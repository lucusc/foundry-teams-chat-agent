import json
import io
import logging
import re
import secrets
import time
from dataclasses import dataclass
from typing import Any
from urllib.parse import quote, unquote, urljoin, urlparse

import httpx
from PIL import Image, UnidentifiedImageError

from foundry import is_bot_mentioned, is_group_conversation

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPE = "https://graph.microsoft.com/.default"
MAX_REDIRECTS = 3
UPLOAD_CHUNK_SIZE = 320 * 1024
SUPPORTED_TYPES = {
    ".pdf": ("application/pdf", "file"),
    ".png": ("image/png", "image"),
    ".jpg": ("image/jpeg", "image"),
    ".jpeg": ("image/jpeg", "image"),
    ".webp": ("image/webp", "image"),
    ".gif": ("image/gif", "image"),
}
MIME_TO_EXTENSION = {
    content_type: extension
    for extension, (content_type, _) in SUPPORTED_TYPES.items()
}

logger = logging.getLogger("demo-foundry-agent.files")


class FileReadError(Exception):
    def __init__(self, reason: str):
        super().__init__(reason)
        self.reason = reason


@dataclass(frozen=True)
class FileSource:
    kind: str
    url: str | None = None
    allowed_host: str | None = None
    graph_content_path: str | None = None
    chat_id: str | None = None
    message_id: str | None = None
    hosted_content_id: str | None = None


@dataclass(frozen=True)
class FileArtifact:
    name: str
    content_type: str
    content: bytes
    web_url: str | None
    source: FileSource


@dataclass(frozen=True)
class SkippedFile:
    name: str
    reason: str


@dataclass
class FileProcessingResult:
    files: list[FileArtifact]
    skipped: list[SkippedFile]


@dataclass(frozen=True)
class PendingExport:
    file_name: str
    content_type: str
    size: int
    source: FileSource
    user_id: str
    conversation_id: str
    expires_at: float


_pending_exports: dict[str, PendingExport] = {}


def _field(value: Any, name: str, default=None):
    snake_name = re.sub(r"(?<!^)(?=[A-Z])", "_", name).lower()
    if isinstance(value, dict):
        return value.get(name, value.get(snake_name, default))
    return getattr(value, name, getattr(value, snake_name, default))


def _attachment_content(attachment: Any) -> dict[str, Any]:
    content = _field(attachment, "content")
    if isinstance(content, str):
        try:
            parsed = json.loads(content)
        except json.JSONDecodeError:
            return {}
        return parsed if isinstance(parsed, dict) else {}
    return content if isinstance(content, dict) else {}


def classify_attachment(attachment: Any) -> str | None:
    """Classify Teams file attachments without inspecting the message text."""
    content_type = str(_field(attachment, "contentType", "") or "").lower()
    if content_type == "application/vnd.microsoft.teams.file.download.info":
        return "file_download_info"
    if content_type == "reference":
        return "reference"
    if content_type == "hostedcontent":
        return "hosted_content"
    if content_type.startswith("image/"):
        return "image"
    return None


def should_process_message(activity: Any) -> bool:
    conversation_type = _field(_field(activity, "conversation"), "conversation_type")
    is_group = conversation_type == "groupChat" or is_group_conversation(activity)
    return conversation_type != "channel" and (
        not is_group or is_bot_mentioned(activity)
    )


def is_file_return_request(text: str) -> bool:
    return bool(
        re.search(
            r"\b(send|return|give|attach)\b.{0,80}\b(file|document|copy)\b",
            text,
            flags=re.IGNORECASE | re.DOTALL,
        )
    )


def requires_user_graph_token(attachments: list[Any]) -> bool:
    """Return True when an attachment can only be read with delegated Graph access.

    Teams represents OneDrive/SharePoint files shared in a chat as ``reference``
    attachments. They have no chat-scoped grant: no resource-specific consent
    (RSC) permission covers files. Reading them with an application permission
    would need tenant-wide Files.Read.All/Sites.Read.All, so they are read only
    on behalf of the signed-in user. See docs/file-access-security.md.
    """
    return any(
        classify_attachment(attachment) == "reference" for attachment in attachments
    )


class TeamsFileHandler:
    def __init__(
        self,
        graph_credential,
        sharepoint_tenant_domain: str,
        max_file_size_bytes: int,
        graph_message_reader=None,
        client: httpx.Client | None = None,
        user_sign_in_enabled: bool = True,
    ):
        # graph_credential is an *application* credential used only for
        # ChatMessage.Read.Chat RSC calls (the triggering message and its inline
        # hostedContents). Those calls are authorized per chat by RSC consent. The
        # application holds no tenant-wide Files/Sites permission, so it can't
        # open OneDrive or SharePoint items.
        self._credential = graph_credential
        self._tenant_domain = sharepoint_tenant_domain.strip().lower()
        self._max_file_size = max_file_size_bytes
        self._graph_message_reader = graph_message_reader
        self._client = client or httpx.Client(timeout=30.0)
        self._missing_user_token_reason = (
            "sign-in required to read OneDrive or SharePoint files"
            if user_sign_in_enabled
            else "Teams sign-in isn't configured in this environment, so "
            "OneDrive or SharePoint files can't be read"
        )

    def requires_connector_token(self, activity: Any) -> bool:
        service_url = _field(activity, "service_url") or _field(
            activity, "serviceUrl"
        )
        for attachment in _field(activity, "attachments", []) or []:
            if classify_attachment(attachment) != "image":
                continue
            content_url = _field(attachment, "contentUrl")
            if content_url and self._is_bot_connector_url(content_url, service_url):
                return True
        return False

    def process_message_files(
        self,
        activity: Any,
        is_group: bool,
        connector_token: str | None = None,
        user_graph_token: str | None = None,
        group_message: dict[str, Any] | None = None,
    ) -> FileProcessingResult:
        """Read only the files attached to the triggering message.

        ``user_graph_token`` is a delegated Microsoft Graph token for the user who
        sent the message. The Azure Bot token service obtains it through Teams SSO
        and the Microsoft Entra on-behalf-of (OBO) flow. Graph enforces that user's
        own file permissions, so the agent can't read anything that user can't
        already open. Only references in this message are resolved.
        ``group_message`` is an optional pre-read copy of the triggering
        group-chat message, so it isn't fetched twice.
        """
        attachments = list(_field(activity, "attachments", []) or [])
        chat_id = None
        message_id = _field(activity, "id")
        sender_id = _field(
            _field(activity, "from_property") or _field(activity, "from"), "id"
        )

        if is_group:
            chat_id = _field(_field(activity, "conversation"), "id")
            if not chat_id or not message_id:
                raise RuntimeError("Group-chat activity is missing its message identity")
            if group_message is not None:
                message = group_message
            elif self._graph_message_reader:
                message = self._graph_message_reader(chat_id, message_id)
            else:
                # A4: the app in webApplicationInfo.id (deployed: the UserAuth app
                # via its managed-identity federated credential; local: the bot
                # app's client secret) gets an app-only Graph token.
                # ChatMessage.Read.Chat RSC authorizes reading only this chat.
                message = self._graph_json(
                    f"/chats/{quote(chat_id, safe='')}/messages/{quote(message_id, safe='')}",
                    permission_failure="consent missing",
                )
            attachments.extend(message.get("attachments") or [])
            sender = message.get("from") or {}
            sender_id = (
                (sender.get("user") or sender.get("application") or {}).get("id")
                or sender_id
            )
            for hosted in message.get("hostedContents") or []:
                attachments.append(
                    {
                        "contentType": "hostedContent",
                        "id": hosted.get("id"),
                        "name": hosted.get("id") or "inline image",
                        "mediaType": hosted.get("contentType"),
                    }
                )

        service_url = _field(activity, "service_url") or _field(
            activity, "serviceUrl"
        )
        result = FileProcessingResult(files=[], skipped=[])
        seen: set[str] = set()
        for attachment in attachments:
            kind = classify_attachment(attachment)
            if kind is None:
                continue

            content = _attachment_content(attachment)
            name = self._safe_name(
                _field(attachment, "name")
                or content.get("name")
                or content.get("fileName")
                or "shared file"
            )
            identity = self._attachment_identity(attachment, content, kind)
            if identity in seen:
                continue
            seen.add(identity)

            try:
                artifact = self._read_attachment(
                    attachment,
                    content,
                    kind,
                    name,
                    sender_id,
                    chat_id,
                    message_id,
                    service_url,
                    connector_token,
                    user_graph_token,
                )
                result.files.append(artifact)
                logger.info(
                    "Read attached file name=%r type=%s size=%d outcome=success",
                    artifact.name,
                    artifact.content_type,
                    len(artifact.content),
                )
            except FileReadError as error:
                logger.info(
                    "Skipped attached file name=%r type=%s size=%s outcome=%s",
                    name,
                    self._declared_type(attachment, content),
                    content.get("size"),
                    error.reason,
                )
                result.skipped.append(SkippedFile(name, error.reason))
            except httpx.HTTPError:
                logger.info(
                    "Skipped attached file name=%r outcome=permission not granted",
                    name,
                )
                result.skipped.append(
                    SkippedFile(name, "permission not granted or consent missing")
                )
        return result

    def prepare_file_consent_card(
        self, artifact: FileArtifact, user_id: str, conversation_id: str
    ) -> dict[str, Any]:
        self._remove_expired_exports()
        if len(_pending_exports) >= 100:
            oldest = min(
                _pending_exports,
                key=lambda export_id: _pending_exports[export_id].expires_at,
            )
            del _pending_exports[oldest]
        export_id = secrets.token_urlsafe(24)
        _pending_exports[export_id] = PendingExport(
            file_name=artifact.name,
            content_type=artifact.content_type,
            size=len(artifact.content),
            source=artifact.source,
            user_id=user_id,
            conversation_id=conversation_id,
            expires_at=time.monotonic() + 900,
        )
        return {
            "contentType": "application/vnd.microsoft.teams.card.file.consent",
            "content": {
                "description": "Send a copy of this attached file to the chat?",
                "sizeInBytes": len(artifact.content),
                "acceptContext": {"exportId": export_id, "action": "accept"},
                "declineContext": {"exportId": export_id, "action": "decline"},
            },
        }

    def requires_consent_connector_token(self, activity: Any) -> bool:
        value = _field(activity, "value", {})
        context = _field(value, "context", {})
        export_id = _field(context, "exportId")
        pending = _pending_exports.get(export_id) if export_id else None
        user_id = _field(
            _field(activity, "from_property") or _field(activity, "from"), "id"
        )
        conversation_id = _field(_field(activity, "conversation"), "id")
        return bool(
            pending
            and pending.expires_at >= time.monotonic()
            and pending.user_id == user_id
            and pending.conversation_id == conversation_id
            and pending.source.kind == "connector_url"
        )

    def requires_consent_user_token(self, activity: Any) -> bool:
        """Return True when a file-consent export must re-read a OneDrive or SharePoint item.

        Those items are re-read with the accepting user's delegated token, never
        with an application permission.
        """
        value = _field(activity, "value", {})
        context = _field(value, "context", {})
        export_id = _field(context, "exportId")
        pending = _pending_exports.get(export_id) if export_id else None
        return bool(
            str(_field(value, "action", "")).lower() == "accept"
            and pending
            and pending.source.kind == "graph_item"
        )

    def handle_file_consent(
        self,
        activity: Any,
        connector_token: str | None = None,
        user_graph_token: str | None = None,
    ) -> dict[str, Any] | None:
        value = _field(activity, "value", {})
        action = str(_field(value, "action", "")).lower()
        context = _field(value, "context", {})
        export_id = _field(context, "exportId")
        pending = _pending_exports.get(export_id) if export_id else None
        user_id = _field(
            _field(activity, "from_property") or _field(activity, "from"), "id"
        )
        conversation_id = _field(_field(activity, "conversation"), "id")
        if (
            not pending
            or pending.expires_at < time.monotonic()
            or pending.user_id != user_id
            or pending.conversation_id != conversation_id
        ):
            if export_id:
                _pending_exports.pop(export_id, None)
            return {"status": "expired"}

        if action == "decline":
            _pending_exports.pop(export_id, None)
            return {"status": "declined"}
        if action != "accept":
            return {"status": "invalid"}

        upload_info = _field(value, "uploadInfo", {})
        upload_url = _field(upload_info, "uploadUrl")
        if not upload_url or not self._is_sharepoint_url(upload_url):
            _pending_exports.pop(export_id, None)
            return {"status": "upload_failed"}

        try:
            # Content is fetched again after consent; only its safe locator is
            # retained between turns, never the file bytes.
            content = self._read_source(
                pending.source,
                pending.file_name,
                pending.content_type,
                connector_token,
                _field(activity, "service_url") or _field(activity, "serviceUrl"),
                user_graph_token,
            )
            uploaded = self._upload_file(upload_url, content)
        except (FileReadError, httpx.HTTPError):
            _pending_exports.pop(export_id, None)
            return {"status": "upload_failed"}

        _pending_exports.pop(export_id, None)
        return {
            "status": "uploaded",
            "fileName": pending.file_name,
            "fileType": pending.content_type.split("/", 1)[-1],
            "uniqueId": str(uploaded.get("id") or ""),
            "contentUrl": str(uploaded.get("webUrl") or ""),
        }

    def _read_attachment(
        self,
        attachment: Any,
        content: dict[str, Any],
        kind: str,
        name: str,
        sender_id: str | None,
        chat_id: str | None,
        message_id: str | None,
        service_url: str | None,
        connector_token: str | None,
        user_graph_token: str | None = None,
    ) -> FileArtifact:
        declared_type = self._declared_type(attachment, content)
        file_name, content_type = self._resolve_file_type(
            name,
            content.get("fileType") or content.get("mediaType"),
            declared_type,
        )
        known_size = self._size(content.get("size") or _field(attachment, "size"))
        if known_size is not None:
            self._check_size(known_size)

        web_url = content.get("webUrl") or _field(attachment, "contentUrl")
        if kind == "hosted_content":
            if not chat_id or not message_id:
                raise FileReadError("unsupported attachment")
            hosted_id = _field(attachment, "id")
            if not hosted_id:
                raise FileReadError("unsupported attachment")
            path = (
                f"/chats/{quote(chat_id, safe='')}/messages/"
                f"{quote(message_id, safe='')}/hostedContents/"
                f"{quote(str(hosted_id), safe='')}/$value"
            )
            source = FileSource(
                "hosted_content",
                graph_content_path=path,
                chat_id=chat_id,
                message_id=message_id,
                hosted_content_id=str(hosted_id),
            )
            data = self._read_graph_content(path, content_type)
        elif kind == "reference":
            # A5: OneDrive/SharePoint references are resolved and downloaded
            # only with the sender's delegated Graph token (Teams SSO + OBO).
            # There's deliberately no application-permission fallback, because
            # that would need tenant-wide Files.Read.All/Sites.Read.All.
            if not user_graph_token:
                raise FileReadError(self._missing_user_token_reason)
            reference_url = _field(attachment, "contentUrl") or content.get("contentUrl")
            source, metadata = self._resolve_reference(
                reference_url, sender_id, user_graph_token
            )
            file_name = self._safe_name(metadata.get("name") or file_name)
            web_url = metadata.get("webUrl") or web_url
            known_size = self._size(metadata.get("size"))
            if known_size is not None:
                self._check_size(known_size)
            file_name, content_type = self._resolve_file_type(
                file_name,
                None,
                (metadata.get("file") or {}).get("mimeType")
                or declared_type,
            )
            data = self._read_graph_content(
                source.graph_content_path, content_type, user_graph_token
            )
        else:
            download_url = (
                content.get("downloadUrl")
                or _field(attachment, "contentUrl")
                or content.get("contentUrl")
            )
            if not download_url:
                raise FileReadError("unsupported attachment")
            if self._is_bot_connector_url(download_url, service_url):
                if not connector_token:
                    raise FileReadError("permission not granted")
                # A2: the Agents SDK connection manager supplies the bot UAMI
                # (deployed) or local bot app (debugging) token for
                # https://api.botframework.com; Azure Bot registration authorizes it.
                source = FileSource(
                    "connector_url",
                    url=download_url,
                    allowed_host=(urlparse(service_url).hostname or "").lower(),
                )
                headers = {"Authorization": f"Bearer {connector_token}"}
                connector_hosts = {source.allowed_host}
            else:
                if not self._is_sharepoint_url(download_url):
                    raise FileReadError("unsupported or untrusted URL")
                # A3: Teams issues a short-lived pre-authenticated SharePoint
                # download URL in local and deployed chats; no app identity or
                # Graph/Bot token is sent to this endpoint.
                source = FileSource("download_url", url=download_url)
                headers = {}
                connector_hosts = set()
            data = self._read_url(
                download_url,
                headers,
                content_type,
                additional_allowed_hosts=connector_hosts,
            )

        self._check_size(len(data))
        self._validate_gif(data, content_type)
        return FileArtifact(
            name=file_name,
            content_type=content_type,
            content=data,
            web_url=web_url if self._is_sharepoint_url(web_url) else None,
            source=source,
        )

    def _read_source(
        self,
        source: FileSource,
        file_name: str,
        content_type: str,
        connector_token: str | None,
        service_url: str | None,
        user_graph_token: str | None = None,
    ) -> bytes:
        if source.kind == "download_url" and source.url:
            if not self._is_sharepoint_url(source.url):
                raise FileReadError("unsupported or untrusted URL")
            return self._read_url(source.url, {}, content_type)
        if source.kind == "connector_url" and source.url:
            if (
                not connector_token
                or not self._is_bot_connector_url(source.url, service_url)
                or (urlparse(source.url).hostname or "").lower()
                != source.allowed_host
            ):
                raise FileReadError("permission not granted")
            return self._read_url(
                source.url,
                {"Authorization": f"Bearer {connector_token}"},
                content_type,
                additional_allowed_hosts={source.allowed_host or ""},
            )
        if source.kind == "graph_item":
            if not source.graph_content_path:
                raise FileReadError("unsupported attachment")
            if not user_graph_token:
                raise FileReadError(self._missing_user_token_reason)
            return self._read_graph_content(
                source.graph_content_path, content_type, user_graph_token
            )
        if source.kind == "hosted_content":
            if not source.graph_content_path:
                raise FileReadError("unsupported attachment")
            return self._read_graph_content(
                source.graph_content_path, content_type
            )
        raise FileReadError("unsupported attachment")

    def _resolve_reference(
        self,
        reference_url: str | None,
        sender_id: str | None,
        user_graph_token: str,
    ) -> tuple[FileSource, dict[str, Any]]:
        if not reference_url or not self._is_sharepoint_url(reference_url):
            raise FileReadError("unsupported or untrusted URL")
        parsed = urlparse(reference_url)
        segments = [unquote(part) for part in parsed.path.split("/") if part]
        if any(part in (".", "..") for part in segments):
            raise FileReadError("unsupported link format")

        host = (parsed.hostname or "").lower()
        personal_host = f"{self._tenant_domain}-my.sharepoint.com"
        if host == personal_host:
            if not sender_id or len(segments) < 4 or segments[0].lower() != "personal":
                raise FileReadError("permission not granted")
            if segments[2].casefold() not in ("documents", "shared documents"):
                raise FileReadError("unsupported link format")
            path = "/" + "/".join(quote(part, safe="") for part in segments[2:])
            user_path = f"/users/{quote(sender_id, safe='')}/drive/root:{path}"
            metadata = self._graph_json(
                user_path,
                permission_failure="permission not granted",
                token=user_graph_token,
            )
            content_path = user_path + ":/content"
            return FileSource("graph_item", graph_content_path=content_path), metadata

        if not segments or segments[0].lower() not in ("sites", "teams"):
            raise FileReadError("unsupported link format")

        # SharePoint document library names can be localized. Try the candidate
        # site paths from longest to shortest and match the path's library
        # segment against Graph's actual drive names rather than assuming a name.
        for site_segment_count in range(len(segments) - 1, 1, -1):
            site_path = "/" + "/".join(
                quote(part, safe="") for part in segments[:site_segment_count]
            )
            site = self._graph_json(
                f"/sites/{quote(host, safe='')}:{site_path}",
                allow_not_found=True,
                permission_failure="permission not granted",
                token=user_graph_token,
            )
            if not site:
                continue
            drives_payload = self._graph_json(
                f"/sites/{quote(str(site.get('id', '')), safe='')}/drives",
                permission_failure="permission not granted",
                token=user_graph_token,
            )
            drives = drives_payload.get("value", [])
            library_name = segments[site_segment_count]
            drive = next(
                (
                    candidate
                    for candidate in drives
                    if str(candidate.get("name", "")).casefold()
                    == library_name.casefold()
                ),
                None,
            )
            if not drive:
                continue
            relative_path = segments[site_segment_count + 1 :]
            if not relative_path:
                raise FileReadError("unsupported link format")
            encoded_path = "/".join(quote(part, safe="") for part in relative_path)
            drive_id = quote(str(drive.get("id", "")), safe="")
            item_path = f"/drives/{drive_id}/root:/{encoded_path}"
            metadata = self._graph_json(
                item_path,
                permission_failure="permission not granted",
                token=user_graph_token,
            )
            item_id = quote(str(metadata.get("id", "")), safe="")
            if not item_id:
                raise FileReadError("unsupported link format")
            return (
                FileSource(
                    "graph_item",
                    graph_content_path=f"/drives/{drive_id}/items/{item_id}/content",
                ),
                metadata,
            )
        raise FileReadError("permission not granted or unsupported link format")

    def _graph_json(
        self,
        path: str,
        permission_failure: str = "permission not granted",
        allow_not_found: bool = False,
        token: str | None = None,
    ) -> dict[str, Any]:
        # A5: `token` is the signed-in user's delegated Graph token (Teams SSO +
        # OBO) for OneDrive/SharePoint lookups. Graph then applies that user's own
        # access, through the user-consentable delegated Files.Read.All and
        # Sites.Read.All scopes. Without `token`, the call uses the app credential,
        # which only has chat-scoped ChatMessage.Read.Chat RSC. No tenant-wide
        # application file permission is ever granted.
        token = token or self._credential.get_token(GRAPH_SCOPE).token
        try:
            response = self._client.get(
                f"{GRAPH_BASE_URL}{path}",
                headers={"Authorization": f"Bearer {token}"},
                follow_redirects=False,
            )
        except httpx.HTTPError:
            raise
        if response.status_code == 404 and allow_not_found:
            return {}
        if response.status_code in (401, 403):
            raise FileReadError(permission_failure)
        if not 200 <= response.status_code < 300:
            raise FileReadError("file could not be read")
        payload = response.json()
        if not isinstance(payload, dict):
            raise FileReadError("file could not be read")
        return payload

    def _read_graph_content(
        self, path: str | None, content_type: str, token: str | None = None
    ) -> bytes:
        if not path or not path.startswith("/"):
            raise FileReadError("unsupported attachment")
        # Delegated token for drive items; app-only RSC token for hostedContents.
        token = token or self._credential.get_token(GRAPH_SCOPE).token
        return self._read_url(
            f"{GRAPH_BASE_URL}{path}",
            {"Authorization": f"Bearer {token}"},
            content_type,
            graph_endpoint=True,
        )

    def _read_url(
        self,
        url: str,
        headers: dict[str, str],
        content_type: str,
        graph_endpoint: bool = False,
        additional_allowed_hosts: set[str] | None = None,
    ) -> bytes:
        current_url = url
        allowed_hosts = self._allowed_hosts() | (additional_allowed_hosts or set())
        allowed_hosts.discard("")
        if graph_endpoint:
            allowed_hosts.add("graph.microsoft.com")

        for _ in range(MAX_REDIRECTS + 1):
            parsed = urlparse(current_url)
            if parsed.scheme != "https" or parsed.hostname not in allowed_hosts:
                raise FileReadError("unsupported or untrusted URL")
            try:
                with self._client.stream(
                    "GET",
                    current_url,
                    headers=headers,
                    follow_redirects=False,
                ) as response:
                    if response.is_redirect:
                        location = response.headers.get("Location")
                        if not location:
                            raise FileReadError("file could not be read")
                        next_url = urljoin(current_url, location)
                        next_host = (urlparse(next_url).hostname or "").lower()
                        if next_host != (parsed.hostname or "").lower():
                            headers = {}
                        current_url = next_url
                        continue
                    if response.status_code in (401, 403):
                        raise FileReadError("permission not granted")
                    if response.status_code == 404:
                        raise FileReadError("consent missing or file unavailable")
                    if not 200 <= response.status_code < 300:
                        raise FileReadError("file could not be read")
                    response_type = response.headers.get("Content-Type", "").split(
                        ";", 1
                    )[0].lower()
                    if response_type and response_type not in (
                        content_type,
                        "application/octet-stream",
                    ):
                        raise FileReadError("unsupported type")
                    content_length = self._size(
                        response.headers.get("Content-Length")
                    )
                    if content_length is not None:
                        self._check_size(content_length)
                    data = bytearray()
                    for chunk in response.iter_bytes():
                        data.extend(chunk)
                        self._check_size(len(data))
                    return bytes(data)
            except httpx.HTTPError:
                raise
        raise FileReadError("too many redirects")

    def _upload_file(self, upload_url: str, content: bytes) -> dict[str, Any]:
        if not self._is_sharepoint_url(upload_url):
            raise FileReadError("unsupported or untrusted URL")
        final_response = None
        for start in range(0, len(content), UPLOAD_CHUNK_SIZE):
            end = min(start + UPLOAD_CHUNK_SIZE, len(content)) - 1
            chunk = content[start : end + 1]
            # A6: the accepted FileConsentCard provides a short-lived,
            # pre-authenticated SharePoint upload URL in local and deployed chats.
            # No Graph or bot token is sent; the user's consent authorizes upload.
            response = self._client.put(
                upload_url,
                content=chunk,
                headers={
                    "Content-Length": str(len(chunk)),
                    "Content-Range": f"bytes {start}-{end}/{len(content)}",
                },
                follow_redirects=False,
            )
            if response.is_redirect:
                location = response.headers.get("Location")
                redirected = urljoin(upload_url, location or "")
                if not self._is_sharepoint_url(redirected):
                    raise FileReadError("upload failed")
                response = self._client.put(
                    redirected,
                    content=chunk,
                    headers={
                        "Content-Length": str(len(chunk)),
                        "Content-Range": f"bytes {start}-{end}/{len(content)}",
                    },
                    follow_redirects=False,
                )
            if not 200 <= response.status_code < 300:
                raise FileReadError("upload failed")
            final_response = response
        if final_response is None:
            raise FileReadError("upload failed")
        try:
            payload = final_response.json()
        except (ValueError, json.JSONDecodeError):
            return {}
        return payload if isinstance(payload, dict) else {}

    def _resolve_file_type(
        self, name: str, file_type: Any, content_type: Any
    ) -> tuple[str, str]:
        declared = str(content_type or "").split(";", 1)[0].strip().lower()
        extension = self._extension(name)
        if not extension and file_type:
            suffix = str(file_type).strip().lower()
            extension = suffix if suffix.startswith(".") else f".{suffix}"
        if not extension and declared in MIME_TO_EXTENSION:
            extension = MIME_TO_EXTENSION[declared]
        supported = SUPPORTED_TYPES.get(extension)
        if not supported:
            raise FileReadError("unsupported type")
        expected_content_type, _kind = supported
        if declared and declared not in (
            expected_content_type,
            "application/octet-stream",
            "reference",
            "application/vnd.microsoft.teams.file.download.info",
        ):
            raise FileReadError("unsupported type")
        if not self._extension(name):
            name = f"{name}{extension}"
        return self._safe_name(name), expected_content_type

    @staticmethod
    def _extension(name: str) -> str:
        match = re.search(r"(\.[a-z0-9]+)$", name.lower())
        return match.group(1) if match else ""

    @staticmethod
    def _declared_type(attachment: Any, content: dict[str, Any]) -> str:
        return str(
            _field(attachment, "mediaType")
            or content.get("mediaType")
            or content.get("mimeType")
            or content.get("contentType")
            or _field(attachment, "contentType")
            or ""
        )

    @staticmethod
    def _safe_name(value: Any) -> str:
        name = str(value or "shared file").replace("\\", "/").split("/")[-1]
        name = re.sub(r"[\x00-\x1f\x7f]", " ", name)
        return name[:255].strip() or "shared file"

    @staticmethod
    def _size(value: Any) -> int | None:
        try:
            return int(value) if value is not None else None
        except (TypeError, ValueError):
            return None

    def _check_size(self, size: int):
        if size < 0:
            raise FileReadError("invalid file size")
        if size > self._max_file_size:
            raise FileReadError("too large")

    @staticmethod
    def _validate_gif(content: bytes, content_type: str):
        if content_type != "image/gif":
            return
        try:
            with Image.open(io.BytesIO(content)) as image:
                if getattr(image, "n_frames", 1) > 1:
                    raise FileReadError("unsupported type (animated GIF)")
        except (UnidentifiedImageError, OSError):
            raise FileReadError("unsupported type") from None

    def _is_sharepoint_url(self, value: Any) -> bool:
        if not isinstance(value, str) or not self._tenant_domain:
            return False
        parsed = urlparse(value)
        allowed = {
            f"{self._tenant_domain}.sharepoint.com",
            f"{self._tenant_domain}-my.sharepoint.com",
        }
        try:
            port = parsed.port
        except ValueError:
            return False
        return (
            parsed.scheme == "https"
            and parsed.hostname is not None
            and parsed.hostname.lower() in allowed
            and port in (None, 443)
            and not parsed.username
            and not parsed.password
        )

    @staticmethod
    def _is_bot_connector_url(value: Any, service_url: str | None) -> bool:
        if not isinstance(value, str) or not isinstance(service_url, str):
            return False
        target = urlparse(value)
        service = urlparse(service_url)
        try:
            target_port = target.port
        except ValueError:
            return False
        return (
            target.scheme == "https"
            and service.scheme == "https"
            and target.hostname is not None
            and target.hostname.lower() == (service.hostname or "").lower()
            and target_port in (None, 443)
        )

    def _allowed_hosts(self) -> set[str]:
        if not self._tenant_domain:
            return set()
        return {
            f"{self._tenant_domain}.sharepoint.com",
            f"{self._tenant_domain}-my.sharepoint.com",
        }

    def _attachment_identity(
        self, attachment: Any, content: dict[str, Any], kind: str
    ) -> str:
        return "|".join(
            (
                kind,
                str(content.get("downloadUrl") or ""),
                str(_field(attachment, "contentUrl") or ""),
                str(_field(attachment, "id") or ""),
                str(_field(attachment, "name") or ""),
            )
        )

    def _remove_expired_exports(self):
        now = time.monotonic()
        for export_id, pending in list(_pending_exports.items()):
            if pending.expires_at < now:
                del _pending_exports[export_id]
