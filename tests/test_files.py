import io
import sys
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import httpx
import pytest
from PIL import Image

sys.path.insert(0, str(Path(__file__).parents[1] / "src"))

from files import (
    FileSource,
    PendingExport,
    TeamsFileHandler,
    _pending_exports,
    classify_attachment,
    is_file_return_request,
    requires_user_graph_token,
    should_process_message,
)


def make_handler(transport, max_size=1024, **kwargs):
    credential = Mock()
    credential.get_token.return_value.token = "graph-token"
    client = httpx.Client(transport=httpx.MockTransport(transport))
    return TeamsFileHandler(credential, "contoso", max_size, client=client, **kwargs)


@pytest.mark.parametrize(
    ("attachment", "expected"),
    [
        (
            {"contentType": "application/vnd.microsoft.teams.file.download.info"},
            "file_download_info",
        ),
        ({"contentType": "image/png"}, "image"),
        ({"contentType": "reference"}, "reference"),
        ({"contentType": "hostedContent"}, "hosted_content"),
    ],
)
def test_classifies_teams_file_attachment_shapes(attachment, expected):
    assert classify_attachment(attachment) == expected


@pytest.mark.parametrize(
    ("activity", "expected"),
    [
        ({"conversation": {"conversation_type": "personal"}}, True),
        (
            {
                "conversation": {"conversation_type": "groupChat"},
                "recipient": {"id": "bot"},
                "entities": [{"type": "mention", "mentioned": {"id": "bot"}}],
            },
            True,
        ),
        (
            {
                "conversation": {"conversation_type": "groupChat"},
                "recipient": {"id": "bot"},
                "entities": [],
            },
            False,
        ),
        ({"conversation": {"conversation_type": "channel"}}, False),
    ],
)
def test_process_gate_matches_chat_and_mention_requirements(activity, expected):
    assert should_process_message(activity) is expected


def test_file_return_intent_is_detected():
    assert is_file_return_request("Could you send me that file back?")
    assert not is_file_return_request("What is this file about?")


def test_builds_onedrive_path_from_attached_reference():
    requests = []

    def respond(request):
        requests.append(request)
        return httpx.Response(
            200,
            json={
                "id": "item-id",
                "name": "a.pdf",
                "size": 8,
                "webUrl": "https://contoso-my.sharepoint.com/personal/a/Documents/a.pdf",
                "file": {"mimeType": "application/pdf"},
            },
        )

    handler = make_handler(respond)
    source, metadata = handler._resolve_reference(
        "https://contoso-my.sharepoint.com/personal/alice/Documents/folder/a.pdf",
        "user-id",
        "user-token",
    )

    assert requests[0].url.path == (
        "/v1.0/users/user-id/drive/root:/Documents/folder/a.pdf"
    )
    assert requests[0].headers["Authorization"] == "Bearer user-token"
    assert source.graph_content_path == (
        "/users/user-id/drive/root:/Documents/folder/a.pdf:/content"
    )
    assert metadata["id"] == "item-id"


def test_resolves_localized_sharepoint_library_by_drive_name():
    requested_paths = []

    def respond(request):
        path = request.url.path
        requested_paths.append(path)
        if path.startswith(
            "/v1.0/sites/contoso.sharepoint.com:/sites/research/"
        ):
            return httpx.Response(404)
        if path == "/v1.0/sites/contoso.sharepoint.com:/sites/research":
            return httpx.Response(200, json={"id": "site-id"})
        if path == "/v1.0/sites/site-id/drives":
            return httpx.Response(
                200, json={"value": [{"id": "drive-id", "name": "Library"}]}
            )
        if path == "/v1.0/drives/drive-id/root:/folder/a.pdf":
            return httpx.Response(
                200,
                json={
                    "id": "item-id",
                    "name": "a.pdf",
                    "size": 7,
                    "webUrl": "https://contoso.sharepoint.com/sites/research/Library/folder/a.pdf",
                    "file": {"mimeType": "application/pdf"},
                },
            )
        raise AssertionError(f"Unexpected Graph request path: {path}")

    handler = make_handler(respond)
    source, _ = handler._resolve_reference(
        "https://contoso.sharepoint.com/sites/research/Library/folder/a.pdf",
        "user-id",
        "user-token",
    )

    assert "/v1.0/sites/site-id/drives" in requested_paths
    assert source.graph_content_path == (
        "/drives/drive-id/items/item-id/content"
    )


@pytest.mark.parametrize(
    "url",
    [
        "http://contoso.sharepoint.com/sites/a/file.pdf",
        "https://attacker.sharepoint.com/sites/a/file.pdf",
        "https://contoso.sharepoint.com.attacker.test/sites/a/file.pdf",
        "https://contoso.sharepoint.com:444/sites/a/file.pdf",
    ],
)
def test_rejects_untrusted_sharepoint_hosts(url):
    handler = make_handler(lambda request: httpx.Response(200))
    assert not handler._is_sharepoint_url(url)


def test_ignores_urls_in_message_text_without_attachments():
    requested = []
    handler = make_handler(
        lambda request: requested.append(request) or httpx.Response(200)
    )
    result = handler.process_message_files(
        {
            "text": "Please read https://contoso.sharepoint.com/private.pdf",
            "attachments": [],
        },
        is_group=False,
    )

    assert result.files == []
    assert result.skipped == []
    assert requested == []


def test_downloads_inline_image_with_bot_connector_credential():
    output = io.BytesIO()
    Image.new("RGB", (1, 1), "red").save(output, format="PNG")
    image_bytes = output.getvalue()
    auth_headers = []

    def respond(request):
        auth_headers.append(request.headers.get("Authorization"))
        return httpx.Response(
            200,
            content=image_bytes,
            headers={
                "Content-Type": "image/png",
                "Content-Length": str(len(image_bytes)),
            },
        )

    handler = make_handler(respond)
    activity = {
        "serviceUrl": "https://smba.trafficmanager.net/amer/",
        "attachments": [
            {
                "contentType": "image/png",
                "name": "pasted.png",
                "contentUrl": "https://smba.trafficmanager.net/amer/attachments/1",
            }
        ],
    }
    assert handler.requires_connector_token(activity)

    result = handler.process_message_files(
        activity, is_group=False, connector_token="bot-connector-token"
    )

    assert len(result.files) == 1
    assert auth_headers == ["Bearer bot-connector-token"]


def test_rejects_docx_and_oversized_attachment_before_download():
    requested = []
    handler = make_handler(
        lambda request: requested.append(request) or httpx.Response(200), max_size=4
    )
    result = handler.process_message_files(
        {
            "attachments": [
                {
                    "contentType": "application/vnd.microsoft.teams.file.download.info",
                    "name": "report.docx",
                    "content": {
                        "name": "report.docx",
                        "downloadUrl": "https://contoso-my.sharepoint.com/report.docx",
                        "size": 2,
                    },
                },
                {
                    "contentType": "application/vnd.microsoft.teams.file.download.info",
                    "name": "large.pdf",
                    "content": {
                        "name": "large.pdf",
                        "downloadUrl": "https://contoso-my.sharepoint.com/large.pdf",
                        "size": 5,
                    },
                },
            ]
        },
        is_group=False,
    )

    assert [file.reason for file in result.skipped] == [
        "unsupported type",
        "too large",
    ]
    assert requested == []


def test_rejects_response_larger_than_limit_without_reading_full_body():
    requested = []

    def respond(request):
        requested.append(request)
        return httpx.Response(
            200,
            content=b"%PDF-12345",
            headers={
                "Content-Type": "application/pdf",
                "Content-Length": "10",
            },
        )

    handler = make_handler(respond, max_size=4)
    result = handler.process_message_files(
        {
            "attachments": [
                {
                    "contentType": "application/vnd.microsoft.teams.file.download.info",
                    "name": "large.pdf",
                    "content": {
                        "downloadUrl": "https://contoso-my.sharepoint.com/large.pdf",
                    },
                }
            ]
        },
        is_group=False,
    )

    assert result.files == []
    assert result.skipped[0].reason == "too large"
    assert len(requested) == 1


def test_skipped_error_and_logs_do_not_expose_url(caplog):
    secret_url = "https://evil.example/private.pdf?sig=private-signature"
    handler = make_handler(lambda request: pytest.fail("Untrusted URL was fetched"))
    result = handler.process_message_files(
        {
            "attachments": [
                {
                    "contentType": "application/vnd.microsoft.teams.file.download.info",
                    "name": "private.pdf",
                    "content": {"downloadUrl": secret_url},
                }
            ]
        },
        is_group=False,
    )

    assert result.skipped[0].reason == "unsupported or untrusted URL"
    assert secret_url not in str(result.skipped)
    assert secret_url not in caplog.text


def test_rejects_animated_gif():
    output = io.BytesIO()
    frame1 = Image.new("RGB", (1, 1), "red")
    frame2 = Image.new("RGB", (1, 1), "blue")
    frame1.save(output, format="GIF", save_all=True, append_images=[frame2])
    data = output.getvalue()

    handler = make_handler(
        lambda request: httpx.Response(
            200,
            content=data,
            headers={
                "Content-Type": "image/gif",
                "Content-Length": str(len(data)),
            },
        )
    )
    result = handler.process_message_files(
        {
            "attachments": [
                {
                    "contentType": "image/gif",
                    "name": "moving.gif",
                    "contentUrl": "https://contoso-my.sharepoint.com/moving.gif",
                }
            ]
        },
        is_group=False,
    )

    assert result.files == []
    assert result.skipped[0].reason == "unsupported type (animated GIF)"


def test_file_consent_retains_only_locator_and_returns_file_info():
    body = b"%PDF-example"
    requested = []

    def respond(request):
        requested.append((request.method, request.url.path))
        if request.method == "GET":
            return httpx.Response(
                200,
                content=body,
                headers={
                    "Content-Type": "application/pdf",
                    "Content-Length": str(len(body)),
                },
            )
        return httpx.Response(
            201,
            json={
                "id": "uploaded-id",
                "webUrl": "https://contoso-my.sharepoint.com/shared/copy.pdf",
            },
        )

    handler = make_handler(respond)
    file_result = handler.process_message_files(
        {
            "attachments": [
                {
                    "contentType": "application/vnd.microsoft.teams.file.download.info",
                    "name": "copy.pdf",
                    "content": {
                        "name": "copy.pdf",
                        "downloadUrl": "https://contoso-my.sharepoint.com/source.pdf",
                        "size": len(body),
                    },
                }
            ]
        },
        is_group=False,
    )
    consent = handler.prepare_file_consent_card(
        file_result.files[0], "user-id", "conversation-id"
    )
    export_id = consent["content"]["acceptContext"]["exportId"]

    from files import _pending_exports

    assert not hasattr(_pending_exports[export_id], "content")
    result = handler.handle_file_consent(
        {
            "value": {
                "action": "accept",
                "context": {"exportId": export_id},
                "uploadInfo": {
                    "uploadUrl": "https://contoso-my.sharepoint.com/upload/session"
                },
            },
            "from": {"id": "user-id"},
            "conversation": {"id": "conversation-id"},
        }
    )

    assert result == {
        "status": "uploaded",
        "fileName": "copy.pdf",
        "fileType": "pdf",
        "uniqueId": "uploaded-id",
        "contentUrl": "https://contoso-my.sharepoint.com/shared/copy.pdf",
    }
    assert requested == [
        ("GET", "/source.pdf"),
        ("GET", "/source.pdf"),
        ("PUT", "/upload/session"),
    ]


def test_file_consent_is_bound_to_user_and_conversation():
    handler = make_handler(lambda request: httpx.Response(200))
    card = handler.prepare_file_consent_card(
        SimpleNamespace(
            name="a.pdf",
            content_type="application/pdf",
            content=b"data",
            source=SimpleNamespace(kind="download_url", url="https://contoso-my.sharepoint.com/a.pdf"),
        ),
        "user-id",
        "conversation-id",
    )
    export_id = card["content"]["acceptContext"]["exportId"]

    result = handler.handle_file_consent(
        {
            "value": {
                "action": "accept",
                "context": {"exportId": export_id},
                "uploadInfo": {
                    "uploadUrl": "https://contoso-my.sharepoint.com/upload/session"
                },
            },
            "from": {"id": "another-user"},
            "conversation": {"id": "conversation-id"},
        }
    )

    assert result == {"status": "expired"}


REFERENCE = {
    "contentType": "reference",
    "name": "a.pdf",
    "contentUrl": "https://contoso-my.sharepoint.com/personal/alice/Documents/a.pdf",
}


def test_requires_user_graph_token_only_for_onedrive_or_sharepoint_references():
    assert requires_user_graph_token([REFERENCE])
    assert requires_user_graph_token([{"contentType": "image/png"}, REFERENCE])
    assert not requires_user_graph_token([])
    assert not requires_user_graph_token(
        [
            {"contentType": "image/png"},
            {"contentType": "hostedContent"},
            {"contentType": "application/vnd.microsoft.teams.file.download.info"},
        ]
    )


@pytest.mark.parametrize(
    ("sign_in_enabled", "reason"),
    [
        (True, "sign-in required to read OneDrive or SharePoint files"),
        (False, "Teams sign-in isn't configured in this environment"),
    ],
)
def test_reference_without_user_token_is_skipped_without_app_fallback(
    sign_in_enabled, reason
):
    handler = make_handler(
        lambda request: pytest.fail("Graph must not be called without a user token"),
        user_sign_in_enabled=sign_in_enabled,
    )

    result = handler.process_message_files({"attachments": [REFERENCE]}, is_group=False)

    assert result.files == []
    assert result.skipped[0].reason.startswith(reason)
    handler._credential.get_token.assert_not_called()


def test_group_reference_uses_user_token_while_inline_image_uses_rsc_app_token():
    body = b"%PDF-example"
    image = io.BytesIO()
    Image.new("RGB", (1, 1), "red").save(image, format="PNG")
    seen = []

    def respond(request):
        path = request.url.path
        seen.append((path, request.headers.get("Authorization")))
        if path == "/v1.0/users/user-id/drive/root:/Documents/a.pdf":
            return httpx.Response(
                200,
                json={
                    "id": "item-id",
                    "name": "a.pdf",
                    "size": len(body),
                    "file": {"mimeType": "application/pdf"},
                },
            )
        if path == "/v1.0/users/user-id/drive/root:/Documents/a.pdf:/content":
            return httpx.Response(
                200, content=body, headers={"Content-Type": "application/pdf"}
            )
        if path.endswith("/hostedContents/hosted-id/$value"):
            return httpx.Response(
                200, content=image.getvalue(), headers={"Content-Type": "image/png"}
            )
        raise AssertionError(f"Unexpected Graph request path: {path}")

    reader = Mock()
    handler = make_handler(respond, graph_message_reader=reader)
    group_message = {
        "from": {"user": {"id": "user-id"}},
        "attachments": [REFERENCE],
        "hostedContents": [{"id": "hosted-id", "contentType": "image/png"}],
    }

    result = handler.process_message_files(
        {"id": "message-id", "conversation": {"id": "19:chat"}, "attachments": []},
        is_group=True,
        user_graph_token="user-token",
        group_message=group_message,
    )

    reader.assert_not_called()
    assert [file.name for file in result.files] == ["a.pdf", "hosted-id.png"]
    auth_by_path = dict(seen)
    assert auth_by_path["/v1.0/users/user-id/drive/root:/Documents/a.pdf"] == (
        "Bearer user-token"
    )
    assert auth_by_path[
        "/v1.0/users/user-id/drive/root:/Documents/a.pdf:/content"
    ] == ("Bearer user-token")
    hosted_path = next(path for path in auth_by_path if "hostedContents" in path)
    assert auth_by_path[hosted_path] == "Bearer graph-token"


def _pending_graph_export(export_id):
    _pending_exports[export_id] = PendingExport(
        source=FileSource(
            "graph_item",
            graph_content_path="/users/user-id/drive/root:/Documents/a.pdf:/content",
        ),
        file_name="a.pdf",
        content_type="application/pdf",
        size=12,
        user_id="user-id",
        conversation_id="conversation-id",
        expires_at=float("inf"),
    )


def _consent_activity(export_id, action="accept"):
    return {
        "value": {
            "action": action,
            "context": {"exportId": export_id},
            "uploadInfo": {
                "uploadUrl": "https://contoso-my.sharepoint.com/upload/session"
            },
        },
        "from": {"id": "user-id"},
        "conversation": {"id": "conversation-id"},
    }


def test_file_consent_for_onedrive_item_needs_and_uses_user_token():
    body = b"%PDF-example"
    seen = []

    def respond(request):
        seen.append((request.method, request.headers.get("Authorization")))
        if request.method == "GET":
            return httpx.Response(
                200, content=body, headers={"Content-Type": "application/pdf"}
            )
        return httpx.Response(201, json={"id": "uploaded-id", "webUrl": ""})

    handler = make_handler(respond)
    _pending_graph_export("graph-export")
    assert handler.requires_consent_user_token(_consent_activity("graph-export"))
    assert not handler.requires_consent_user_token(
        _consent_activity("graph-export", "decline")
    )

    result = handler.handle_file_consent(
        _consent_activity("graph-export"), user_graph_token="user-token"
    )

    assert result["status"] == "uploaded"
    assert seen[0] == ("GET", "Bearer user-token")
    handler._credential.get_token.assert_not_called()


def test_file_consent_for_onedrive_item_without_user_token_fails_closed():
    handler = make_handler(
        lambda request: pytest.fail("Graph must not be called without a user token")
    )
    _pending_graph_export("graph-export-no-token")

    result = handler.handle_file_consent(_consent_activity("graph-export-no-token"))

    assert result == {"status": "upload_failed"}
