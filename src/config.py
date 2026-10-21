import re
from urllib.parse import urlparse

USER_AUTH_HANDLER = "GRAPH"

class Config:
    """Runtime configuration for the Teams-hosted Foundry agent."""

    def __init__(self, env):
        self.project_endpoint = self._required(env, "AZURE_AI_PROJECT_ENDPOINT")
        self._validate_project_endpoint(self.project_endpoint)
        self.model_deployment_name = self._required(
            env, "AZURE_AI_MODEL_DEPLOYMENT_NAME"
        )
        self.managed_identity_client_id = env.get("AZURE_CLIENT_ID", "").strip() or None
        self.sharepoint_tenant_domain = env.get(
            "SHAREPOINT_TENANT_DOMAIN", ""
        ).strip().lower()
        if self.sharepoint_tenant_domain and not re.fullmatch(
            r"[a-z0-9](?:[a-z0-9-]*[a-z0-9])?", self.sharepoint_tenant_domain
        ):
            raise RuntimeError(
                "SHAREPOINT_TENANT_DOMAIN must be the SharePoint tenant prefix"
            )

        raw_file_size = env.get("MAX_FILE_SIZE_BYTES", "20971520")
        try:
            self.max_file_size_bytes = int(raw_file_size)
        except ValueError as error:
            raise RuntimeError("MAX_FILE_SIZE_BYTES must be an integer") from error
        if self.max_file_size_bytes < 1:
            raise RuntimeError("MAX_FILE_SIZE_BYTES must be greater than zero")

        # GRAPH_AUTH_MODE selects the *application* identity used only for
        # ChatMessage.Read.Chat RSC reads. That identity holds no tenant-wide
        # Graph permission. File contents are read with the user's delegated token.
        self.graph_auth_mode = env.get("GRAPH_AUTH_MODE", "default").strip()
        if self.graph_auth_mode not in ("default", "bot", "federated"):
            raise RuntimeError(
                "GRAPH_AUTH_MODE must be 'default', 'bot', or 'federated'"
            )
        self.graph_federated_app: tuple[str, str] | None = None
        if self.graph_auth_mode == "federated":
            if not self.managed_identity_client_id:
                raise RuntimeError(
                    "AZURE_CLIENT_ID is required when GRAPH_AUTH_MODE is 'federated'"
                )
            self.graph_federated_app = (
                self._required(
                    env, "CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID"
                ),
                self._required(env, "SSO_APP_ID"),
            )
        self.graph_bot_credentials: tuple[str, str, str] | None = None
        if self.graph_auth_mode == "bot":
            self.graph_bot_credentials = (
                self._required(
                    env, "CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID"
                ),
                self._required(
                    env, "CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTID"
                ),
                self._required(
                    env, "CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTSECRET"
                ),
            )

        # Delegated (on-behalf-of) file access is enabled only when an Azure Bot
        # OAuth connection is configured for the GRAPH user-authorization handler.
        # Without it, OneDrive/SharePoint attachments are skipped rather than
        # read with an application permission.
        self.user_auth_handler: str | None = None
        if env.get(
            "AGENTAPPLICATION__USERAUTHORIZATION__HANDLERS__"
            f"{USER_AUTH_HANDLER}__SETTINGS__AZUREBOTOAUTHCONNECTIONNAME",
            "",
        ).strip():
            self.user_auth_handler = USER_AUTH_HANDLER

        raw_count = env.get("GRAPH_HISTORY_MESSAGE_COUNT", "10")
        try:
            self.graph_history_message_count = int(raw_count)
        except ValueError as error:
            raise RuntimeError(
                "GRAPH_HISTORY_MESSAGE_COUNT must be an integer"
            ) from error
        if not 1 <= self.graph_history_message_count <= 50:
            raise RuntimeError(
                "GRAPH_HISTORY_MESSAGE_COUNT must be between 1 and 50"
            )

    @staticmethod
    def _required(env, name):
        value = env.get(name, "").strip()
        if not value:
            raise RuntimeError(f"{name} is required")
        return value

    @staticmethod
    def _validate_project_endpoint(value):
        parsed = urlparse(value)
        path_parts = parsed.path.strip("/").split("/")
        if (
            parsed.scheme != "https"
            or not parsed.netloc
            or len(path_parts) != 3
            or path_parts[:2] != ["api", "projects"]
            or not path_parts[2]
            or parsed.params
            or parsed.query
            or parsed.fragment
        ):
            raise RuntimeError(
                "AZURE_AI_PROJECT_ENDPOINT must be a Foundry project endpoint "
                "ending in /api/projects/<project>, not an agent or model endpoint"
            )
