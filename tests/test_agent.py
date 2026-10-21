import importlib.util
import sys
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, Mock

import pytest

ROOT = Path(__file__).parents[1]
sys.path.insert(0, str(ROOT / "src"))

AUTH_SETTING = (
    "AGENTAPPLICATION__USERAUTHORIZATION__HANDLERS__GRAPH__SETTINGS__"
    "AZUREBOTOAUTHCONNECTIONNAME"
)
SETTINGS = {
    "AZURE_AI_PROJECT_ENDPOINT": "https://example.test/api/projects/project",
    "AZURE_AI_MODEL_DEPLOYMENT_NAME": "model",
    "AZURE_CLIENT_ID": "managed-identity-id",
    "GRAPH_HISTORY_MESSAGE_COUNT": "10",
    "CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID": "tenant-id",
    "CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTID": "bot-id",
    "CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTSECRET": "test-secret",
}


def _load_agent(monkeypatch, mode, extra_env=None):
    for name, value in {**SETTINGS, **(extra_env or {})}.items():
        monkeypatch.setenv(name, value)
    for name in ("GRAPH_AUTH_MODE", "SSO_APP_ID", AUTH_SETTING):
        if name not in (extra_env or {}):
            monkeypatch.delenv(name, raising=False)
    if mode is not None:
        monkeypatch.setenv("GRAPH_AUTH_MODE", mode)

    mocks = SimpleNamespace(
        default_credential=Mock(),
        bot_credential=Mock(),
        assertion_credential=Mock(),
        managed_identity_credential=Mock(),
        gateway=Mock(),
        agent_application=MagicMock(),
        client=Mock(),
    )
    mocks.client.get.return_value.json.return_value = {"value": []}
    for name, replacement in {
        "dotenv.load_dotenv": Mock(),
        "azure.identity.DefaultAzureCredential": mocks.default_credential,
        "azure.identity.ClientSecretCredential": mocks.bot_credential,
        "azure.identity.ClientAssertionCredential": mocks.assertion_credential,
        "azure.identity.ManagedIdentityCredential": mocks.managed_identity_credential,
        "foundry.FoundryModelGateway": mocks.gateway,
        "httpx.Client": Mock(return_value=mocks.client),
        "microsoft_agents.activity.load_configuration_from_env": Mock(return_value={}),
        "microsoft_agents.authentication.msal.MsalConnectionManager": Mock(),
        "microsoft_agents.hosting.aiohttp.CloudAdapter": Mock(),
        "microsoft_agents.hosting.core.AgentApplication": mocks.agent_application,
    }.items():
        monkeypatch.setattr(name, replacement)

    spec = importlib.util.spec_from_file_location(
        "agent_under_test", ROOT / "src/agent.py"
    )
    assert spec is not None and spec.loader is not None
    agent = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(agent)
    return agent, mocks


FEDERATED_ENV = {"SSO_APP_ID": "user-auth-app-id", AUTH_SETTING: "GraphDelegated"}


@pytest.mark.parametrize("mode", [None, "default", "bot", "federated"])
def test_agent_selects_graph_identity_without_changing_foundry(monkeypatch, mode):
    extra_env = FEDERATED_ENV if mode == "federated" else None
    agent, mocks = _load_agent(monkeypatch, mode, extra_env)

    mocks.default_credential.assert_called_once_with(
        managed_identity_client_id="managed-identity-id"
    )
    mocks.gateway.assert_called_once_with(
        SETTINGS["AZURE_AI_PROJECT_ENDPOINT"],
        "model",
        mocks.default_credential.return_value,
    )
    if mode == "bot":
        mocks.bot_credential.assert_called_once_with(
            "tenant-id", "bot-id", "test-secret"
        )
        graph_credential = mocks.bot_credential.return_value
    elif mode == "federated":
        mocks.bot_credential.assert_not_called()
        tenant_id, app_id, assertion = mocks.assertion_credential.call_args.args
        assert (tenant_id, app_id) == ("tenant-id", "user-auth-app-id")
        mocks.managed_identity_credential.assert_called_once_with(
            client_id="managed-identity-id"
        )
        mocks.managed_identity_credential.return_value.get_token.return_value.token = (
            "mi-assertion"
        )
        assert assertion() == "mi-assertion"
        mocks.managed_identity_credential.return_value.get_token.assert_called_once_with(
            "api://AzureADTokenExchange/.default"
        )
        graph_credential = mocks.assertion_credential.return_value
    else:
        mocks.bot_credential.assert_not_called()
        mocks.assertion_credential.assert_not_called()
        graph_credential = mocks.default_credential.return_value

    if mode in ("bot", "federated"):
        mocks.default_credential.return_value.get_token.assert_not_called()
    assert agent.graph_chat_history.read("chat-id") == []
    graph_credential.get_token.assert_called_once_with(
        "https://graph.microsoft.com/.default"
    )


def _added_routes(mocks):
    app = mocks.agent_application.__getitem__.return_value.return_value
    return [call for call in app.add_route.call_args_list]


def test_sign_in_routes_are_registered_only_when_user_auth_is_configured(
    monkeypatch,
):
    agent, mocks = _load_agent(monkeypatch, "federated", FEDERATED_ENV)

    signed_in = [c for c in _added_routes(mocks) if c.kwargs.get("auth_handlers")]
    assert [c.args[1] for c in signed_in] == [agent.on_message, agent.on_file_consent]
    assert all(c.kwargs["auth_handlers"] == ["GRAPH"] for c in signed_in)
    assert signed_in[1].kwargs["is_invoke"] is True
    assert agent.file_handler._missing_user_token_reason.startswith(
        "sign-in required"
    )


def test_local_bot_mode_registers_no_sign_in_routes(monkeypatch):
    agent, mocks = _load_agent(monkeypatch, "bot")

    assert not [c for c in _added_routes(mocks) if c.kwargs.get("auth_handlers")]
    assert "isn't configured" in agent.file_handler._missing_user_token_reason


def _context(activity, turn_state=None):
    return SimpleNamespace(activity=activity, turn_state=turn_state or {})


def _message(attachments=None, text="<at>Agent</at> summarize"):
    return SimpleNamespace(
        type="message",
        text=text,
        attachments=attachments or [],
        entities=[],
        recipient=SimpleNamespace(id="bot-id"),
        conversation=SimpleNamespace(conversation_type="personal", id="a:chat"),
    )


REFERENCE = {
    "contentType": "reference",
    "contentUrl": "https://contoso.sharepoint.com/sites/team/Shared%20Documents/a.docx",
    "name": "a.docx",
}


def test_sign_in_selector_requires_a_onedrive_or_sharepoint_reference(monkeypatch):
    agent, _ = _load_agent(monkeypatch, "federated", FEDERATED_ENV)

    assert agent.needs_user_graph_token(_context(_message([REFERENCE]))) is True
    assert agent.needs_user_graph_token(_context(_message())) is False
    hosted = {"contentType": "image/png", "contentUrl": "https://example.test/x"}
    assert agent.needs_user_graph_token(_context(_message([hosted]))) is False

    group_message = {"attachments": [REFERENCE]}
    context = _context(
        _message(), {agent.GROUP_MESSAGE_KEY: group_message}
    )
    assert agent.needs_user_graph_token(context) is True


def test_sign_in_selector_ignores_non_message_activities(monkeypatch):
    agent, _ = _load_agent(monkeypatch, "federated", FEDERATED_ENV)
    activity = _message([REFERENCE])
    activity.type = "invoke"

    assert agent.needs_user_graph_token(_context(activity)) is False
