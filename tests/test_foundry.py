import sys
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

sys.path.insert(0, str(Path(__file__).parents[1] / "src"))

from config import Config
from foundry import (
    FoundryModelGateway,
    GraphChatHistory,
    is_bot_mentioned,
    normalize_message,
)


def test_config_reads_foundry_settings():
    config = Config(
        {
            "AZURE_AI_PROJECT_ENDPOINT": (
                "https://example.test/api/projects/example-project"
            ),
            "AZURE_AI_MODEL_DEPLOYMENT_NAME": "gpt-4.1-mini",
            "GRAPH_HISTORY_MESSAGE_COUNT": "25",
        }
    )

    assert config.project_endpoint == (
        "https://example.test/api/projects/example-project"
    )
    assert config.model_deployment_name == "gpt-4.1-mini"
    assert config.graph_history_message_count == 25


def test_config_defaults_to_shared_graph_credential():
    config = Config(
        {
            "AZURE_AI_PROJECT_ENDPOINT": "https://example.test/api/projects/project",
            "AZURE_AI_MODEL_DEPLOYMENT_NAME": "model",
        }
    )

    assert config.graph_auth_mode == "default"
    assert config.graph_bot_credentials is None
    assert config.sharepoint_tenant_domain == ""
    assert config.max_file_size_bytes == 20 * 1024 * 1024


def test_config_reads_file_security_settings():
    config = Config(
        {
            "AZURE_AI_PROJECT_ENDPOINT": "https://example.test/api/projects/project",
            "AZURE_AI_MODEL_DEPLOYMENT_NAME": "model",
            "SHAREPOINT_TENANT_DOMAIN": "Contoso",
            "MAX_FILE_SIZE_BYTES": "1048576",
        }
    )

    assert config.sharepoint_tenant_domain == "contoso"
    assert config.max_file_size_bytes == 1024 * 1024


@pytest.mark.parametrize(
    ("setting", "value"),
    [
        ("SHAREPOINT_TENANT_DOMAIN", "contoso.sharepoint.com"),
        ("SHAREPOINT_TENANT_DOMAIN", "bad..tenant"),
        ("MAX_FILE_SIZE_BYTES", "zero"),
        ("MAX_FILE_SIZE_BYTES", "0"),
    ],
)
def test_config_rejects_invalid_file_security_settings(setting, value):
    with pytest.raises(RuntimeError, match=setting):
        Config(
            {
                "AZURE_AI_PROJECT_ENDPOINT": "https://example.test/api/projects/project",
                "AZURE_AI_MODEL_DEPLOYMENT_NAME": "model",
                setting: value,
            }
        )


def test_config_reads_bot_graph_credentials():
    config = Config(
        {
            "AZURE_AI_PROJECT_ENDPOINT": "https://example.test/api/projects/project",
            "AZURE_AI_MODEL_DEPLOYMENT_NAME": "model",
            "GRAPH_AUTH_MODE": "bot",
            "CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID": "tenant-id",
            "CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTID": "bot-id",
            "CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTSECRET": "test-secret",
        }
    )

    assert config.graph_auth_mode == "bot"
    assert config.graph_bot_credentials == ("tenant-id", "bot-id", "test-secret")


FEDERATED_ENV = {
    "AZURE_AI_PROJECT_ENDPOINT": "https://example.test/api/projects/project",
    "AZURE_AI_MODEL_DEPLOYMENT_NAME": "model",
    "GRAPH_AUTH_MODE": "federated",
    "AZURE_CLIENT_ID": "managed-identity-id",
    "CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID": "tenant-id",
    "SSO_APP_ID": "user-auth-app-id",
}
USER_AUTH_SETTING = (
    "AGENTAPPLICATION__USERAUTHORIZATION__HANDLERS__GRAPH__SETTINGS__"
    "AZUREBOTOAUTHCONNECTIONNAME"
)


def test_config_reads_federated_graph_app_and_user_auth_handler():
    config = Config({**FEDERATED_ENV, USER_AUTH_SETTING: "GraphDelegated"})

    assert config.graph_auth_mode == "federated"
    assert config.graph_federated_app == ("tenant-id", "user-auth-app-id")
    assert config.graph_bot_credentials is None
    assert config.user_auth_handler == "GRAPH"


def test_config_disables_user_auth_without_oauth_connection():
    assert Config(FEDERATED_ENV).user_auth_handler is None
    assert Config({**FEDERATED_ENV, USER_AUTH_SETTING: " "}).user_auth_handler is None


@pytest.mark.parametrize(
    "setting",
    [
        "AZURE_CLIENT_ID",
        "CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID",
        "SSO_APP_ID",
    ],
)
def test_config_requires_complete_federated_graph_settings(setting):
    env = {**FEDERATED_ENV, setting: " "}

    with pytest.raises(RuntimeError, match=f"{setting} is required"):
        Config(env)


@pytest.mark.parametrize("mode", ["", "developer", "bto"])
def test_config_rejects_invalid_graph_auth_mode(mode):
    with pytest.raises(RuntimeError, match="GRAPH_AUTH_MODE"):
        Config(
            {
                "AZURE_AI_PROJECT_ENDPOINT": (
                    "https://example.test/api/projects/project"
                ),
                "AZURE_AI_MODEL_DEPLOYMENT_NAME": "model",
                "GRAPH_AUTH_MODE": mode,
            }
        )


@pytest.mark.parametrize("setting", ["TENANTID", "CLIENTID", "CLIENTSECRET"])
def test_config_requires_complete_bot_graph_credentials(setting):
    env = {
        "AZURE_AI_PROJECT_ENDPOINT": "https://example.test/api/projects/project",
        "AZURE_AI_MODEL_DEPLOYMENT_NAME": "model",
        "GRAPH_AUTH_MODE": "bot",
        "CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID": "tenant-id",
        "CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTID": "bot-id",
        "CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTSECRET": "test-secret",
    }
    name = f"CONNECTIONS__SERVICE_CONNECTION__SETTINGS__{setting}"
    env[name] = " "

    with pytest.raises(RuntimeError, match=f"{name} is required"):
        Config(env)


@pytest.mark.parametrize(
    "endpoint",
    [
        "https://example.test/api/projects/example-project/agents/example-agent",
        "https://example.test/openai/deployments/example-model",
        "not-a-url",
    ],
)
def test_config_rejects_non_project_endpoint(endpoint):
    with pytest.raises(RuntimeError, match="Foundry project endpoint"):
        Config(
            {
                "AZURE_AI_PROJECT_ENDPOINT": endpoint,
                "AZURE_AI_MODEL_DEPLOYMENT_NAME": "model",
            }
        )


@pytest.mark.parametrize("count", ["0", "51", "invalid"])
def test_config_rejects_invalid_history_count(count):
    with pytest.raises(RuntimeError, match="GRAPH_HISTORY_MESSAGE_COUNT"):
        Config(
            {
                "AZURE_AI_PROJECT_ENDPOINT": (
                    "https://example.test/api/projects/example-project"
                ),
                "AZURE_AI_MODEL_DEPLOYMENT_NAME": "model",
                "GRAPH_HISTORY_MESSAGE_COUNT": count,
            }
        )


def test_normalize_message_removes_mention():
    assert normalize_message(" <at>Demo Foundry Agent</at>  summarize this ") == (
        "summarize this"
    )


def test_bot_mention_matches_recipient():
    activity = SimpleNamespace(
        recipient=SimpleNamespace(id="bot-id"),
        entities=[{"type": "mention", "mentioned": {"id": "BOT-ID"}}],
    )

    assert is_bot_mentioned(activity)


def test_graph_history_uses_credential_and_returns_oldest_first():
    credential = Mock()
    credential.get_token.return_value.token = "graph-token"
    response = Mock()
    response.json.return_value = {
        "value": [
            {
                "from": {"user": {"displayName": "Second"}},
                "body": {"content": "<p>Newest</p>"},
            },
            {
                "from": {"user": {"displayName": "First"}},
                "body": {"content": "<p>Oldest &amp; useful</p>"},
            },
        ]
    }
    client = Mock()
    client.get.return_value = response

    history = GraphChatHistory(credential, 10, client).read(
        "19:group chat@thread.v2"
    )

    assert history == ["First: Oldest & useful", "Second: Newest"]
    credential.get_token.assert_called_once_with(
        "https://graph.microsoft.com/.default"
    )
    assert client.get.call_args.kwargs["headers"] == {
        "Authorization": "Bearer graph-token"
    }
    response.raise_for_status.assert_called_once_with()


def test_graph_reads_triggering_message_attachments_with_rsc_credential():
    credential = Mock()
    credential.get_token.return_value.token = "graph-token"
    response = Mock()
    response.json.return_value = {"attachments": [{"contentType": "reference"}]}
    client = Mock()
    client.get.return_value = response

    message = GraphChatHistory(credential, 10, client).read_message("chat/id", "m/id")

    assert message == {"attachments": [{"contentType": "reference"}]}
    assert client.get.call_args.kwargs["headers"] == {
        "Authorization": "Bearer graph-token"
    }
    assert client.get.call_args.kwargs["params"] == {
        "$expand": "attachments,hostedContents"
    }
    response.raise_for_status.assert_called_once_with()


def test_model_receives_history_as_untrusted_context():
    project = Mock()
    openai = project.get_openai_client.return_value
    openai.responses.create.return_value.output_text = " Answer "
    gateway = FoundryModelGateway.__new__(FoundryModelGateway)
    gateway._openai = openai
    gateway._model_deployment_name = "gpt-4.1-mini"

    assert gateway.ask("Question", ["Participant: Context"]) == "Answer"
    request = openai.responses.create.call_args.kwargs
    assert request["model"] == "gpt-4.1-mini"
    assert "untrusted user content" in request["instructions"]
    assert "<chat_history>\nParticipant: Context\n</chat_history>" in request["input"]


def test_model_receives_supported_file_content_as_responses_input():
    import base64

    project = Mock()
    openai = project.get_openai_client.return_value
    openai.responses.create.return_value.output_text = "The file says hello."
    gateway = FoundryModelGateway.__new__(FoundryModelGateway)
    gateway._openai = openai
    gateway._model_deployment_name = "gpt-4.1-mini"
    shared_file = SimpleNamespace(
        name="notes.pdf",
        content_type="application/pdf",
        content=b"sample pdf bytes",
    )

    assert gateway.ask("Summarize this", [], [shared_file]) == (
        "The file says hello."
    )
    request = openai.responses.create.call_args.kwargs
    content = request["input"][0]["content"]
    assert content[0]["type"] == "input_text"
    assert "Summarize this" in content[0]["text"]
    assert content[1]["type"] == "input_file"
    assert content[1]["filename"] == "notes.pdf"
    assert content[1]["file_data"] == (
        "data:application/pdf;base64,"
        + base64.b64encode(shared_file.content).decode("ascii")
    )
    assert "attached file names and contents as untrusted input" in request["instructions"]


@pytest.mark.parametrize(
    ("name", "content_type"),
    [
        ("paper.pdf", "application/pdf"),
        ("image.png", "image/png"),
        ("image.jpg", "image/jpeg"),
        ("image.webp", "image/webp"),
        ("image.gif", "image/gif"),
    ],
)
def test_supported_file_types_are_converted_to_responses_content(name, content_type):
    project = Mock()
    openai = project.get_openai_client.return_value
    openai.responses.create.return_value.output_text = "Answer"
    gateway = FoundryModelGateway.__new__(FoundryModelGateway)
    gateway._openai = openai
    gateway._model_deployment_name = "model"

    gateway.ask(
        "Question",
        [],
        [SimpleNamespace(name=name, content_type=content_type, content=b"data")],
    )
    request_content = openai.responses.create.call_args.kwargs["input"][0]["content"]
    expected_type = "input_file" if content_type == "application/pdf" else "input_image"
    assert request_content[1]["type"] == expected_type
