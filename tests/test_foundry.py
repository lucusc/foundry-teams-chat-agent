import sys
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

import httpx
import pytest

sys.path.insert(0, str(Path(__file__).parents[1] / "src"))

from config import Config
from foundry import (
    FoundryAgentGateway,
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
    assert config.general_agent_name == "teams-poc-general"
    assert config.summary_agent_name == "teams-poc-summary"
    assert config.graph_history_message_count == 25


def test_config_reads_custom_agent_names():
    config = Config({
        "AZURE_AI_PROJECT_ENDPOINT": "https://example.test/api/projects/project",
        "AZURE_AI_MODEL_DEPLOYMENT_NAME": "model",
        "AZURE_AI_GENERAL_AGENT_NAME": " custom-general ",
        "AZURE_AI_SUMMARY_AGENT_NAME": " custom-summary ",
    })
    assert config.general_agent_name == "custom-general"
    assert config.summary_agent_name == "custom-summary"


@pytest.mark.parametrize(
    "names",
    [
        {"AZURE_AI_GENERAL_AGENT_NAME": ""},
        {"AZURE_AI_SUMMARY_AGENT_NAME": "bad name"},
        {"AZURE_AI_GENERAL_AGENT_NAME": "x" * 64},
        {"AZURE_AI_GENERAL_AGENT_NAME": "same", "AZURE_AI_SUMMARY_AGENT_NAME": "same"},
    ],
)
def test_config_rejects_invalid_or_duplicate_agents(names):
    with pytest.raises(RuntimeError, match="AGENT_NAME|must be different"):
        Config({
            "AZURE_AI_PROJECT_ENDPOINT": "https://example.test/api/projects/project",
            "AZURE_AI_MODEL_DEPLOYMENT_NAME": "model",
            **names,
        })


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
    assert "params" not in client.get.call_args.kwargs
    assert client.get.call_args.args == (
        "https://graph.microsoft.com/v1.0/chats/chat%2Fid/messages/m%2Fid",
    )
    response.raise_for_status.assert_called_once_with()


@pytest.mark.parametrize("inline_image", [False, True])
def test_graph_message_uses_supported_endpoints(inline_image):
    credential = Mock()
    credential.get_token.return_value.token = "graph-token"
    message = {
        "attachments": [{"contentType": "reference", "name": "notes.pdf"}],
        "body": {
            "content": (
                '<img src="../messages/message-id/hostedContents/image-id/$value">'
                if inline_image else "<p>Read the attachment</p>"
            )
        },
    }
    requests = []

    def respond(request):
        requests.append(request)
        if request.url.query:
            return httpx.Response(400, json={"error": {"code": "BadRequest"}})
        if request.url.path.endswith("/hostedContents"):
            return httpx.Response(
                200, json={"value": [{"id": "image-id", "contentType": "image/png"}]}
            )
        return httpx.Response(200, json=message)

    client = httpx.Client(transport=httpx.MockTransport(respond))
    result = GraphChatHistory(credential, 10, client).read_message(
        "19:group@thread.v2", "message-id"
    )

    assert result["attachments"] == message["attachments"]
    assert len(requests) == (2 if inline_image else 1)
    assert requests[0].url.raw_path == (
        b"/v1.0/chats/19%3Agroup%40thread.v2/messages/message-id"
    )
    if inline_image:
        assert result["hostedContents"] == [
            {"id": "image-id", "contentType": "image/png"}
        ]
        assert requests[1].headers["Authorization"] == (
            requests[0].headers["Authorization"]
        )


@pytest.mark.parametrize(
    ("status", "payload", "error"),
    [
        (403, {"error": {"code": "Forbidden"}}, httpx.HTTPStatusError),
        (200, {"value": {}}, RuntimeError),
    ],
)
def test_graph_hosted_content_failures_are_reported(status, payload, error):
    credential = Mock()
    credential.get_token.return_value.token = "graph-token"

    def respond(request):
        if request.url.path.endswith("/hostedContents"):
            return httpx.Response(status, json=payload)
        return httpx.Response(
            200,
            json={"body": {"content": '<img src="../hostedContents/image-id/$value">'}},
        )

    client = httpx.Client(transport=httpx.MockTransport(respond))
    with pytest.raises(error):
        GraphChatHistory(credential, 10, client).read_message("chat-id", "message-id")


@pytest.mark.parametrize("timeouts", [1, 2])
def test_graph_history_retries_one_timeout(timeouts):
    credential = Mock()
    credential.get_token.return_value.token = "graph-token"
    calls = []

    def respond(request):
        calls.append(request)
        if len(calls) <= timeouts:
            raise httpx.ReadTimeout("timed out", request=request)
        return httpx.Response(
            200,
            json={"value": [{"from": {"user": {"displayName": "A"}}, "body": {"content": "hi"}}]},
        )

    history = GraphChatHistory(
        credential, 10, httpx.Client(transport=httpx.MockTransport(respond))
    )
    if timeouts == 1:
        assert history.read("chat-id") == ["A: hi"]
    else:
        with pytest.raises(httpx.ReadTimeout):
            history.read("chat-id")
    assert len(calls) == 2


def make_gateway(route='{"agent":"general"}', answer=" Answer "):
    with patch("foundry.AIProjectClient"):
        gateway = FoundryAgentGateway(
            "https://example.test/api/projects/project", "gpt-4.1-mini", Mock(),
            "general-agent", "summary-agent",
        )
    gateway._openai.responses.create.side_effect = [
        SimpleNamespace(output_text=route),
        SimpleNamespace(output_text=answer),
    ]
    return gateway


@pytest.mark.parametrize("route", ["general", "summary"])
def test_selected_foundry_agent_receives_history(route):
    gateway = make_gateway(f'{{"agent":"{route}"}}')

    assert gateway.ask("Question", ["Participant: Context"]) == (
        f"**{route}-agent**\n\nAnswer"
    )
    calls = gateway._openai.responses.create.call_args_list
    assert len(calls) == 2
    router = calls[0].kwargs
    assert router["model"] == "gpt-4.1-mini"
    assert router["input"] == "Question"
    assert router["store"] is False
    assert router["text"]["format"]["strict"] is True
    schema = router["text"]["format"]["schema"]
    assert schema["properties"]["agent"]["enum"] == ["general", "summary"]
    request = calls[1].kwargs
    assert request["extra_body"]["agent_reference"] == {
        "name": f"{route}-agent", "type": "agent_reference",
    }
    assert "model" not in request
    assert "instructions" not in request
    assert request["store"] is False
    assert "previous_response_id" not in request
    assert "conversation" not in request
    assert "<chat_history>\nParticipant: Context\n</chat_history>" in request["input"]


@pytest.mark.parametrize(
    "route",
    ["", "not json", "null", "[]", "{}", '{"agent":"other"}',
     '{"agent":[]}', '{"agent":"general","extra":true}'],
)
def test_invalid_routing_fails_without_calling_an_agent(route):
    gateway = make_gateway(route)
    with pytest.raises(RuntimeError, match="Foundry routing"):
        gateway.ask("Question", [])
    assert gateway._openai.responses.create.call_count == 1


def test_routing_failure_does_not_silently_fall_back():
    gateway = make_gateway()
    gateway._openai.responses.create.side_effect = RuntimeError("routing unavailable")
    with pytest.raises(RuntimeError, match="routing unavailable"):
        gateway.ask("Question", [])
    assert gateway._openai.responses.create.call_count == 1


@pytest.mark.parametrize("answer", ["", " ", None])
def test_empty_agent_response_fails(answer):
    gateway = make_gateway(answer=answer)
    with pytest.raises(RuntimeError, match="agent returned no text"):
        gateway.ask("Question", [])


def test_agent_error_propagates_without_falling_back_to_direct_model():
    gateway = make_gateway()
    gateway._openai.responses.create.side_effect = [
        SimpleNamespace(output_text='{"agent":"general"}'),
        RuntimeError("agent not found"),
    ]
    with pytest.raises(RuntimeError, match="agent not found"):
        gateway.ask("Question", [])
    assert gateway._openai.responses.create.call_count == 2


def test_model_receives_supported_file_content_as_responses_input():
    import base64

    gateway = make_gateway('{"agent":"summary"}', "The file says hello.")
    shared_file = SimpleNamespace(
        name="notes.pdf",
        content_type="application/pdf",
        content=b"sample pdf bytes",
    )

    assert gateway.ask("Summarize this", [], [shared_file]) == (
        "**summary-agent**\n\nThe file says hello."
    )
    request = gateway._openai.responses.create.call_args.kwargs
    content = request["input"][0]["content"]
    assert content[0]["type"] == "input_text"
    assert "Summarize this" in content[0]["text"]
    assert content[1]["type"] == "input_file"
    assert content[1]["filename"] == "notes.pdf"
    assert content[1]["file_data"] == (
        "data:application/pdf;base64,"
        + base64.b64encode(shared_file.content).decode("ascii")
    )
    router = gateway._openai.responses.create.call_args_list[0].kwargs
    assert router["input"] == "Summarize this"


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
    gateway = make_gateway()

    gateway.ask(
        "Question",
        [],
        [SimpleNamespace(name=name, content_type=content_type, content=b"data")],
    )
    request_content = gateway._openai.responses.create.call_args.kwargs["input"][0]["content"]
    expected_type = "input_file" if content_type == "application/pdf" else "input_image"
    assert request_content[1]["type"] == expected_type
