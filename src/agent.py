import asyncio
import logging
import os

import httpx
from azure.core.exceptions import AzureError
from azure.identity import (
    ClientAssertionCredential,
    ClientSecretCredential,
    DefaultAzureCredential,
    ManagedIdentityCredential,
)
from dotenv import load_dotenv
from microsoft_agents.activity import (
    Activity,
    ActivityTypes,
    Attachment,
    InvokeResponse,
    load_configuration_from_env,
)
from microsoft_agents.authentication.msal import MsalConnectionManager
from microsoft_agents.hosting.aiohttp import CloudAdapter
from microsoft_agents.hosting.core import (
    AgentApplication,
    MemoryStorage,
    RouteRank,
    TurnContext,
    TurnState,
)
from openai import OpenAIError

from config import Config
from files import (
    FileReadError,
    TeamsFileHandler,
    is_file_return_request,
    requires_user_graph_token,
    should_process_message,
)
from foundry import (
    FoundryModelGateway,
    GraphChatHistory,
    is_group_conversation,
    normalize_message,
)

load_dotenv()
logging.basicConfig(
    level=os.getenv("LOG_LEVEL", "INFO"),
    format="%(asctime)s %(levelname)s %(name)s | %(message)s",
)
logger = logging.getLogger("demo-foundry-agent")

config = Config(os.environ)
agents_sdk_config = load_configuration_from_env(os.environ)

# DefaultAzureCredential provides the application identity used for outbound Azure
# calls. In App Service, AZURE_CLIENT_ID selects the provisioned user-assigned
# managed identity. During local development, the chain can instead use the
# developer's Azure CLI or VS Code credential. This is separate from the Agents
# SDK connection manager below, which authenticates incoming Bot Service traffic.
credential = DefaultAzureCredential(
    managed_identity_client_id=config.managed_identity_client_id
)
# A7: this TokenCredential obtains Foundry tokens for the AI-project audience.
# The deployed caller is the UAMI authorized by the Foundry User RBAC assignment;
# local development uses the developer's Azure CLI or VS Code credential.
gateway = FoundryModelGateway(
    config.project_endpoint,
    config.model_deployment_name,
    credential,
)
# SECURITY: Microsoft Graph access is split on purpose. See
# docs/file-access-security.md.
# - The application credential below can read only chats where the Teams app is
#   installed and RSC-consented (ChatMessage.Read.Chat). It has NO tenant-wide
#   Files.Read.All/Sites.Read.All application permission, because that would let
#   this service read every OneDrive and SharePoint file in the organization.
# - OneDrive/SharePoint attachments are read with the sending user's delegated
#   token instead. Teams SSO plus the Microsoft Entra on-behalf-of (OBO)
#   exchange, run by the Azure Bot token service, bounds access to files that
#   user can already open, attached to the message they sent.
if config.graph_bot_credentials is not None:
    # Local Teams: the bot app (also webApplicationInfo.id) holds the RSC grant.
    graph_credential = ClientSecretCredential(*config.graph_bot_credentials)
elif config.graph_federated_app is not None:
    # Deployed: the RSC grant belongs to the single-tenant UserAuth app. Its
    # federated identity credential trusts this App Service's managed identity,
    # so no client secret exists. The MI token for api://AzureADTokenExchange is
    # the client assertion.
    graph_tenant_id, graph_app_id = config.graph_federated_app
    _assertion_source = ManagedIdentityCredential(
        client_id=config.managed_identity_client_id
    )
    graph_credential = ClientAssertionCredential(
        graph_tenant_id,
        graph_app_id,
        lambda: _assertion_source.get_token(
            "api://AzureADTokenExchange/.default"
        ).token,
    )
else:
    graph_credential = credential
logger.info("Microsoft Graph authentication mode: %s", config.graph_auth_mode)
logger.info(
    "Delegated OneDrive/SharePoint access: %s",
    "enabled (Teams SSO + OBO)" if config.user_auth_handler else "not configured",
)
graph_chat_history = GraphChatHistory(
    graph_credential,
    config.graph_history_message_count,
)

storage = MemoryStorage()
connection_manager = MsalConnectionManager(**agents_sdk_config)
adapter = CloudAdapter(connection_manager=connection_manager)
file_handler = TeamsFileHandler(
    graph_credential,
    config.sharepoint_tenant_domain,
    config.max_file_size_bytes,
    graph_message_reader=graph_chat_history.read_message,
    user_sign_in_enabled=config.user_auth_handler is not None,
)
agent_app = AgentApplication[TurnState](
    storage=storage,
    adapter=adapter,
    connection_manager=connection_manager,
    **agents_sdk_config,
)

GROUP_MESSAGE_KEY = "demo_foundry_agent.group_message"


@agent_app.before_turn
async def prefetch_group_message(context: TurnContext, _state: TurnState) -> bool:
    # Teams sends group-chat file references only through Graph. Reading the
    # triggering message (app-only, chat-scoped RSC) before routing lets the
    # sign-in route below ask for delegated access only when the message actually
    # carries a OneDrive/SharePoint file.
    activity = context.activity
    if (
        getattr(activity, "type", None) == ActivityTypes.message
        and should_process_message(activity)
        and is_group_conversation(activity)
    ):
        chat_id = getattr(getattr(activity, "conversation", None), "id", None)
        message_id = getattr(activity, "id", None)
        if chat_id and message_id:
            try:
                context.turn_state[GROUP_MESSAGE_KEY] = await asyncio.to_thread(
                    graph_chat_history.read_message, chat_id, message_id
                )
            except (AzureError, httpx.HTTPError, RuntimeError, ValueError):
                # on_message retries the read and reports the failure to the user.
                logger.warning(
                    "Could not prefetch the triggering group-chat message",
                    exc_info=True,
                )
    return True


def needs_user_graph_token(context: TurnContext) -> bool:
    activity = context.activity
    if getattr(activity, "type", None) != ActivityTypes.message:
        return False
    if not should_process_message(activity):
        return False
    attachments = list(getattr(activity, "attachments", None) or [])
    group_message = context.turn_state.get(GROUP_MESSAGE_KEY)
    if isinstance(group_message, dict):
        attachments.extend(group_message.get("attachments") or [])
    return requires_user_graph_token(attachments)


def is_file_consent_invoke(context: TurnContext) -> bool:
    activity = context.activity
    return (
        getattr(activity, "type", None) == ActivityTypes.invoke
        and getattr(activity, "name", None) == "fileConsent/invoke"
    )


def file_consent_needs_user_graph_token(context: TurnContext) -> bool:
    return is_file_consent_invoke(context) and file_handler.requires_consent_user_token(
        context.activity
    )


async def get_user_graph_token(context: TurnContext) -> str | None:
    # The token comes from the Azure Bot OAuth connection. Teams SSO supplies the
    # user's token for api://botid-<bot id>, and the Bot token service exchanges
    # it (Microsoft Entra on-behalf-of) for a delegated Graph token. The Agents
    # SDK caches it only for this turn, after a route with auth_handlers ran.
    if not config.user_auth_handler:
        return None
    token_response = await agent_app.auth.get_token(context, config.user_auth_handler)
    return getattr(token_response, "token", None) or None


async def on_graph_sign_in_failure(
    context: TurnContext, _state: TurnState, _handler_id: str | None
):
    await context.send_activity(
        "Sign-in didn't complete, so I can't read OneDrive or SharePoint files "
        "attached to that message. I only read those files on your behalf, using "
        "your own access. Mention me again to retry."
    )


if config.user_auth_handler:
    agent_app.auth.on_sign_in_failure(on_graph_sign_in_failure)


@agent_app.conversation_update("membersAdded")
async def on_members_added(context: TurnContext, _state: TurnState):
    await context.send_activity(
        "Demo Foundry Agent is ready. Ask a question in a personal chat, "
        "or mention it in a consenting group chat."
    )


async def on_message(context: TurnContext, _state: TurnState):
    activity = context.activity
    if not should_process_message(activity):
        if (
            getattr(
                getattr(activity, "conversation", None),
                "conversation_type",
                None,
            )
            == "channel"
        ):
            logger.warning("Ignoring unsupported Teams channel message")
        return

    group_chat = is_group_conversation(activity)
    user_text = normalize_message(getattr(activity, "text", None))
    try:
        chat_history = []
        if group_chat:
            chat_id = getattr(getattr(activity, "conversation", None), "id", None)
            if not chat_id:
                raise RuntimeError("Teams group-chat activity has no conversation ID")
            chat_history = await asyncio.to_thread(graph_chat_history.read, chat_id)

        connector_token = None
        if file_handler.requires_connector_token(activity):
            # A2: the Agents SDK's service connection authenticates as the bot
            # UAMI in Azure or the local bot app during Teams debugging, for the
            # https://api.botframework.com audience authorized by the Bot registration.
            connector_token = (
                await connection_manager.get_default_connection().get_access_token(
                    "https://api.botframework.com", []
                )
            )
        user_graph_token = None
        if needs_user_graph_token(context):
            user_graph_token = await get_user_graph_token(context)
        file_result = await asyncio.to_thread(
            file_handler.process_message_files,
            activity,
            group_chat,
            connector_token,
            user_graph_token,
            context.turn_state.get(GROUP_MESSAGE_KEY),
        )

        if not user_text and not file_result.files:
            if file_result.skipped:
                details = "\n".join(
                    f"- {skipped.name}: {skipped.reason}"
                    for skipped in file_result.skipped
                )
                await context.send_activity(f"No supported file was read:\n{details}")
            else:
                await context.send_activity("Send a question for Demo Foundry Agent.")
            return

        request_text = user_text or "Please summarize the attached file(s)."
        answer = await asyncio.to_thread(
            gateway.ask,
            request_text,
            chat_history,
            file_result.files,
        )
    except (
        AzureError,
        FileReadError,
        httpx.HTTPError,
        OpenAIError,
        RuntimeError,
        ValueError,
    ):
        logger.exception("Foundry agent request failed")
        await context.send_activity(
            "Demo Foundry Agent could not read the chat context or call the model. "
            "Confirm its Foundry access and group-chat consent, then try again."
        )
        return

    reply = answer
    links = [
        f"- {file.name}: {file.web_url}"
        for file in file_result.files
        if file.web_url
    ]
    if links:
        reply += "\n\nAttached file locations (recipients still need access):\n"
        reply += "\n".join(links)
    if file_result.skipped:
        reply += "\n\nSkipped files:\n" + "\n".join(
            f"- {skipped.name}: {skipped.reason}"
            for skipped in file_result.skipped
        )
    if group_chat and is_file_return_request(user_text):
        reply += (
            "\n\nTeams does not support sending a bot-uploaded file to group chats; "
            "use the file location above."
        )

    await context.send_activity(reply)
    if (
        not group_chat
        and is_file_return_request(user_text)
        and file_result.files
    ):
        user_id = getattr(getattr(activity, "from_property", None), "id", None)
        conversation_id = getattr(
            getattr(activity, "conversation", None), "id", None
        )
        if not user_id or not conversation_id:
            await context.send_activity(
                "The file cannot be returned because the chat identity is unavailable."
            )
            return
        for shared_file in file_result.files:
            card = file_handler.prepare_file_consent_card(
                shared_file, user_id, conversation_id
            )
            await context.send_activity(
                Activity(
                    type=ActivityTypes.message,
                    attachments=[
                        Attachment(
                            content_type=card["contentType"],
                            content=card["content"],
                        )
                    ],
                )
            )


async def on_file_consent(context: TurnContext, _state: TurnState):
    activity = context.activity
    value = getattr(activity, "value", None) or {}
    action = value.get("action") if isinstance(value, dict) else None
    connector_token = None
    if (
        action == "accept"
        and file_handler.requires_consent_connector_token(activity)
    ):
        # A2: file-consent callbacks use the same Bot Connector audience and
        # local/deployed service identity as attachment-image downloads.
        connector_token = (
            await connection_manager.get_default_connection().get_access_token(
                "https://api.botframework.com", []
            )
        )

    user_graph_token = None
    if file_consent_needs_user_graph_token(context):
        user_graph_token = await get_user_graph_token(context)

    result = await asyncio.to_thread(
        file_handler.handle_file_consent,
        activity,
        connector_token,
        user_graph_token,
    )
    if result and result.get("status") == "uploaded":
        await context.send_activity(
            Activity(
                type=ActivityTypes.message,
                attachments=[
                    Attachment(
                        content_type="application/vnd.microsoft.teams.card.file.info",
                        content={
                            "uniqueId": result["uniqueId"],
                            "fileType": result["fileType"],
                            "fileName": result["fileName"],
                            "contentUrl": result["contentUrl"],
                        },
                    )
                ],
            )
        )
    elif result and result.get("status") == "declined":
        await context.send_activity("File transfer declined.")
    elif result and result.get("status") == "expired":
        await context.send_activity(
            "That file-transfer request expired. Ask me to send the file again."
        )
    elif result and result.get("status") != "uploaded":
        await context.send_activity(
            "The file could not be transferred. Check the file permission and try again."
        )

    await context.send_activity(
        Activity(
            type=ActivityTypes.invoke_response,
            value=InvokeResponse(status=200),
        )
    )


@agent_app.error
async def on_error(context: TurnContext, error: Exception):
    logger.error(
        "Unhandled agent error",
        exc_info=(type(error), error, error.__traceback__),
    )
    await context.send_activity("The agent encountered an unexpected error.")


# Messages and file-consent exports that need a OneDrive/SharePoint item use the
# GRAPH user-authorization handler. Teams SSO signs in only the user who sent
# the message, and the original message is replayed after sign-in. All other
# messages skip sign-in entirely.
if config.user_auth_handler:
    agent_app.add_route(
        needs_user_graph_token,
        on_message,
        rank=RouteRank.FIRST,
        auth_handlers=[config.user_auth_handler],
    )
    agent_app.add_route(
        file_consent_needs_user_graph_token,
        on_file_consent,
        is_invoke=True,
        rank=RouteRank.FIRST,
        auth_handlers=[config.user_auth_handler],
    )
agent_app.activity(ActivityTypes.message)(on_message)
# AgentApplication has no invoke() decorator; register invoke routes directly.
agent_app.add_route(is_file_consent_invoke, on_file_consent, is_invoke=True)
