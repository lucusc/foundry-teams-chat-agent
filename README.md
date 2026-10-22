# Demo Foundry Agent

A Microsoft 365 Agents Toolkit Python proof of concept that connects Teams to a
shared Microsoft Foundry project.

> [!WARNING]
> **Proof of concept (POC) - work in progress.**
> This project is intended for experimentation and demonstration, not production
> use. Features, configuration, and deployment steps may change without notice.

## Development and deployment

Choose the guide for your task:

- [Teams + Foundry file-access TL;DR](docs/teams-foundry-file-access-tldr.md):
  the setup sequence, permission model, and demo limitations.
- [Local development](docs/local-development.md): Python setup and unit tests.
- [Agents Playground](docs/agents-playground.md): local simulator, without Teams.
- [Local Teams debugging](docs/local-teams-debugging.md): real chats with local breakpoints.
- [Deploy an environment](docs/deploy-environment.md): separate dev or staging deployments and updates.
- [Foundry platform](foundry-platform/README.md): provision or update the shared model infrastructure.

## Behavior

- Uses `AIProjectClient` and the Responses API with `DefaultAzureCredential`.
- Answers personal-chat messages without reading Microsoft Graph history.
- Responds in group chats only when mentioned.
- Explicitly enables the SDK typing indicator with no initial delay and a
  three-second refresh while processing a turn, including chat-context reads,
  attachment processing, and model calls. It stops when a reply is sent or the
  turn ends (including errors). Its appearance depends on the Teams client.
- Reads a bounded window of recent group-chat messages after the chat grants
  `ChatMessage.Read.Chat` resource-specific consent.
- Reads supported PDF and image attachments from the triggering message and
  sends their content to Foundry. It never resolves links typed in message text.
- Reads OneDrive and SharePoint attachments **on behalf of the user who sent
  the message** (Teams SSO + on-behalf-of). It holds no organization-wide
  Microsoft Graph file permission. See
  [File access security](#file-access-security).
- Ignores channel messages and treats chat history as untrusted model context.
- When deployed, uses one user-assigned managed identity for App Service,
  Azure Bot, and Foundry. The same identity authenticates through a federated
  credential as the single-tenant sign-in app for chat-scoped Graph RSC calls.
- In local Teams debugging, uses the local bot's app identity for Graph RSC
  and your developer credential for Foundry. OneDrive and SharePoint
  attachments aren't read locally, because Teams SSO is provisioned only in Azure.

The Teams agent lifecycle is rooted in `m365agents*.yml`. The shared Foundry
account, project, and model are independently managed with AZD under
[`foundry-platform/`](foundry-platform/README.md).

### Current chat and file capabilities

Personal chat means a user's one-to-one conversation with this agent, not their
private conversations with other people.

| Capability | Personal chat with the agent | Group chat |
| --- | --- | --- |
| Read the current message | Yes | Yes, when @mentioned |
| Read previous messages | No; earlier turns are not included in the model request | Yes; a bounded window of recent messages, default 10, after chat-specific consent |
| Read files attached to the current request | Yes, for supported files | Yes, for supported files on the message that @mentions the agent |
| Find and read files shared in earlier messages | No | No; historical attachments are not retrieved |

## Configure

Set these values for the workflow you want to run:

- Local Teams debugging: `env/.env.local.user`
- Agents Playground: `env/.env.playground.user`
- Deployed Teams testing: `env/.env.dev`

```dotenv
AZURE_AI_PROJECT_ENDPOINT=https://<account>.services.ai.azure.com/api/projects/<project>
AZURE_AI_MODEL_DEPLOYMENT_NAME=gpt-4.1-mini
GRAPH_HISTORY_MESSAGE_COUNT=10
SHAREPOINT_TENANT_DOMAIN=<tenant-prefix>
MAX_FILE_SIZE_BYTES=20971520
```

`SHAREPOINT_TENANT_DOMAIN` is the SharePoint tenant prefix (for example,
`contoso`, not a URL). `MAX_FILE_SIZE_BYTES` defaults to 20 MiB. Supported
files are PDF, PNG, JPEG, WebP, and non-animated GIF; Office and text files are
reported as unsupported.

The local Teams workflow automatically sets `GRAPH_AUTH_MODE=bot` and reuses
the generated bot connection's tenant ID, client ID, and secret for Graph RSC.
No extra Graph secret or tenant-wide permission is required. The deployed
environment sets `GRAPH_AUTH_MODE=federated`, `SSO_APP_ID`, and the
`AGENTAPPLICATION__USERAUTHORIZATION__HANDLERS__GRAPH__...` settings from
Bicep.

## Credential flow

`DefaultAzureCredential` is the outbound credential chain. In the deployed App
Service, `AZURE_CLIENT_ID` selects the provisioned user-assigned managed
identity; locally, Azure CLI or VS Code credentials can satisfy the chain.

The application passes credentials—not reusable raw tokens—to the downstream
clients:

1. The Graph history reader requests an app-only token for
   `https://graph.microsoft.com/.default`. Microsoft Graph additionally checks
   that the installed app has `ChatMessage.Read.Chat` consent for the requested
   group chat. Teams grants that consent to `webApplicationInfo.id`, so:
   - local Teams debugging uses `ClientSecretCredential` for the installed
     local bot;
   - deployed testing uses `ClientAssertionCredential` for the sign-in app,
     with a managed-identity assertion through a federated identity credential.
     No secret is stored.
2. OneDrive and SharePoint attachments use a **delegated** Graph token for the
   sender. The Agents SDK `GRAPH` user-authorization handler gets it from the
   Azure Bot OAuth connection `GraphDelegated`, which runs the on-behalf-of
   exchange of the user's Teams SSO token. The app-only credential is never
   used for these files.
3. `AIProjectClient` requests Foundry-scoped tokens and configures its
   OpenAI-compatible Responses client to use them for inference. This always
   uses `DefaultAzureCredential`, independently of Graph authentication.

Incoming Bot Service authentication remains separate and is handled by the
Microsoft 365 Agents SDK connection manager.

## Authentication flow

| Hop | Identity and credential | Audience | Authorization and environment |
| --- | --- | --- | --- |
| A1. Teams → bot | Bot Framework validates the inbound JWT | `BOT_ID` | Azure Bot registration; SDK validation applies locally and in Azure. |
| A2. Bot → Bot Connector | UAMI in Azure; local bot app via the Agents SDK connection manager while debugging | `https://api.botframework.com` | Azure Bot registration. Used for replies and authenticated inline-image downloads; local auth uses the generated, Git-ignored bot secret. |
| A3. Bot → personal-chat file download URL | No credential | SharePoint | Teams-issued, short-lived pre-authenticated URL. The URL is accepted only for the configured SharePoint tenant and is never logged. |
| A4. Bot → triggering group-chat message | App-only. In Azure, the sign-in app via a UAMI federated credential (`GRAPH_AUTH_MODE=federated`). Locally, the bot app's `ClientSecretCredential`. | `https://graph.microsoft.com/.default` | `ChatMessage.Read.Chat` RSC, configured in the Teams manifest and approved for that chat only. |
| A5. Bot → attached files in OneDrive or SharePoint | **Delegated** token for the message sender: Teams SSO → Azure Bot OAuth connection `GraphDelegated` → on-behalf-of. Deployed only. | Microsoft Graph delegated `Files.Read.All` and `Sites.Read.All` | Limited to files that user can already open. The code requests only attachments on the triggering message. Each user consents once. There are **no** application file permissions. |
| A6. Bot → personal-chat upload URL | No credential | SharePoint | Teams supplies a short-lived pre-authenticated URL after the user accepts the `FileConsentCard`. |
| A7. Bot → Foundry | `DefaultAzureCredential` | `https://ai.azure.com` | Foundry User RBAC assignment. In Azure the selected identity is the UAMI; locally it is your developer credential. |

## File access security

> [!IMPORTANT]
> **Don't grant this agent Microsoft Graph *application* permissions
> `Files.Read.All` or `Sites.Read.All`.** Those permissions are
> organization-wide: the identity could read every OneDrive and SharePoint file
> in the tenant, with no user present, including files never shared in any chat.
> No app-only permission can limit file access to one chat.

Teams resource-specific consent (RSC) can scope message access to one chat, but
it has no file permission. A file shared in a group chat lives in the sender's
OneDrive, and its permissions belong to OneDrive, not to the chat. Selected
permissions need explicit resource grants in addition to consent; a site grant
can cover multiple files, but it isn't a chat-scoped grant.

So the agent uses **Teams single sign-on with the Microsoft Entra on-behalf-of
(OBO) flow**:

- **The user's permissions are the ceiling.** Graph enforces the sender's own
  access, so the agent can't open anything that user can't open.
- **The code narrows it further.** Only attachments on the message that
  mentioned the agent are resolved, with the token of the user who sent it.
  Typed URLs are never fetched, and the model can't request files.
- **No secrets and no admin role assignment.** The Azure Bot token service and
  App Service authenticate as a single-tenant sign-in app through federated
  identity credentials.
- **Delegated admin consent.** Microsoft's default user-consent policy blocks
  user consent to `Files.Read.All` and `Sites.Read.All`, so an application
  administrator grants tenant-wide consent for the sign-in app's delegated
  scopes. Users still sign in once, and each user's token is limited to that
  user's access.

Residual risk: delegated `*.Read.All` scopes still cover everything *that
user* can open. The code is what limits use to the attachment. For the
trade-off analysis, diagrams, compliance checks, and Microsoft documentation
links, see [ADR-001: Graph access model for chat files](docs/file-access-security.md).

Before deployment, configure `SHAREPOINT_TENANT_DOMAIN` with the tenant prefix
so the bot accepts only that tenant's `sharepoint.com` and `-my.sharepoint.com`
hosts. If it is empty, SharePoint attachments are rejected. The default
20 MiB limit can be changed with `MAX_FILE_SIZE_BYTES` locally or the
`maxFileSizeBytes` parameter in [`infra/azure.bicep`](infra/azure.bicep).

## Agents Playground

In VS Code, run **Debug in Microsoft 365 Agents Playground**. This provisions
the local Playground tooling, starts the Python agent on port 3978, and opens
the Playground.

> [!NOTE]
> **Use for:** Fast local debugging of personal-chat responses, Foundry
> connectivity, and model prompts, with Python breakpoints.
>
> **Requires:** Your local Python service and an Azure credential with Foundry
> access. No Teams app installation, bot registration, or development tunnel is
> needed; inference still calls the live Foundry model.
>
> **Limitations:** Playground simulates Teams conversations but does not create
> real Teams chats or grant `ChatMessage.Read.Chat` consent. This agent's
> group-chat path calls Microsoft Graph, so a simulated group chat cannot test
> history retrieval and can return the chat-context error even when personal
> chat works. See [Agents Playground](docs/agents-playground.md).

## Local Teams debugging

Sign in to Microsoft 365 and Azure through Agents Toolkit, then run one of:

- **Debug in Teams (Edge)**
- **Debug in Teams (Chrome)**
- **Debug in Teams (Desktop)**

The workflow creates the local app registration, starts a development tunnel,
packages the app, and starts the Python process.

> [!NOTE]
> **Use for:** Debugging real Teams message handling, personal-chat responses,
> and consenting group-chat history retrieval with local Python breakpoints.
>
> **Requires:** A Teams tenant that allows custom app installation, the local
> bot app registration, a running development tunnel, and installation of the
> local app with `ChatMessage.Read.Chat` consent in each test group chat. Stop
> Playground before switching workflows; both generate the root `.localConfigs`.
>
> **Limitations:** Local testing uses a separate app identity from the deployed
> managed identity; each app needs its own chat installation and consent.
> OneDrive and SharePoint attachments aren't read locally. Teams SSO and the
> OAuth connection exist only in Azure, and the local bot app has no Graph file
> permissions. The agent reports that sign-in isn't configured.
> Use non-sensitive messages because recent history is sent to Foundry.
> Playground cannot reproduce this chat consent.
> See [Local Teams debugging](docs/local-teams-debugging.md).

## Deployed Teams testing

Configure `env/.env.dev`, then run:

```bash
atk auth login azure
atk auth login m365
atk provision --env dev -i false
atk deploy --env dev -i false
```

Grant the generated user-assigned managed identity access to the shared Foundry
project. Install the generated app package in a test group chat, approve its
request to read that chat's messages, and mention the agent. The first time
you attach a OneDrive or SharePoint file, select **Sign in** and approve the
delegated permissions.

> [!NOTE]
> **Use for:** End-to-end verification of the deployed bot's personal-chat
> responses and consenting group-chat history retrieval with its managed
> identity.
>
> **Requires:** Deployed Azure resources, Foundry access for the managed identity,
> and app installation with `ChatMessage.Read.Chat` consent in each test group
> chat. An administrator must grant consent for the delegated file scopes, and
> each user who shares OneDrive or SharePoint files must add the app for
> personal use.
> Use non-sensitive messages because recent history and file content are sent
> to Foundry.
>
> **Limitations:** This runs in Azure, not under the local Python debugger; use
> App Service logs to investigate failures. Personal-chat success or personal app
> installation does not prove group-chat access. Group replies require a real
> bot mention, and Teams channels are unsupported.
> See [Deploy an environment](docs/deploy-environment.md#4-install-and-verify-in-teams).

## Tests

```bash
python -m pytest tests -q
```

The tests intentionally focus on the POC's critical behavior and package
contract rather than exhaustive coverage.

## Development resources

Official guidance for building Microsoft 365 Teams bots:

- [File access security](docs/file-access-security.md)
  - Why this agent uses delegated on-behalf-of access instead of
    organization-wide Graph file permissions, with links to Microsoft documentation.
- [Enable SSO for your bot app](https://learn.microsoft.com/microsoftteams/platform/bots/how-to/authentication/bot-sso-overview)
  - Teams SSO for bots, which supplies the user token used for on-behalf-of.
- [Agents SDK user authorization with a federated credential](https://learn.microsoft.com/microsoft-365/agents-sdk/azure-bot-user-authorization-federated-credentials)
  - Configure the Azure Bot OAuth connection without client secrets.

- [Microsoft 365 Agents Toolkit (ATK) overview](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/agents-toolkit-fundamentals)
  - Set up the toolkit and understand the development lifecycle.
- [ATK CLI reference](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/microsoft-365-agents-toolkit-cli)
  - Learn the commands and options for provisioning, deploying, and managing apps.
- [Microsoft 365 Agents SDK documentation](https://learn.microsoft.com/en-us/microsoft-365/agents-sdk/)
  - Follow SDK development guidance, including Python quickstarts and samples.
- [Teams bot design guidelines](https://learn.microsoft.com/en-us/microsoftteams/platform/bots/design/bots)
  - Design bot interactions, onboarding, and personal and shared chat experiences.
- [Teams app use cases and capabilities](https://learn.microsoft.com/en-us/microsoftteams/platform/concepts/design/understand-use-cases)
  - Choose appropriate app capabilities and conversation scopes.
