# Token workflow sequence diagrams

These Mermaid sequence diagrams describe the deployed implementation, not a live
Azure trace. A1-A7 match the [README authentication table](../README.md#authentication-flow).
View them on GitHub or in a Mermaid-enabled Markdown preview.

The user-assigned managed identity (UAMI) is the bot identity. The separate
single-tenant sign-in app (`SSO_APP_ID`) handles Teams SSO and holds the
chat-scoped Graph grants. Tokens for Connector, Graph, and Foundry are not
interchangeable; a `/.default` scope request is not the JWT audience itself.
Acquisition steps below run when a suitable cached token is unavailable.

## A1-A2: Incoming activities and Bot Connector authentication

```mermaid
sequenceDiagram
    autonumber
    actor User
    participant Teams
    participant Bot as Azure Bot / Connector
    participant App as App Service / Agents SDK
    participant Identity as Managed identity / Entra ID

    User->>Teams: Send personal message or mention bot in group chat
    Teams->>Bot: Deliver activity
    Bot->>App: A1 POST /api/messages + signed JWT
    Note over Bot,App: JWT audience = BOT_ID (UAMI client ID)
    App->>App: SDK middleware validates JWT
    alt Invalid authentication
        App-->>Bot: Reject request
    else Valid authentication
        App->>App: Route authenticated activity
        Note over App: Ignore channels and unmentioned group messages
        App->>Identity: SDK requests Connector token as bot UAMI
        Identity-->>App: Access token for https://api.botframework.com
        opt Protected Connector-hosted inline image
            App->>Bot: A2 GET trusted attachment URL + Connector bearer
            Bot-->>App: Image bytes
        end
        Note over App: Read context / files and call Foundry as shown below
        App->>Bot: A2 Send reply or typing activity + Connector bearer
        Bot-->>Teams: Deliver activity
        Teams-->>User: Display reply / typing
    end
```

The inbound Bot Framework JWT authenticates the activity; it is **not** a
delegated token for the sender's Graph files.

## A4: App-only Graph tokens for chat-scoped reads

```mermaid
sequenceDiagram
    autonumber
    participant App as App Service
    participant MI as Managed identity endpoint
    participant Entra as Microsoft Entra ID
    participant Graph as Microsoft Graph

    Note over App,Graph: Group chats only, app installed with ChatMessage.Read.Chat RSC consent
    App->>MI: ManagedIdentityCredential selects AZURE_CLIENT_ID
    Note over App,MI: Requested scope = api://AzureADTokenExchange/.default
    MI-->>App: UAMI token for api://AzureADTokenExchange
    App->>Entra: ClientAssertionCredential authenticates as SSO_APP_ID
    Note over App,Entra: UAMI token is client assertion, request Graph /.default scope
    Entra->>Entra: Validate sign-in app federation trusts UAMI principal
    Entra-->>App: App-only Graph access token for sign-in app
    App->>Graph: GET triggering chat message + app-only bearer
    Graph->>Graph: Enforce ChatMessage.Read.Chat for this chat
    Graph-->>App: Message body and attachment references
    opt Triggering message has inline hosted images
        App->>Graph: GET hostedContents metadata / bytes + app-only bearer
        Graph-->>App: Hosted image data
    end
    App->>Graph: GET recent chat messages + app-only bearer
    Graph-->>App: Bounded history (default 10 messages)
```

Teams grants RSC to `webApplicationInfo.id = SSO_APP_ID`, not directly to the
deployed UAMI. This path grants chat-message access, **not** OneDrive/SharePoint
file access. Personal chats do not request Graph history.

## A5: Teams SSO and delegated Graph file access

```mermaid
sequenceDiagram
    autonumber
    actor User
    participant Teams
    participant App as App Service / GRAPH auth handler
    participant TokenService as Azure Bot token service
    participant Entra as Microsoft Entra ID
    participant Graph as Microsoft Graph
    participant Files as SharePoint / OneDrive

    Note over Teams,App: Current message contains a OneDrive / SharePoint reference
    App->>TokenService: Request user authorization via GraphDelegated connection
    opt No reusable delegated token
        Teams->>Entra: Obtain sender's SSO token for sign-in app
        Note over Teams,Entra: API resource = api://botid-BOT_ID, scope = access_as_user
        Entra-->>Teams: Sender's SSO token
        Teams->>App: SSO token-exchange invoke
        App->>TokenService: SDK forwards SSO exchange
        TokenService->>Entra: OBO exchange with SSO token and federated client authentication
        Entra->>Entra: Validate Bot token-store federation on sign-in app
        Note over TokenService,Entra: Federation audience = api://AzureADTokenExchange
        Entra-->>TokenService: Delegated Graph token for sender
    end
    TokenService-->>App: SDK auth.get_token returns delegated Graph token
    Note over App,Graph: Delegated Files.Read.All / Sites.Read.All, never application file roles
    App->>Graph: Resolve current attachment site / drive / item + delegated bearer
    Graph->>Graph: Enforce sender's existing file permissions
    Graph-->>App: File metadata
    App->>Graph: GET item content + delegated bearer
    alt Graph returns bytes directly
        Graph-->>App: File bytes
    else Graph redirects to file download
        Graph-->>App: Pre-authenticated download location
        App->>App: Validate HTTPS host allowlist, strip bearer on host change
        App->>Files: GET allowed download URL without Graph bearer
        Files-->>App: File bytes
    end
```

The connection requests `openid profile offline_access Files.Read.All Sites.Read.All`.
Consent is required according to tenant policy; the user must sign in and have
the app installed for personal use. The SDK can prompt for sign-in when silent
SSO is unavailable. Failed sign-in means files are not read; there is no
app-only file fallback.

Graph limits access to what the sender can open. The **application code**
further limits requests to current-message attachments; the delegated scopes
themselves are broader. OneDrive lookup assumes the sender's own drive.

## A3: Personal-chat files through pre-authenticated download URLs

```mermaid
sequenceDiagram
    autonumber
    participant Teams
    participant Bot as Azure Bot
    participant App as App Service
    participant Files as SharePoint / OneDrive

    Teams->>Bot: Personal-chat attachment with downloadUrl
    Bot->>App: A1 Authenticated activity with attachment
    App->>App: Validate configured SharePoint tenant and HTTPS URL
    App->>Files: A3 GET short-lived pre-authenticated downloadUrl
    Note over App,Files: No Graph or Connector Authorization header
    Files-->>App: File bytes
    App->>App: Validate file type and size
```

The URL is a credential by possession, not an anonymous public file link.
Do not log or share it. Supported inputs are PDF, PNG, JPEG, WebP, and
non-animated GIF, with a default 20 MiB limit per file. These file checks also
apply to Connector and Graph downloads. Unsupported or unreadable files are
reported as skipped; typed URLs and historical attachments are not fetched.

## A7: Foundry inference and the reply

```mermaid
sequenceDiagram
    autonumber
    participant App as App Service
    participant Identity as DefaultAzureCredential / UAMI
    participant Foundry as Foundry project / Responses API
    participant Model as Model deployment
    participant Bot as Bot Connector
    participant Teams

    Note over App: Prompt + bounded group history + supported attachment bytes
    App->>Identity: AIProjectClient requests Foundry authentication
    Identity-->>App: Access token for https://ai.azure.com
    App->>Foundry: A7 responses.create with Foundry bearer and model input
    Note over App,Foundry: SDK manages authentication, no Graph token or API key forwarded
    Foundry->>Foundry: Enforce project-scoped Foundry User RBAC for UAMI
    Foundry->>Model: Inference on configured deployment
    Model-->>Foundry: Generated answer
    Foundry-->>App: Response text
    App->>Bot: A2 Reply + Connector bearer
    Bot-->>Teams: Answer, existing file locations and skipped-file reasons
```

Foundry authentication is independent of `GRAPH_AUTH_MODE`. PDF/image bytes
are encoded into the model request by the bot; the model does not fetch files.
Replies in group chats are visible to participants: sender access does not
prove every recipient has access to the source file. File bytes are handled
in memory; this POC does not establish a Foundry zero-retention guarantee.

## A6: Returning an existing file after personal-chat consent

```mermaid
sequenceDiagram
    autonumber
    actor User
    participant Teams
    participant Bot as Azure Bot / Connector
    participant App as App Service
    participant Source as Original file source
    participant Files as SharePoint / OneDrive

    Note over User,App: Personal chat only, user requested return of a supported existing file
    App->>Bot: A2 Send FileConsentCard + Connector bearer
    Bot-->>Teams: Deliver consent card
    Teams-->>User: Ask permission to upload
    User->>Teams: Accept or decline
    Teams->>Bot: fileConsent/invoke
    Bot->>App: A1 Authenticated callback
    App->>App: Check pending request, expiry, user and conversation
    alt Declined or expired
        App->>Bot: A2 Report decline / expiry, no upload
    else Accepted and valid
        Note over App: Validate callback's pre-authenticated uploadUrl
        App->>Source: Re-read original bytes using that source's authorization
        Note over App,Source: A3 URL, A2 Connector token, or appropriate A4/A5 Graph token
        Source-->>App: Existing file bytes
        App->>Files: A6 PUT chunks to Teams-issued uploadUrl
        Note over App,Files: No Graph or Connector Authorization header
        Files-->>App: Uploaded file metadata
        App->>Bot: A2 Send file-info card + Connector bearer
        Bot-->>Teams: Deliver file-info card
        Teams-->>User: Show returned file
    end
```

File-return consent does not grant Graph access to the original source; the
source must still be readable when fetched again. Group chats receive existing
file locations instead of bot-uploaded files.

## Local-debugging differences

| Path | Local Teams debugging |
| --- | --- |
| A1 inbound | Same SDK JWT validation path. |
| A2 Connector | Generated local bot app and Git-ignored client secret instead of UAMI. |
| A4 chat Graph | `ClientSecretCredential` for the local bot app, which is also the local `webApplicationInfo.id`; per-chat RSC still required. |
| A5 cloud-file references | Not provisioned locally; skipped without delegated authorization, never read app-only. |
| A3/A6 file URLs | Same pre-authenticated URL mechanism, subject to tenant/host checks and personal-chat consent for upload. |
| A7 Foundry | Developer's Azure CLI / VS Code credential, with Foundry RBAC. |

Agents Playground simulates personal-chat handling but does not establish real
Teams SSO or chat RSC consent; inference still calls the live Foundry model.

Sources: [HTTP authentication](../src/app.py),
[runtime credentials and routing](../src/agent.py),
[file operations](../src/files.py), [Graph and inference clients](../src/foundry.py),
[sign-in app federation](../infra/userAuth/userAuthApp.bicep),
and [Azure Bot OAuth connection](../infra/botRegistration/azurebot.bicep).
