# Requirements: chat file access POC

## Goal

Show how the Teams agent can securely authenticate to read files shared in
**personal** and **group** chats without storing secrets in the deployed
environment (local debugging is an accepted exception; see D-1). The agent must be able to read a shared file and send it back to
the user, or send a link to where it's stored.

Supported ways to share a file:

| ID | User action |
| --- | --- |
| U1 | Drag and drop a file into the compose box |
| U2 | Copy and paste a file or image into the compose box |
| U3 | Attach a file with the **+** button (upload from device or attach cloud file) |
| U4 | Share a file from SharePoint or OneDrive into the chat |

The code must be commented so a reader can follow every authentication step
(see [Code and comment requirements](#code-and-comment-requirements)).

File access must be limited to what the requesting user can already open, and
must never use organization-wide Graph application permissions. The trade-off
between application and delegated access is recorded in
[ADR-001](file-access-security.md) (D-6).

## Baseline (what exists today)

- One user-assigned managed identity (UAMI) is the bot identity, the Foundry
  caller, and the Microsoft Graph caller in the deployed environment
  ([infra/azure.bicep](../infra/azure.bicep)).
- Group chats grant `ChatMessage.Read.Chat` resource-specific consent (RSC),
  and [`GraphChatHistory`](../src/foundry.py) reads recent messages with an
  app-only Graph token.
- Local Teams debugging uses an Entra app registration with a generated
  client secret ([m365agents.local.yml](../m365agents.local.yml)).
- The agent ignores `activity.attachments`.

### Uncommitted manifest edits that must be resolved

[appPackage/manifest.json](../appPackage/manifest.json) has three edits
without matching tests, infrastructure, or code changes. `python -m pytest
tests` currently fails in
[test_manifest.py](../tests/test_manifest.py).

| Edit | Decision | Required companion changes |
| --- | --- | --- |
| `supportsFiles: true` | **Keep.** Required for personal-chat uploads and file consent cards. | Update the manifest test, the manifest description (MAN-4), the README, and add attachment handling (FR-*). |
| `supportsTargetedMessages: true` | **Revert to `false`.** Enables slash commands and is out of scope. | None after reverting. |
| `webApplicationInfo.id: ${{RSC_APP_CLIENT_ID}}` | **Replace with `${{SSO_APP_ID}}`.** No env file or provisioning step defines `RSC_APP_CLIENT_ID`. Teams SSO requires the sign-in app here, and Teams grants RSC consent to this app ID, so app-only Graph calls must authenticate as the same app (D-6). | Provision the sign-in app (INF-1), set `webApplicationInfo.resource` (MAN-1), and update the manifest test (TEST-1). |

## Platform constraints

These facts determine the design. The sources are listed under
[References](#references).

1. **Personal chat uploads need no Graph permission.** With `supportsFiles`
   enabled, a file the user uploads goes to their OneDrive. The bot receives
   an `application/vnd.microsoft.teams.file.download.info` attachment whose
   `content.downloadUrl` is a short-lived, pre-authenticated URL.
2. **Group chats need Microsoft Graph.** File consent and
   `file.download.info` work only in personal chats. In group chats, files
   stay in the sender's OneDrive or a SharePoint site and must be read
   through Graph.
3. **`Files.Read.All` can't resolve sharing URLs.** Accessing a file through a
   sharing URL needs these least-privileged permissions:

   | Graph call | Least-privileged delegated permission |
   | --- | --- |
   | `GET /users/{id}/drive/root:/{path}` and `/drives/{id}/items/{id}/content` | `Files.Read.All` (needed for files other users shared) |
   | `GET /sites/{hostname}:/{path}` (URL to site lookup) | `Sites.Read.All` |
   | `GET /shares/{encodedUrl}` (sharing URL to item) | `Files.ReadWrite` |

   A file in a user's OneDrive (U1–U3) is resolved from the uploader and the
   drive-relative path. A file on a SharePoint team site (U4) also needs
   `Sites.Read.All`, so the POC requests both delegated scopes (D-2).
   `/shares` isn't used because it needs write access.
4. **No chat-scoped file permission exists.** Teams RSC has no file
   permission. The application versions of `Files.Read.All` and
   `Sites.Read.All` are organization-wide and need admin consent: the identity
   could read **every** file in the tenant without a user. Delegated versions
   are limited to what the signed-in user can open, and users can consent to
   them for themselves. The POC therefore uses delegated access through
   Teams SSO and on-behalf-of (D-6).
5. **Reading a single message works with RSC.** `GET
   /chats/{chatId}/messages/{messageId}` accepts `ChatMessage.Read.Chat`, so
   the agent can read the triggering message's `reference` attachments and
   inline images. The Teams activity ID is the Graph message ID.
6. **Creating Entra apps with Bicep requires Bicep v0.36.1 or later.** The
   Microsoft Graph Bicep extension creates the sign-in app and its federated
   credentials. [m365agents.yml](../m365agents.yml) must pin at least this version.
7. **Teams grants RSC to `webApplicationInfo.id`.** Teams SSO requires
   `webApplicationInfo.id` to be the sign-in app. So app-only RSC calls (A4)
   must authenticate as that app. The deployed UAMI does that through a
   federated identity credential, without a secret.

## Identity and authentication design

**Decision:** the UAMI (`BOT_ID`) remains the bot and Foundry identity. A
single-tenant sign-in app (`<resourceBaseName>-userauth`) is created by Bicep.
It's the Teams SSO resource and the OBO client, and holds only delegated Graph
scopes. It has two federated credentials and no secret:
- one for the Azure Bot token service;
- one for the UAMI, used for app-only RSC calls.

No identity holds Graph application file permissions.

| Hop | Who authenticates | Credential | Audience | What authorizes the call |
| --- | --- | --- | --- | --- |
| A1. Teams → bot (inbound) | Bot Framework | JWT checked by `jwt_authorization_middleware` | `BOT_ID` | Azure Bot registration |
| A2. Bot → Bot Connector (replies, inline image download) | UAMI | `MsalConnectionManager` (`UserManagedIdentity`) | `https://api.botframework.com` | Bot registration |
| A3. Bot → personal-chat `downloadUrl` | None (pre-authenticated URL) | — | SharePoint | Teams-issued, short-lived URL |
| A4. Bot → Graph chat messages | Sign-in app (app-only) | `ClientAssertionCredential` with a UAMI assertion (federated credential) | `https://graph.microsoft.com` | `ChatMessage.Read.Chat` RSC for that chat |
| A5. Bot → Graph files | Message sender (delegated) | Teams SSO token → Azure Bot OAuth connection `GraphDelegated` → OBO | `https://graph.microsoft.com` | Delegated `Files.Read.All` and `Sites.Read.All`, consented by the user. Limited to that user's access. |
| A6. Bot → personal-chat upload URL (send file) | None (pre-authenticated URL) | — | SharePoint | User accepted the file consent card |
| A7. Bot → Foundry | UAMI | `DefaultAzureCredential` | `https://ai.azure.com` | Foundry RBAC role |

For local Teams debugging, A4 uses the local bot app registration through
`ClientSecretCredential` (`GRAPH_AUTH_MODE=bot`). That registration requests
**no** Graph file permissions. Teams SSO isn't provisioned locally, so A5 is
unavailable, and the agent reports that sign-in isn't configured. The local
client secret is an accepted exception (D-1): it is
generated by Agents Toolkit, stored only in Git-ignored files, and never used
in the deployed environment. "No secrets" is a guarantee for the deployed
environment only.

## Functional requirements

### Receiving files

The agent must be able to read every file referenced in a message that
mentions it in a group chat, or any message in a personal chat.

| ID | Chat | Action | Expected attachment (verify in S-1) | Read path |
| --- | --- | --- | --- | --- |
| FR-1 | Personal | U1 drag and drop, U3 upload, U2 paste a file | `file.download.info` | A3: `GET downloadUrl` |
| FR-2 | Personal | U2 paste an image | `image/*` with a Bot Connector `contentUrl` | A2: bearer token for the bot identity |
| FR-3 | Personal | U3 attach cloud file, U4 share | `reference` or link (S-1) | A5 |
| FR-4 | Group | U1, U2 file, U3 | Read the triggering message through Graph (A4); `reference` attachment in the sender's OneDrive | A5: `/users/{senderId}/drive/root:/{path}:/content` |
| FR-5 | Group | U2 paste an image | Inline image in the message body or `hostedContents` | A2 or A4 (S-3) |
| FR-6 | Group | U3 attach cloud file, U4 share from SharePoint | `reference` attachment pointing to a SharePoint site | A5: `/sites/{hostname}:/{path}` to find the site, then read its drive item |

- **FR-7:** Process files only when the agent is triggered (any personal
  message, or a group message that mentions the agent). Channels remain
  unsupported.
- **FR-8:** Read only the files attached to the triggering message, with the
  delegated token of the user who sent it. Never fetch an arbitrary URL from
  message text (SEC-2).
- **FR-9:** When a file can't be read, tell the user which file was skipped
  and why (unsupported type, too large, permission not granted, or consent
  missing) without exposing URLs or tokens.

### Using and returning files

- **FR-10:** Send supported file content to the Foundry model as Responses
  API `input_file` or `input_image` items, along with the user's text.
  Supported types (D-3): PDF (`input_file`) and PNG, JPEG, WebP, and
  non-animated GIF (`input_image`). Other types, including Office and text
  files, get the FR-9 "unsupported type" reply. Text extraction for those
  formats is a later enhancement.
- **FR-11:** Reply with a link to where the file is stored (`webUrl`) in
  both chat types. The link doesn't change SharePoint sharing permissions.
  Recipients still need their own access to the file.
- **FR-12 (personal chat only):** Send a file back with a
  `FileConsentCard`, handle the `fileConsent/invoke` accept and decline
  actions, upload to the returned `uploadUrl` (A6), and then send a
  `FileInfoCard`. Group chats can't receive bot-uploaded files, so the agent
  falls back to FR-11.

## Infrastructure and deployment requirements

- **INF-1:** Grant **no** Microsoft Graph application roles. In
  [infra/userAuth/userAuthApp.bicep](../infra/userAuth/userAuthApp.bicep),
  use the Microsoft Graph Bicep extension to create a single-tenant sign-in
  app with:
  - `api://botid-<bot id>/access_as_user`, pre-authorized for the Teams
    clients;
  - only the delegated `openid`, `profile`, `offline_access`, `Files.Read.All`,
    and `Sites.Read.All` scopes;
  - two federated credentials, one for the Azure Bot token service and one for
    the UAMI.
- **INF-2:** In
  [infra/botRegistration/azurebot.bicep](../infra/botRegistration/azurebot.bicep),
  add the *AAD v2 with Federated Credentials* OAuth connection
  `GraphDelegated`. Set `GRAPH_AUTH_MODE=federated`, `SSO_APP_ID`, and the
  Agents SDK `GRAPH` user-authorization handler in the App Service settings.
- **INF-3:** Raise or remove `bicepCliVersion` in
  [m365agents.yml](../m365agents.yml) so it is at least v0.36.1.
- **INF-4:** Document the required deployer role. Creating the sign-in app
  needs permission to register applications, such as **Application
  Developer**. No privileged role is needed to grant application permissions.
- **INF-5:** Local debugging: keep `requiredResourceAccess` empty in
  [aad.manifest.json](../aad.manifest.json), and keep the existing generated
  client secret for the local bot (D-1). Document that local debugging can't
  read OneDrive or SharePoint attachments.
- **INF-6:** Document how to check that no identity holds Graph application
  file roles, and how to remove any that are found.
- **INF-7:** Add no client secret, certificate, or API key to deployed
  configuration.

## Manifest requirements

- **MAN-1:** Keep `supportsFiles: true`, and revert `supportsTargetedMessages`
  to `false`. Set `webApplicationInfo` to
  `{"id": "${{SSO_APP_ID}}", "resource": "api://botid-${{BOT_ID}}"}`, and add
  `token.botframework.com` to `validDomains` (D-6).
- **MAN-2:** Keep `ChatMessage.Read.Chat` as the only RSC permission.
  Tenant-wide Graph permissions don't belong in the manifest.
- **MAN-3:** Increase `version` so Teams picks up the change, and reinstall
  the app in existing test chats.
- **MAN-4:** Update `description.full` to state that the agent reads shared
  files and sends their content to the Foundry model.

## Security requirements

- **SEC-1:** Never grant the application versions of `Files.Read.All` or
  `Sites.Read.All`. They're organization-wide. Document this risk and why
  on-behalf-of was chosen in the README, in
  [file-access-security.md](file-access-security.md), and in code comments
  (D-6).
- **SEC-2 (confused deputy):** The agent must read only files that Teams
  attached to the triggering message (from the activity or the Graph message
  in A4). It must use the delegated token of that message's sender, and never
  fall back to an application identity. It must never read URLs typed in
  message text.
- **SEC-3:** Accept only HTTPS hosts in the tenant's SharePoint domains
  (`<tenant>.sharepoint.com` and `<tenant>-my.sharepoint.com`) plus the Bot
  Connector service URL from the activity.
- **SEC-4:** Enforce a configurable maximum file size (default 20 MB) and
  the FR-10 list of allowed file types before downloading the whole file.
- **SEC-5:** Keep file content in memory for the request only. Don't persist
  or cache it.
- **SEC-6:** Never log `downloadUrl`, `uploadUrl`, tokens, or file content.
  Log file names, content types, sizes, and outcomes only.
- **SEC-7:** Pass file content to the model as untrusted input, the same way
  chat history is handled today.

## Code and comment requirements

- **CODE-1:** Put file handling in a dedicated module (for example,
  `src/files.py`) with a single entry point used by `on_message`.
- **CODE-2:** At every authentication hop (A1–A7), add a comment that states:
  1. Which identity authenticates.
  2. The token audience or scope.
  3. Which permission or consent authorizes the call, and where it's
     configured (Bicep, YAML, manifest, or admin consent).
  4. What happens in local versus deployed environments.
- **CODE-3:** Use the existing Graph credential selection
  (`graph_credential` in [agent.py](../src/agent.py)) for app-only RSC calls
  only. OneDrive and SharePoint calls use the delegated user token from the
  Agents SDK user-authorization handler.
- **CODE-4:** Pass `TokenCredential` objects to clients, not raw tokens, as
  the existing code already does.
- **CODE-5:** Add an "Authentication flow" section to the README that mirrors
  the A1–A7 table.

## Test requirements

- **TEST-1:** Update [test_manifest.py](../tests/test_manifest.py) to expect
  `supportsFiles: true`, `supportsTargetedMessages: false`, and
  `webApplicationInfo.id == "${{SSO_APP_ID}}"`.
- **TEST-2:** Add infrastructure tests that confirm:
  - No Bicep template or Entra manifest grants Graph application roles
    (`appRoleAssignedTo` or `type: 'Role'`).
  - The sign-in app is single-tenant and requests only the delegated scopes in
    INF-1.
  - The Bicep templates don't request `Files.ReadWrite.All` or
    `Sites.ReadWrite.All`.
  - `bicepCliVersion` is at least v0.36.1.
  - No secret settings appear in the Bicep template.
- **TEST-3:** Add unit tests for:
  - Classifying each attachment type (FR-1 through FR-6).
  - Building OneDrive paths and SharePoint site lookups.
  - Rejecting disallowed hosts (SEC-3).
  - Ignoring URLs in message text (SEC-2).
  - Size limits and the FR-10 type list, including rejecting `.docx`
    (SEC-4).
  - Group-chat file handling only when the agent is mentioned (FR-7).
  - Error messages that don't expose URLs (FR-9).
- **TEST-4:** Mock Graph and HTTP calls. Unit tests must not call live
  services.

## Acceptance criteria

Run in the deployed `dev` environment with non-sensitive test files.

| # | Scenario | Pass condition |
| --- | --- | --- |
| AC-1 | Personal chat: drag and drop a PDF | Answer reflects the file's content |
| AC-2 | Personal chat: paste an image | Answer reflects the image |
| AC-3 | Personal chat: **+** → upload from device | Same as AC-1 |
| AC-4 | Group chat: attach a file and mention the agent | Same as AC-1 |
| AC-5 | Group chat: paste an image and mention the agent | Same as AC-2 |
| AC-6 | Group chat: share a file from a SharePoint team site and mention the agent | Same as AC-1 |
| AC-6a | Any chat: attach a `.docx` file | Agent replies that the type isn't supported yet (FR-9) |
| AC-7 | Any chat: ask for the file location | Reply contains a working link |
| AC-8 | Personal chat: ask for a file back | Consent card, then a file card in OneDrive |
| AC-9 | Paste a SharePoint URL as text for a file the user can't open | Agent refuses to read it (SEC-2) |
| AC-9a | Group chat: first OneDrive or SharePoint attachment from a user | Sign-in card; after consent the original message is answered |
| AC-9b | Review the bot managed identity and local bot app in Enterprise applications | No Graph application permissions |
| AC-10 | Review the deployed configuration | No secrets in app settings, Bicep, or `env/.env.dev` |

## Decisions

All decisions were accepted on 2026-10-20 as recommended.

| ID | Question | Decision |
| --- | --- | --- |
| D-1 | Can local Teams debugging use the generated client secret, or must the POC be secretless everywhere? | **Accepted:** the local client secret is allowed. It's already in use and stays in Git-ignored files. "No secrets" applies to the deployed environment (INF-5, INF-7, AC-10). |
| D-2 | Add `Sites.Read.All` so files on SharePoint team sites (FR-6, U4) can be resolved? | **Accepted:** request the *delegated* `Files.Read.All` and `Sites.Read.All` scopes. Don't use `Files.ReadWrite.All` (INF-1, AC-6). |
| D-3 | Which file types must the model understand? | **Accepted:** PDF and PNG, JPEG, WebP, and non-animated GIF images. Other types get an "unsupported type" reply. Office and text extraction is a later enhancement (FR-10, AC-6a). |
| D-4 | Should group chats process files only when the agent is mentioned? | **Accepted:** yes, matching current behavior (FR-7). |
| D-5 | What does "send the file" mean? | **Accepted:** send content to the model (FR-10) and reply with a link (FR-11) in both chat types, plus send the file back through a consent card in personal chat (FR-12). |
| D-6 | Should files be read with Graph application permissions or with delegated access? | **Accepted: delegated access with Teams SSO and on-behalf-of** (Option 2 in [ADR-001](file-access-security.md)). Application `Files.Read.All` and `Sites.Read.All` are organization-wide. No chat-scoped file permission exists, and selected permissions need per-file admin grants. Delegated access limits the agent to what the sender can open, and the code limits it to the triggering message's attachments. The app is single-tenant, each user consents individually, and file reads are deployed-only; locally, they report that sign-in isn't configured. |

## Spikes to run before implementation

- **S-1:** In the deployed environment, log the attachment content types and
  fields (with URLs redacted) for every row in the
  [Receiving files](#receiving-files) table. Then confirm the expected formats.
- **S-2:** Confirm whether group-chat activities include file attachments or
  only HTML, and whether the Graph lookup in A4 is always needed.
- **S-3:** Confirm that `hostedContents` download works with
  `ChatMessage.Read.Chat` RSC, or whether inline images in group chats must
  come through the Bot Connector.
- **S-4:** Confirm that downloading an inline image from the Bot Connector
  works with the UAMI connection.
- **S-5:** Confirm how to map a OneDrive `contentUrl` to
  `/users/{id}/drive/root:/{path}`, including document libraries whose names
  are localized.
- **S-6:** Confirm that the configured Foundry model deployment accepts
  `input_file` (PDF) and `input_image`.

## Out of scope

- Teams channels and meetings.
- Delegated OneDrive and SharePoint reads in local Teams debugging (D-6).
- Text extraction from Office and text files (D-3).
- Uploading files into group chats.
- Persisting files, indexing them, or using them as RAG sources.
- Other production hardening, such as admin-consent policy and monitoring.

## References

- [Send and receive files with bots](https://learn.microsoft.com/microsoftteams/platform/bots/how-to/bots-filesv4)
- [Resource-specific consent](https://learn.microsoft.com/microsoftteams/platform/graph-api/rsc/resource-specific-consent)
- [Get chatMessage](https://learn.microsoft.com/graph/api/chatmessage-get)
- [Get driveItem content](https://learn.microsoft.com/graph/api/driveitem-get-content)
- [Access shared items](https://learn.microsoft.com/graph/api/shares-get)
- [Get site by path](https://learn.microsoft.com/graph/api/site-getbypath)
- [Microsoft Graph Bicep quickstart](https://learn.microsoft.com/graph/templates/bicep/quickstart-create-bicep-interactive-mode)
- [Enable SSO for your bot app](https://learn.microsoft.com/microsoftteams/platform/bots/how-to/authentication/bot-sso-overview)
- [Agents SDK user authorization with a federated credential](https://learn.microsoft.com/microsoft-365/agents-sdk/azure-bot-user-authorization-federated-credentials)
- [On-behalf-of flow](https://learn.microsoft.com/entra/identity-platform/v2-oauth2-on-behalf-of-flow)
- [File access security](file-access-security.md)
- [Teams app manifest schema v1.30](https://developer.microsoft.com/json-schemas/teams/v1.30/MicrosoftTeams.schema.json)
