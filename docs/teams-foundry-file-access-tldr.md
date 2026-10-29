# TL;DR: Teams bot + Foundry with attached-file access

## The design in one minute

**The bot reads attachments; Foundry analyzes their contents.** This Python
demo automatically selects between two Foundry-managed prompt agents through
`AIProjectClient` and the Responses API: general help and summarization.
It requires those two Agent Service agents but no File Search index.

| Operation | Authorization |
| --- | --- |
| Personal-chat upload | Teams supplies a short-lived pre-authenticated download URL; no Graph file permission. |
| Read the triggering group message, recent history, and inline hosted images | App-only `ChatMessage.Read.Chat` resource-specific consent (RSC), approved for that chat. |
| Read an attached OneDrive/SharePoint reference | Sender's delegated Graph token, obtained through Teams SSO and the Azure Bot OAuth connection's on-behalf-of exchange. |
| Classify the request and call the selected Foundry agent | Bot managed identity with project-scoped **Foundry User** RBAC. |

**Never grant Graph application `Files.Read.All` or `Sites.Read.All`.**
Delegated permissions with those names are intentional: Graph limits each
token to what its user can access. RSC authorizes chat messages, not files.
Application code further limits file reads to the current message's attachments.

## Set up your own copy

1. **Prepare accounts and tools.** Use Azure and Microsoft 365 in the same
   Entra tenant, Python 3.11, Node.js 22, Azure CLI, Azure Developer CLI, and
   Microsoft 365 Agents Toolkit (ATK). You need permission to create Azure
   resources and Entra apps, assign Foundry RBAC (or help from an administrator),
   upload custom Teams apps, and approve chat RSC. Plan for Azure/model charges.
2. **Create the model platform.** Follow the
   [Foundry platform guide](../foundry-platform/README.md) to provision the
   account, project, and model with AZD, or reuse an existing compatible project.
   Record the project endpoint, model deployment name, and project resource ID.
   The supplied default is `gpt-4.1-mini`; PDF/image input requires a model that
   supports those inputs. This platform lifecycle is separate from the bot.
3. **Configure a fresh bot environment.** Follow
   [deployment steps 1-2](deploy-environment.md#1-choose-the-environment-and-sign-in).
   Set `AZURE_AI_PROJECT_ENDPOINT`, `AZURE_AI_MODEL_DEPLOYMENT_NAME`,
   `AZURE_AI_GENERAL_AGENT_NAME=teams-poc-general`,
   `AZURE_AI_SUMMARY_AGENT_NAME=teams-poc-summary`,
   `SHAREPOINT_TENANT_DOMAIN` (the prefix, not a URL), and a unique resource
   suffix in the chosen ATK environment. The project endpoint ends in
   `/api/projects/<project>`. Don't copy another deployment's generated IDs or
   overwrite an existing environment unintentionally.
   With developer Foundry access, run the
   [two-agent setup](../foundry-platform/README.md#create-the-two-example-agents)
   before using the bot.
4. **Provision the Teams bot with ATK.** From the repository root, run
   `atk provision --env dev -i false` for a configured dev environment.
   [The templates](../infra/azure.bicep) create App Service, a user-assigned
   managed identity, Azure Bot with Teams enabled, a single-tenant sign-in app,
   and the `GraphDelegated` OAuth connection. They configure:
   - Delegated `openid`, `profile`, `offline_access`, `Files.Read.All`, and
     `Sites.Read.All`; no tenant-wide Graph application file role.
   - `api://botid-<BOT_ID>/access_as_user`, pre-authorized for Teams clients.
   - Two federated credentials: Bot token service for user authorization and
     App Service's managed identity for app-only RSC calls.
   - `GRAPH_AUTH_MODE=federated`, `SSO_APP_ID`, and the SDK `GRAPH`
     user-authorization handler. No deployed client secret is needed.
5. **Grant the two independent permissions.** Follow
   [deployment step 3](deploy-environment.md#3-provision-grant-access-and-deploy-code):
   assign project-scoped **Foundry User** to the bot managed identity's
   **principal/object ID**. Have an Application Administrator or another
   authorized admin consent to the sign-in app's **delegated** Graph scopes
   where tenant policy requires it. Declaring scopes in Bicep doesn't grant
   consent, and provisioning the platform doesn't grant the bot inference access.
6. **Deploy code and install the Teams package.** Run
   `atk deploy --env dev -i false`, then upload the generated
   `appPackage/build/appPackage.dev.zip` to Teams. The
   [manifest](../appPackage/manifest.json) needs `personal`/`groupChat` scopes,
   `supportsFiles: true`, `ChatMessage.Read.Chat`, and `token.botframework.com`.
   Its `webApplicationInfo.id` is the **sign-in app ID**, while `bots[].botId`
   is the **bot managed identity client ID**; don't interchange them.
7. **Install and consent in each chat.** Add the bot to the group chat and
   approve reading that chat's messages. Each user who signs in for cloud-file
   access must also install the app for personal use. Attach a supported file
   **in the same message** as a real bot mention. Complete sign-in if prompted;
   the SDK resumes the original request. With admin consent already granted,
   users don't need to consent to the same scopes individually.

## Verify before demonstrating

- Personal chat: ask a simple question, then attach a small PDF or image.
- Group chat: post a fact, mention the bot to ask about it, then attach a
  supported file in a new message mentioning the bot. Unmentioned messages
  should receive no reply.
- Cloud access: test both a sender-owned OneDrive PDF and a supported SharePoint
  file. Repeat as another user to exercise that user's sign-in.
- Negative test: use a controlled file reference whose access has been revoked
  for the sender; the bot should report it as unreadable. Also try an unsupported
  file and one over the size limit. Do not add app-only permissions to fix failures.
- If sign-in fails, run **Test Connection** for `GraphDelegated` in Azure Bot.
  Check delegated admin consent, personal installation, and App Service logs.
  Renew chat consent after changing the sign-in identity or manifest permissions.

## Important demo limits

- PDF, PNG, JPEG, WebP, and non-animated GIF only; **20 MiB per file** by default.
- Direct attachment paths in the configured SharePoint tenant only. Arbitrary
  sharing/short links, typed URLs, and earlier-message attachments aren't resolved.
  OneDrive lookup assumes the file is in the sender's own drive; another owner's
  file can fail even when the sender has access.
- Channels are unsupported. File-consent uploads work only in personal chats;
  group replies can provide existing file locations.
- This repository provisions cloud-file SSO **only in Azure**. Local Teams
  debugging can test history/personal uploads; Playground cannot prove real
  Graph consent or cloud-file access.
- Sender access isn't an access check for every group participant. Responses
  are visible in the chat. Use non-sensitive data: history and file bytes go to
  Foundry, and no zero-retention guarantee is established by this demo.

## Consistency review

Reviewed 2026-10-22: the supplied ADR snapshot matched the repository's
[file-access ADR](file-access-security.md) exactly before this review. The core
design agrees with the runtime, manifest, and infrastructure. The ADR now
clarifies supported link formats, the sender-owned OneDrive assumption,
deployment permissions, and the limits of per-user access/retention claims.

The attachment, authentication, Foundry, platform, and manifest checks passed:
**101 mocked tests**. This verifies local contracts, not an actual tenant's
grants, Teams SSO token exchange, or live file/model access. Run the Teams
checks above against your deployment.

Details: [deployment guide](deploy-environment.md) |
[access-model rationale](file-access-security.md) |
[Microsoft Teams file APIs](https://learn.microsoft.com/microsoftteams/platform/bots/how-to/bots-filesv4) |
[federated user authorization](https://learn.microsoft.com/microsoft-365/agents-sdk/azure-bot-user-authorization-federated-credentials).
