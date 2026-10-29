# Local Teams debugging

Use real Teams chats while Python runs on your computer with breakpoints.
This creates a local Teams app and Entra bot registration, not a dev or staging
App Service deployment.

## Prerequisites

- Complete [Local development](local-development.md).
- Install Azure CLI and the Microsoft 365 Agents Toolkit extension for VS Code.
- Use a Microsoft 365 tenant that permits custom app uploads, Entra app
  registration, and group-chat resource-specific consent (RSC).
- Obtain the Foundry project endpoint, model deployment name, and project
  resource ID from its owner or the [Foundry platform guide](../foundry-platform/README.md).
- Use the same Entra tenant for your Microsoft 365 and Azure accounts.

## Sign in and configure

Run from the repository root in Bash:

```bash
read -r -p "Azure tenant ID: " TENANT_ID
az login --tenant "$TENANT_ID"
```

Your CLI user needs **Foundry User** on the project. If missing, have a role
administrator assign role `53ca6127-db72-4b80-b1b0-d745d6d5456d` to your
user object ID at the project resource scope. This is separate from Teams
chat consent.

For a new local setup only:

```bash
read -r -p "Foundry project endpoint: " PROJECT_ENDPOINT
read -r -p "Model deployment name: " MODEL_DEPLOYMENT
read -r -p "SharePoint tenant prefix (for example contoso): " SHAREPOINT_TENANT_DOMAIN

cat > env/.env.local <<EOF
TEAMSFX_ENV=local
APP_NAME_SUFFIX=local
EOF

cat > env/.env.local.user <<EOF
AZURE_AI_PROJECT_ENDPOINT=$PROJECT_ENDPOINT
AZURE_AI_MODEL_DEPLOYMENT_NAME=$MODEL_DEPLOYMENT
AZURE_AI_GENERAL_AGENT_NAME=teams-poc-general
AZURE_AI_SUMMARY_AGENT_NAME=teams-poc-summary
GRAPH_HISTORY_MESSAGE_COUNT=10
SHAREPOINT_TENANT_DOMAIN=$SHAREPOINT_TENANT_DOMAIN
EOF
```

For an existing setup, edit these files and retain generated IDs and secrets.
The project endpoint has the form
`https://<account>.services.ai.azure.com/api/projects/<project>`.

Register the two Foundry prompt agents before starting:

```bash
python foundry-platform/create_agents.py --env-file env/.env.local.user --smoke-test
```

## Start and verify

1. Stop Playground; both workflows overwrite the root `.localConfigs`.
2. Sign in to Microsoft 365 and Azure in VS Code's **Agents Toolkit** view.
3. In **Run and Debug**, choose **Debug in Teams (Edge)**, **(Chrome)**, or
   **(Desktop)**, then press **F5**.
4. Let the tasks create the app registration, start the public development
   tunnel on port `3978`, build the package, and start Python.
5. Install the **local** app in Teams and send a personal-chat message.
6. Set a breakpoint in [agent.py](../src/agent.py) and send another message.
7. Add the local app to a test group chat and approve **Read this chat's
   messages** (`ChatMessage.Read.Chat`). Post a non-sensitive fact, then use
   Teams' mention picker to ask the bot about it.

The local app needs its own installation and consent in each chat. Consent
for the deployed app does not cover it. Channels are unsupported. Recent
group history is sent to the live Foundry model.

## Authentication and limitations

The workflow generates the ignored root `.localConfigs` with `GRAPH_AUTH_MODE=bot`.
Graph history uses the local bot app's generated secret; Foundry uses your
developer credential. Never copy the secret into tracked files or set
`AZURE_CLIENT_SECRET` globally.

Personal-chat uploads, inline images, and consenting group history work.
OneDrive/SharePoint attachments report that sign-in is not configured:
Teams SSO and the OAuth connection exist only in the
[deployed environment](deploy-environment.md). Do not add Graph application
file permissions to work around this.

## Troubleshooting

| Symptom | Check |
| --- | --- |
| Bot returns 401 | Restart the full debug workflow; root `.localConfigs` must contain the bot client ID, secret, and `CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID`. |
| Graph uses `AzureCliCredential` or returns 403 | Restart until startup shows `Microsoft Graph authentication mode: bot`; verify local-app consent in that chat. |
| Foundry returns 401/403 | Check Azure CLI sign-in, the project endpoint, and your project-scoped Foundry User role; allow propagation time. |
| Group chat has no reply | Use a real mention and approve chat consent. Update/reinstall the local app if its permissions changed. |
| `aadApp/update` reports `Resource ''` not found | Keep editor-only `$schema` metadata out of [aad.manifest.json](../aad.manifest.json); do not delete the bot's credentials. |
| Configuration changes have no effect | Restart the full **Debug in Teams** workflow to regenerate `.env`. |

When finished, stop debugging. Remove the local Teams app and its separate
Entra app registration when no longer needed; deleting cloud resource groups
does not remove them.
