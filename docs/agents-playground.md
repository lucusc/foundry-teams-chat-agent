# Agents Playground

Run the Python bot on your computer and chat with it in the local simulator.
No Teams account, bot registration, tunnel, or App Service deployment is needed.
Foundry calls are live and incur model charges.

## Prerequisites

- Complete [Local development](local-development.md).
- Install Node.js 22 LTS, Azure CLI, and the Microsoft 365 Agents Toolkit
  extension for VS Code.
- Obtain a Foundry project endpoint, model deployment name, and project
  resource ID from its owner, or [provision the platform](../foundry-platform/README.md).
- Your Azure CLI user needs project-scoped **Foundry User** access.

## Sign in and grant access

Run from the repository root in Bash. Use the tenant that owns the project:

```bash
read -r -p "Azure tenant ID: " TENANT_ID
read -r -p "Foundry project resource ID: " PROJECT_ID
az login --tenant "$TENANT_ID"
```

If access has not already been granted, an Azure role administrator can run:

```bash
DEVELOPER_ID=$(az ad signed-in-user show --query id -o tsv)
az role assignment create \
  --assignee-object-id "$DEVELOPER_ID" \
  --assignee-principal-type User \
  --role 53ca6127-db72-4b80-b1b0-d745d6d5456d \
  --scope "$PROJECT_ID"
```

This assigns access to the signed-in CLI user; if an administrator runs it
on your behalf, give them your user object ID instead. Contributor alone
does not grant inference access. Allow several minutes for role propagation.

## Configure and run

For a new Playground setup, run:

```bash
read -r -p "Foundry project endpoint: " PROJECT_ENDPOINT
read -r -p "Model deployment name: " MODEL_DEPLOYMENT

cat > env/.env.playground <<EOF
TEAMSFX_ENV=playground
TEAMSAPPTESTER_PORT=56150
EOF

cat > env/.env.playground.user <<EOF
AZURE_AI_PROJECT_ENDPOINT=$PROJECT_ENDPOINT
AZURE_AI_MODEL_DEPLOYMENT_NAME=$MODEL_DEPLOYMENT
AZURE_AI_GENERAL_AGENT_NAME=teams-poc-general
AZURE_AI_SUMMARY_AGENT_NAME=teams-poc-summary
GRAPH_HISTORY_MESSAGE_COUNT=10
EOF
```

For an existing setup, edit those files instead of overwriting them.
The endpoint must have the form
`https://<account>.services.ai.azure.com/api/projects/<project>`,
not a published-agent endpoint.

Register the two prompt agents before starting:

```bash
python foundry-platform/create_agents.py --env-file env/.env.playground.user --smoke-test
```

1. Stop any local Teams debugging session.
2. Select **Debug in Microsoft 365 Agents Playground** in VS Code's
   **Run and Debug** menu and press **F5**.
3. Wait for Python on port `3978` and Playground on port `56150`.
4. Send a personal-chat message: `Say hello in one sentence`.
5. Confirm a model answer and no authentication errors in the Python terminal.
   Set a breakpoint in [agent.py](../src/agent.py) to debug.
   The reply should identify `teams-poc-general`. Send `Summarize: We chose
   blue; Sam will send the draft Friday` and confirm `teams-poc-summary`.

The workflow installs Playground tooling and generates the ignored root
`.env`. Restart the full debug workflow after changing configuration.
Incoming messages are anonymous in this workflow only; never expose it
publicly. Outbound Foundry calls still use your Azure credential.

## Limits and troubleshooting

| Symptom or limitation | Action |
| --- | --- |
| Foundry 401/403 | Check `az login`, the project endpoint, and your Foundry User assignment. |
| Model not found | Use the exact model deployment name, not an agent name. |
| Agent not found | Run the agent setup script and check both configured agent names in the same project. |
| Startup fails | Select `.venv`, activate it, and check ports `3978` and `56150` are free. |
| Group-history request fails | Playground cannot grant real Teams `ChatMessage.Read.Chat` consent. Use [local Teams](local-teams-debugging.md). |
| OneDrive/SharePoint sign-in is unavailable | Expected: Teams SSO exists only in the [deployed environment](deploy-environment.md). |

Playground and local Teams both overwrite the root `.localConfigs`; do not run them
at the same time.
