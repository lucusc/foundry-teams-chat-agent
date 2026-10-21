# Deploy a dev or staging environment

Deploy the Teams bot to Azure. No local debugger, Playground, or tunnel is
involved. Run commands from the repository root in the same Bash terminal
(Linux, macOS, or Windows with WSL).

## Prerequisites

- Install Python 3.11, Node.js 22 LTS, Azure CLI, and the ATK CLI:
  `npm install -g @microsoft/m365agentstoolkit-cli`.
- Use Azure and Microsoft 365 accounts in the same Entra tenant.
- Have resource creation permissions (for example Contributor) and permission
  to assign Foundry roles (for example Role Based Access Control Administrator
  at the relevant scope, or help from a role administrator).
- Be allowed to register Entra applications. Your Teams tenant must allow
  custom app uploads and chat-specific consent.
- Obtain a Foundry project endpoint, model deployment name, and project
  resource ID, or [provision Foundry separately](../foundry-platform/README.md).
  ATK does not create Foundry or grant inference access.

The current template creates a paid B1 Linux App Service running Python 3.11,
an Azure Bot, a user-assigned managed identity, and a single-tenant sign-in app
with a federated OAuth connection for Teams SSO. Model usage is also billed.

## 1. Choose the environment and sign in

Choose **one**:

```bash
ENVIRONMENT=dev
APP_RG=rg-foundry-teams-app-dev
```

Or, for a separate staging deployment:

```bash
ENVIRONMENT=staging
APP_RG=rg-foundry-teams-app-staging
```

Then run:

```bash
read -r -p "Azure tenant ID: " TENANT_ID
read -r -p "Azure subscription ID: " SUBSCRIPTION_ID
read -r -p "Azure region: " LOCATION
read -r -p "Foundry project endpoint: " PROJECT_ENDPOINT
read -r -p "Model deployment name: " MODEL_DEPLOYMENT
read -r -p "Foundry project resource ID: " PROJECT_ID
read -r -p "SharePoint tenant prefix (for example contoso): " SHAREPOINT_TENANT_DOMAIN

az login --tenant "$TENANT_ID"
az account set --subscription "$SUBSCRIPTION_ID"
atk auth login azure
atk auth login m365
atk auth list
az account show --query '{subscription:id,tenant:tenantId,name:name}' -o table
```

Verify that both tools use the intended accounts. Sign-ins are separate.
For an existing environment, use its actual resource group and subscription,
not a newly chosen name.

| Environment | ATK settings | Teams package |
| --- | --- | --- |
| dev | `env/.env.dev` | `appPackage/build/appPackage.dev.zip` |
| staging | `env/.env.staging` | `appPackage/build/appPackage.staging.zip` |

These are **separate deployments**, not App Service deployment slots.
Keep resource groups, `RESOURCE_SUFFIX`, and generated Teams/bot IDs distinct.
Both may use the same shared Foundry project if that is intentional; for
model/data isolation, provision a separate platform and use its outputs.

**Updating an existing environment?** Keep its settings, suffix, and generated
IDs. Skip step 2 and use [Update an existing environment](#update-an-existing-environment).

## 2. Create fresh environment settings

Register providers if needed; a subscription administrator can do this:

```bash
az provider register --namespace Microsoft.Web --wait
az provider register --namespace Microsoft.ManagedIdentity --wait
az provider register --namespace Microsoft.BotService --wait
az group create --name "$APP_RG" --location "$LOCATION"
```

**New deployment only:** the next block replaces `env/.env.$ENVIRONMENT`.
The repository's existing dev file may contain another deployment's IDs.
Replace it only if you intend to create a fresh dev system, not update that
system. Do not copy generated dev IDs or its suffix into staging.

```bash
RESOURCE_SUFFIX=$(python3.11 -c 'import secrets; print(secrets.token_hex(3))')

cat > "env/.env.$ENVIRONMENT" <<EOF
TEAMSFX_ENV=$ENVIRONMENT
APP_NAME_SUFFIX=$ENVIRONMENT
AZURE_SUBSCRIPTION_ID=$SUBSCRIPTION_ID
AZURE_RESOURCE_GROUP_NAME=$APP_RG
RESOURCE_SUFFIX=$RESOURCE_SUFFIX
AZURE_AI_PROJECT_ENDPOINT=$PROJECT_ENDPOINT
AZURE_AI_MODEL_DEPLOYMENT_NAME=$MODEL_DEPLOYMENT
GRAPH_HISTORY_MESSAGE_COUNT=10
SHAREPOINT_TENANT_DOMAIN=$SHAREPOINT_TENANT_DOMAIN
EOF
```

- The project endpoint must look like
  `https://<account>.services.ai.azure.com/api/projects/<project>`, not an
  agent endpoint. Use the exact model deployment name.
- `bot` plus `RESOURCE_SUFFIX` names the web app and identity and must be
  4-20 lowercase letters/digits. Use a new suffix if the web app name is taken;
  keep it stable after provisioning.
- The SharePoint prefix is a name such as `contoso`, not a URL.
  History count must be an integer from `1` to `50`.
- Cloud environment files are not ignored. Review before committing; never
  store secrets in them. Personal/secret values belong in the ignored
  `env/.env.<environment>.user` file.

## 3. Provision, grant access, and deploy code

```bash
atk provision --env "$ENVIRONMENT" -i false
```

Wait for success. ATK writes generated IDs into the selected environment file
and builds its Teams package. Do not manually substitute `TEAMS_APP_ID` for
`BOT_ID`: they identify different things.

Read the generated web app and managed identity, then grant project-scoped
**Foundry User** to the identity's **principal/object ID**, not its client ID:

```bash
WEBAPP_ID=$(az deployment group show \
  --resource-group "$APP_RG" --name Create-demo-foundry-agent-resources \
  --query properties.outputs.BOT_AZURE_APP_SERVICE_RESOURCE_ID.value -o tsv)
WEBAPP_NAME=$(az webapp show --ids "$WEBAPP_ID" --query name -o tsv)
BOT_PRINCIPAL_ID=$(az identity show \
  --resource-group "$APP_RG" --name "$WEBAPP_NAME" \
  --query principalId -o tsv)

az role assignment create \
  --assignee-object-id "$BOT_PRINCIPAL_ID" \
  --assignee-principal-type ServicePrincipal \
  --role 53ca6127-db72-4b80-b1b0-d745d6d5456d \
  --scope "$PROJECT_ID"

atk deploy --env "$ENVIRONMENT" -i false
```

Ask a role administrator to grant access if needed. Allow time for role
propagation and the App Service remote dependency build/startup.
Provisioning alone does not upload Python; deployment does not install Teams.

**Do not grant Graph application `Files.Read.All` or `Sites.Read.All`.**
History uses chat-scoped RSC; files use the sender's delegated Teams SSO token.
No organization-wide application file permission or bot secret is needed.
If user consent is restricted, an administrator can approve the **delegated**
permissions on the generated `-userauth` app. See
[File access security](file-access-security.md).

## 4. Install and verify in Teams

Use non-sensitive test messages and files: their contents go to Foundry.

1. Sign in to Teams in the deployment's tenant.
2. Open **Apps > Manage your apps > Upload an app > Upload a custom app**.
3. Upload `appPackage/build/appPackage.dev.zip` for dev, or
   `appPackage/build/appPackage.staging.zip` for staging.
4. In a personal chat, ask `Say hello in one sentence` and confirm an answer.
5. Add that environment's app to a **group chat**, approve **Read this chat's
   messages**, and post a short fact. Use the mention picker to ask the bot
   about the fact. Confirm it uses the history and ignores unmentioned messages.
6. Attach a supported PDF or image from OneDrive/SharePoint and mention the
   bot. Complete **Sign in to share files** and confirm an answer about the file.
   Office/text files are unsupported; the default file limit is 20 MiB.
7. In Azure Bot **Configuration > OAuth Connection Settings**, run **Test
   Connection** on `GraphDelegated` if file sign-in fails.

Consent is per app and per group chat. Personal-chat success does not prove
history access. Update/reinstall packages and renew consent after permission
or sign-in identity changes. Teams channels are unsupported.

For runtime errors:

```bash
az webapp log config --resource-group "$APP_RG" --name "$WEBAPP_NAME" \
  --docker-container-logging filesystem
az webapp log tail --resource-group "$APP_RG" --name "$WEBAPP_NAME"
```

Stop log streaming with **Ctrl+C**; this does not stop the app.
Expect startup to report `Microsoft Graph authentication mode: federated`
and delegated access enabled. The only messaging route is `POST /api/messages`;
a browser GET or a running Azure resource is not an end-to-end test.

## Update an existing environment

Select `ENVIRONMENT` and sign in as in step 1. Preserve the selected file's
generated IDs and `RESOURCE_SUFFIX`; do not run the fresh-file block.

| Change | Action from the repository root |
| --- | --- |
| Python code or dependencies | Run `python -m pytest tests -q` in your [development environment](local-development.md), then `atk deploy --env "$ENVIRONMENT" -i false`. |
| Endpoint, model, history count, or tenant prefix | Edit the selected environment file, then `atk provision --env "$ENVIRONMENT" -i false`. Code deployment alone does not apply settings. |
| Infrastructure or file-size limit | Edit [Bicep](../infra/azure.bicep) or [parameters](../infra/azure.parameters.json), then provision the selected environment. Shared template changes affect every environment next provisioned. |
| Manifest or icons | Edit [appPackage](../appPackage), bump the manifest version, provision, then upload/update that environment's rebuilt package in Teams. Renew consent if needed. |
| Foundry platform | Follow the [platform update guide](../foundry-platform/README.md#update-or-adopt-an-existing-platform). Copy changed outputs into the selected ATK file and provision again. A new project also needs a new Foundry role assignment. |

Repeat the Teams checks after updates. To submit this environment's package
to the organization's app catalog, run `atk publish --env "$ENVIRONMENT" -i false`;
an administrator must approve it. Publishing neither deploys code nor grants
chat consent. Replace demo developer/privacy/terms URLs in the manifest before
wider distribution.

## Troubleshooting

| Symptom | Check |
| --- | --- |
| Authorization failure during provisioning | ATK accounts, selected subscription/group, Azure permissions, and Entra app-registration policy. |
| Custom identifier URI rejected | Ask the tenant administrator whether app policy permits the sign-in app's `api://botid-<BOT_ID>` URI. |
| Foundry 401/403 or model not found | Endpoint, exact deployment name, identity principal ID, project-scoped Foundry User assignment, and propagation time. |
| Group history fails | The correct environment's app must have consent in that chat. Its `webApplicationInfo.id` must match the generated sign-in app. |
| Repeated sign-in or admin approval prompt | Test `GraphDelegated`; verify the latest package and tenant delegated-consent policy. Never substitute application file permissions. |
| App fails to start | ATK build output and App Service logs; a 404 at `/` alone is not a startup failure. |
| Settings do not change | Provision the selected environment; deploying code alone does not update App Service settings. |

## Remove a test environment

Remove its Teams installations/catalog entry. Verify the selected app resource
group with `az resource list --resource-group "$APP_RG" -o table` before running
`az group delete --name "$APP_RG"` and confirming deletion.
Deleting only the web app leaves its paid App Service plan.

Remove the generated `-userauth` Entra app registration separately.
Do not delete shared Foundry resources; use the
[platform cleanup guide](../foundry-platform/README.md#remove-a-test-platform)
only for an unused, dedicated test platform.
