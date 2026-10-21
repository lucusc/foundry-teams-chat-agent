# Shared Foundry Platform

This infrastructure-only AZD project owns the shared Microsoft Foundry account,
project, and model deployment. The Teams application remains managed by
Microsoft 365 Agents Toolkit at the repository root.

This guide covers only the model infrastructure. To run the bot, use
[Agents Playground](../docs/agents-playground.md),
[local Teams debugging](../docs/local-teams-debugging.md), or
[deploy a dev or staging environment](../docs/deploy-environment.md).

## Provision

Install [Azure Developer CLI](https://learn.microsoft.com/en-us/azure/developer/azure-developer-cli/install-azd)
and Azure CLI. You need resource creation permissions, permission to register
providers, and quota for `gpt-4.1-mini` version `2025-04-14` with Global Standard
capacity 10 in the selected region. Provisioning incurs Azure charges.

Run these commands from the repository root. Enter your actual tenant,
subscription, dedicated platform resource group, region, and AZD environment
name (independent of the ATK dev/staging environment):

```bash
read -r -p "Azure tenant ID: " TENANT_ID
read -r -p "Azure subscription ID: " SUBSCRIPTION_ID
read -r -p "Foundry resource group: " FOUNDRY_RG
read -r -p "Azure region: " LOCATION
read -r -p "Foundry AZD environment name: " FOUNDRY_ENV

az login --tenant "$TENANT_ID"
az account set --subscription "$SUBSCRIPTION_ID"
az provider register --namespace Microsoft.CognitiveServices --wait
az group create --name "$FOUNDRY_RG" --location "$LOCATION"

cd foundry-platform
azd auth login --tenant-id "$TENANT_ID"
azd env new "$FOUNDRY_ENV" \
  --subscription "$SUBSCRIPTION_ID" --location "$LOCATION"
azd env set AZURE_RESOURCE_GROUP "$FOUNDRY_RG"
azd provision --no-prompt
```

Run all platform AZD commands inside `foundry-platform`, not the repository
root. The configuration has no application services, hosted-agent extensions,
or Teams publishing hooks; `azd deploy` is not needed.

The [parameter file](infra/main.parameters.json) binds the AZD environment
name and location to the existing resource-group-scoped
[Bicep template](infra/main.bicep). Project and model defaults are unchanged.
To customize them, update the Bicep defaults or add explicit values for
`projectName`, `modelDeploymentName`, `modelVersion`, or `modelCapacity` in the
parameter file before provisioning.

## Read outputs

From `foundry-platform`:

```bash
azd env get-value AZURE_AI_PROJECT_ENDPOINT
azd env get-value AZURE_AI_MODEL_DEPLOYMENT_NAME
azd env get-value AZURE_AI_PROJECT_ID
```

Copy the endpoint and model deployment name into the ATK environment files at
the repository root. Use the project resource ID when assigning Foundry access
to the developer and the bot's managed identity, as described in the
[Playground access setup](../docs/agents-playground.md#sign-in-and-grant-access)
and [cloud bot access setup](../docs/deploy-environment.md#3-provision-grant-access-and-deploy-code).
Provisioning does not grant these callers access automatically.

AZD saves environment settings and deployment outputs under
`foundry-platform/.azure/`, which is ignored by Git. It does not update the
ATK environment files automatically.

## Update or adopt an existing platform

- For an environment already created in this folder, run `azd env select`
  with its environment name, then `azd provision --no-prompt`.
- For resources provisioned before this folder became an AZD project, create
  an AZD environment here and set the **same subscription, resource group,
  location, and project/model parameters** before provisioning. The account
  name is derived from the subscription and resource group; changing either
  creates a different account. Changing `projectName` creates another project.
- Keep existing environment names when possible to preserve their tag values.
  Root-level AZD environment state from the former hosted-agent implementation
  is not used by this separate project. Do not run its old deployment hooks.

After changing infrastructure, run `azd provision --no-prompt` again.
Refresh the ATK settings and reprovision the Teams app if the project endpoint
or model deployment name changes.

If provisioning fails, check the selected AZD environment with
`azd env get-values` and verify regional model availability and quota.
Do not change an existing environment's resource group or project/model names
unless you intend to create different resources.

## Remove a test platform

From `foundry-platform`, verify the selected environment and resource group:

```bash
azd env get-values
azd down
```

Only confirm deletion for a dedicated test platform that no other application
uses. The Teams app has a separate ATK lifecycle and resource group.
