@maxLength(20)
@minLength(4)
@description('Used to generate names for all resources in this file')
param resourceBaseName string

@description('Microsoft Foundry project endpoint used by the agent')
param azureAiProjectEndpoint string
@description('Model deployment used to classify requests for agent routing')
param azureAiModelDeploymentName string
@description('Existing Foundry prompt agent for general questions')
param azureAiGeneralAgentName string = 'teams-poc-general'
@description('Existing Foundry prompt agent for summaries')
param azureAiSummaryAgentName string = 'teams-poc-summary'
param graphHistoryMessageCount string = '10'
@description('SharePoint tenant prefix, for example contoso')
param sharePointTenantDomain string = ''
@description('Maximum allowed attached file size in bytes')
param maxFileSizeBytes int = 20971520

param webAppSKU string
param linuxFxVersion string

@maxLength(42)
param botDisplayName string

param serverfarmsName string = resourceBaseName
param webAppName string = resourceBaseName
param identityName string = resourceBaseName
param location string = resourceGroup().location
param pythonVersion string = linuxFxVersion

resource identity 'Microsoft.ManagedIdentity/userAssignedIdentities@2023-01-31' = {
  location: location
  name: identityName
}

// SECURITY: Do not grant the managed identity Microsoft Graph application roles
// such as Files.Read.All or Sites.Read.All. Those are organization-wide: they
// would let this App Service read every OneDrive and SharePoint file in the
// tenant. Files shared in a chat are read with the sending user's delegated token
// (Teams SSO + on-behalf-of) instead. See docs/file-access-security.md (ADR-001).
var oauthConnectionName = 'GraphDelegated'
var oauthConnectionUniqueId = guid(resourceBaseName, 'BotServiceOauthConnection')

module userAuthApp './userAuth/userAuthApp.bicep' = {
  name: 'User-auth-app'
  params: {
    appName: '${resourceBaseName}-userauth'
    botId: identity.properties.clientId
    managedIdentityPrincipalId: identity.properties.principalId
    oauthConnectionUniqueId: oauthConnectionUniqueId
  }
}

// Compute resources for your Web App
resource serverfarm 'Microsoft.Web/serverfarms@2021-02-01' = {
  kind: 'app,linux'
  location: location
  name: serverfarmsName
  sku: {
    name: webAppSKU
  }
  properties:{
    reserved: true
  }
}

// Web App that hosts your agent
resource webApp 'Microsoft.Web/sites@2021-02-01' = {
  kind: 'app,linux'
  location: location
  name: webAppName
  properties: {
    serverFarmId: serverfarm.id
    siteConfig: {
      alwaysOn: true
      appCommandLine: 'gunicorn --bind 0.0.0.0 --worker-class aiohttp.worker.GunicornWebWorker --timeout 600 app:app'
      linuxFxVersion: pythonVersion
      appSettings: [
        {
          name: 'WEBSITES_CONTAINER_START_TIME_LIMIT'
          value: '600'
        }
        {
          name: 'SCM_DO_BUILD_DURING_DEPLOYMENT'
          value: 'true'
        }
        {
          name: 'BOT_ID'
          value: identity.properties.clientId
        }
        {
          name: 'CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTID'
          value: identity.properties.clientId
        }
        {
          name: 'CONNECTIONS__SERVICE_CONNECTION__SETTINGS__AUTHTYPE'
          value: 'UserManagedIdentity'
        }
        {
          name: 'CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID'
          value: identity.properties.tenantId
        }
        {
          name: 'AZURE_CLIENT_ID'
          value: identity.properties.clientId
        }
        {
          name: 'AZURE_AI_PROJECT_ENDPOINT'
          value: azureAiProjectEndpoint
        }
        {
          name: 'AZURE_AI_MODEL_DEPLOYMENT_NAME'
          value: azureAiModelDeploymentName
        }
        {
          name: 'AZURE_AI_GENERAL_AGENT_NAME'
          value: azureAiGeneralAgentName
        }
        {
          name: 'AZURE_AI_SUMMARY_AGENT_NAME'
          value: azureAiSummaryAgentName
        }
        {
          name: 'GRAPH_HISTORY_MESSAGE_COUNT'
          value: graphHistoryMessageCount
        }
        {
          name: 'SHAREPOINT_TENANT_DOMAIN'
          value: sharePointTenantDomain
        }
        {
          name: 'MAX_FILE_SIZE_BYTES'
          value: string(maxFileSizeBytes)
        }
        {
          name: 'BOT_TENANT_ID'
          value: identity.properties.tenantId
        }
        { 
          name: 'BOT_TYPE'
          value: 'UserAssignedMsi' 
        }
        {
          // App-only Graph calls use the user-auth app (through the managed
          // identity's federated credential) because Teams grants chat-scoped RSC
          // to webApplicationInfo.id.
          name: 'GRAPH_AUTH_MODE'
          value: 'federated'
        }
        {
          name: 'SSO_APP_ID'
          value: userAuthApp.outputs.appId
        }
        {
          // Enables the Agents SDK user-authorization handler named GRAPH, which
          // gets the delegated token from the Azure Bot OAuth connection.
          name: 'AGENTAPPLICATION__USERAUTHORIZATION__HANDLERS__GRAPH__SETTINGS__AZUREBOTOAUTHCONNECTIONNAME'
          value: oauthConnectionName
        }
        {
          name: 'AGENTAPPLICATION__USERAUTHORIZATION__HANDLERS__GRAPH__SETTINGS__TITLE'
          value: 'Sign in to share files'
        }
        {
          name: 'AGENTAPPLICATION__USERAUTHORIZATION__HANDLERS__GRAPH__SETTINGS__TEXT'
          value: 'Sign in so the agent can read the files you attached, using your own access.'
        }
      ]
      ftpsState: 'FtpsOnly'
    }
  }
  identity: {
    type: 'UserAssigned'
    userAssignedIdentities: {
      '${identity.id}': {}
    }
  }
}

// Register your web service as a bot with the Bot Framework
module azureBotRegistration './botRegistration/azurebot.bicep' = {
  name: 'Azure-Bot-registration'
  params: {
    resourceBaseName: resourceBaseName
    identityClientId: identity.properties.clientId
    identityResourceId: identity.id
    identityTenantId: identity.properties.tenantId
    botAppDomain: webApp.properties.defaultHostName
    botDisplayName: botDisplayName
    userAuthAppId: userAuthApp.outputs.appId
    userAuthIdentifierUri: userAuthApp.outputs.identifierUri
    oauthConnectionName: oauthConnectionName
    oauthConnectionUniqueId: oauthConnectionUniqueId
  }
}

// The output will be persisted in .env.{envName}. Visit https://aka.ms/teamsfx-actions/arm-deploy for more details.
output BOT_AZURE_APP_SERVICE_RESOURCE_ID string = webApp.id
output BOT_DOMAIN string = webApp.properties.defaultHostName
output BOT_ID string = identity.properties.clientId
output BOT_TENANT_ID string = identity.properties.tenantId
output CONNECTIONS__SERVICE_CONNECTION__SETTINGS__CLIENTID string = identity.properties.clientId
output CONNECTIONS__SERVICE_CONNECTION__SETTINGS__AUTHTYPE string = 'UserManagedIdentity'
output CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID string = identity.properties.tenantId
output SSO_APP_ID string = userAuthApp.outputs.appId
output SSO_APP_ID_URI string = userAuthApp.outputs.identifierUri
