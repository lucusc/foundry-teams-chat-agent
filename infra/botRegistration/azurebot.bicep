@maxLength(20)
@minLength(4)
@description('Used to generate names for all resources in this file')
param resourceBaseName string

@maxLength(42)
param botDisplayName string

param botServiceName string = resourceBaseName
param botServiceSku string = 'F0'
param identityResourceId string
param identityClientId string
param identityTenantId string
param botAppDomain string

@description('Client ID of the single-tenant Microsoft Entra app used for Teams SSO')
param userAuthAppId string
@description('Application ID URI that Teams SSO tokens are issued for, api://botid-<bot id>')
param userAuthIdentifierUri string
@description('Name of the Azure Bot OAuth connection that the agent code uses')
param oauthConnectionName string
@description('Unique identifier that matches the Bot Service federated identity credential subject')
param oauthConnectionUniqueId string

// Register your web service as a bot with the Bot Framework
resource botService 'Microsoft.BotService/botServices@2021-03-01' = {
  kind: 'azurebot'
  location: 'global'
  name: botServiceName
  properties: {
    displayName: botDisplayName
    endpoint: 'https://${botAppDomain}/api/messages'
    msaAppId: identityClientId
    msaAppMSIResourceId: identityResourceId
    msaAppTenantId:identityTenantId
    msaAppType:'UserAssignedMSI'
  }
  sku: {
    name: botServiceSku
  }
}

// Connect the bot service to Microsoft Teams
resource botServiceMsTeamsChannel 'Microsoft.BotService/botServices/channels@2021-03-01' = {
  parent: botService
  location: 'global'
  name: 'MsTeamsChannel'
  properties: {
    channelName: 'MsTeamsChannel'
  }
}

// OAuth connection used for Teams SSO. The Azure Bot token service exchanges the
// user's Teams SSO token for a DELEGATED Microsoft Graph token (Microsoft Entra
// on-behalf-of), authenticating as the user-auth app with a federated identity
// credential instead of a client secret. Delegated scopes bound file reads to
// what the signed-in user can already open. See docs/file-access-security.md.
resource graphOAuthConnection 'Microsoft.BotService/botServices/connections@2022-09-15' = {
  parent: botService
  name: oauthConnectionName
  location: 'global'
  properties: {
    serviceProviderId: 'c00b44ab-5e16-c44c-af26-2fd5bc55eb18'
    serviceProviderDisplayName: 'AAD v2 with Federated Credentials'
    clientId: userAuthAppId
    scopes: 'openid profile offline_access Files.Read.All Sites.Read.All'
    parameters: [
      {
        key: 'ClientId'
        value: userAuthAppId
      }
      {
        key: 'UniqueIdentifier'
        value: oauthConnectionUniqueId
      }
      {
        key: 'TokenExchangeUrl'
        value: userAuthIdentifierUri
      }
      {
        key: 'TenantId'
        value: identityTenantId
      }
    ]
  }
}
