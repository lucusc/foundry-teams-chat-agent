// Single-tenant Microsoft Entra app used for Teams SSO and delegated (on-behalf-of)
// OneDrive/SharePoint file reads. See docs/file-access-security.md.
//
// SECURITY: This app requests only DELEGATED Microsoft Graph permissions.
// Delegated Files.Read.All/Sites.Read.All let the bot read only files that the
// signed-in user can already open. The app deliberately has NO application
// (app-only) Files.Read.All or Sites.Read.All role, which would let this service
// read every file and site in the organization without any user present.
//
// The bot's own identity stays the user-assigned managed identity. This app adds:
// - api://botid-<bot id>/access_as_user, pre-authorized for Teams, so Teams can
//   issue an SSO token for the user without a prompt.
// - A federated identity credential trusted by the Azure Bot token service, which
//   exchanges the SSO token for a delegated Graph token (Microsoft Entra OBO)
//   through the bot's OAuth connection. No client secret is created.
// - A federated identity credential trusted by the App Service managed identity,
//   so the runtime can authenticate as this app (app-only) for the chat-scoped
//   RSC grant (ChatMessage.Read.Chat). Teams grants RSC to webApplicationInfo.id,
//   which is this app.

extension microsoftGraphV1

@description('Unique name and display name of the Microsoft Entra app')
param appName string

@description('Bot ID (managed identity client ID). Teams SSO for bots requires api://botid-<bot id>.')
param botId string

@description('Principal (object) ID of the App Service managed identity')
param managedIdentityPrincipalId string

@description('Unique identifier shared by the token-store federated credential and the Azure Bot OAuth connection')
param oauthConnectionUniqueId string

param tenantId string = tenant().tenantId

// Bot Service token store federated-credential subjects need the tenant GUID as
// base64url-encoded little-endian bytes without padding. This pure-Bicep encoder
// avoids a deployment script. For example, 72f988bf-86f1-41af-91ab-2d7cd011db47
// encodes to v4j5cvGGr0GRqy180BHbRw.
func base64UrlTenantId(guidValue string) string =>
  first(map(
    [replace(toLower(guidValue), '-', '')],
    hex =>
      first(map(
        [
          '${substring(hex, 6, 2)}${substring(hex, 4, 2)}${substring(hex, 2, 2)}${substring(hex, 0, 2)}${substring(hex, 10, 2)}${substring(hex, 8, 2)}${substring(hex, 14, 2)}${substring(hex, 12, 2)}${substring(hex, 16, 16)}'
        ],
        le =>
          join(
            map(
              range(0, 11),
              i =>
                i < 10
                  ? first(map(
                      [
                        indexOf('0123456789abcdef', substring(le, i * 3, 1)) * 256 + indexOf(
                          '0123456789abcdef',
                          substring(le, i * 3 + 1, 1)
                        ) * 16 + indexOf('0123456789abcdef', substring(le, i * 3 + 2, 1))
                      ],
                      v =>
                        '${substring('ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_', v / 64, 1)}${substring('ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_', max(v, 0) % 64, 1)}'
                    ))!
                  : first(map(
                      [
                        indexOf('0123456789abcdef', substring(le, 30, 1)) * 16 + indexOf(
                          '0123456789abcdef',
                          substring(le, 31, 1)
                        )
                      ],
                      v =>
                        '${substring('ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_', v / 4, 1)}${substring('ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_', (v % 4) * 16, 1)}'
                    ))!
            ),
            ''
          )
      ))!
  ))!

var graphAppId = '00000003-0000-0000-c000-000000000000'
var accessAsUserScopeId = guid(appName, 'access_as_user')
var identifierUri = 'api://botid-${botId}'
var issuer = '${environment().authentication.loginEndpoint}${tenantId}/v2.0'
// Base64url-encoded application ID of the Bot Service token store, from
// https://learn.microsoft.com/microsoft-365/agents-sdk/azure-bot-user-authorization-federated-credentials
var botServiceTokenStoreEncodedAppId = '9ExAW52n_ky4ZiS_jhpJIQ'
// Teams desktop/mobile and Teams web clients.
var teamsClientAppIds = [
  '1fec8e78-bce4-4aaf-ab1b-5451cc387264'
  '5e3ce6c0-2b1f-4285-8d4b-75ee78787346'
]

resource userAuthApp 'Microsoft.Graph/applications@v1.0' = {
  uniqueName: appName
  displayName: appName
  // Single tenant: only users in this organization can sign in.
  signInAudience: 'AzureADMyOrg'
  identifierUris: [
    identifierUri
  ]
  api: {
    requestedAccessTokenVersion: 2
    oauth2PermissionScopes: [
      {
        id: accessAsUserScopeId
        value: 'access_as_user'
        type: 'User'
        isEnabled: true
        adminConsentDisplayName: 'Use the Teams bot as the signed-in user'
        adminConsentDescription: 'Lets Teams sign the user in to the bot so it can read files attached to that user\'s messages.'
        userConsentDisplayName: 'Use the Teams bot as you'
        userConsentDescription: 'Lets the bot read files you attach to messages that mention it.'
      }
    ]
    preAuthorizedApplications: [
      for clientAppId in teamsClientAppIds: {
        appId: clientAppId
        delegatedPermissionIds: [
          accessAsUserScopeId
        ]
      }
    ]
  }
  web: {
    redirectUris: [
      'https://token.botframework.com/.auth/web/redirect'
    ]
    implicitGrantSettings: {
      enableAccessTokenIssuance: false
      enableIdTokenIssuance: false
    }
  }
  requiredResourceAccess: [
    {
      resourceAppId: graphAppId
      // Every entry is type 'Scope' (delegated). Never add type 'Role' (application)
      // Files.Read.All or Sites.Read.All here; that grants organization-wide access.
      resourceAccess: [
        {
          id: '37f7f235-527c-4136-accd-4a02d197296e' // openid
          type: 'Scope'
        }
        {
          id: '14dad69e-099b-42c9-810b-d002981feec1' // profile
          type: 'Scope'
        }
        {
          id: '7427e0e9-2fba-42fe-b0c0-848c9e6a8182' // offline_access
          type: 'Scope'
        }
        {
          id: 'df85f4d6-205c-4ac5-a5ea-6bf408dba283' // Files.Read.All (delegated)
          type: 'Scope'
        }
        {
          id: '205e70e5-aba6-4c52-a976-6d2d46c48043' // Sites.Read.All (delegated)
          type: 'Scope'
        }
      ]
    }
  ]
}

resource userAuthServicePrincipal 'Microsoft.Graph/servicePrincipals@v1.0' = {
  appId: userAuthApp.appId
}

// Trusted by the Azure Bot token service for the OAuth connection's SSO token
// exchange (Microsoft Entra on-behalf-of).
resource botTokenStoreCredential 'Microsoft.Graph/applications/federatedIdentityCredentials@v1.0' = {
  name: '${userAuthApp.uniqueName}/${oauthConnectionUniqueId}'
  audiences: [
    'api://AzureADTokenExchange'
  ]
  issuer: issuer
  subject: '/eid1/c/pub/t/${base64UrlTenantId(tenantId)}/a/${botServiceTokenStoreEncodedAppId}/${oauthConnectionUniqueId}'
  description: 'Azure Bot token service SSO token exchange'
}

// Trusted by the App Service managed identity for app-only RSC reads.
resource managedIdentityCredential 'Microsoft.Graph/applications/federatedIdentityCredentials@v1.0' = {
  name: '${userAuthApp.uniqueName}/bot-managed-identity'
  audiences: [
    'api://AzureADTokenExchange'
  ]
  issuer: issuer
  subject: managedIdentityPrincipalId
  description: 'App Service managed identity for chat-scoped RSC Graph reads'
}

output appId string = userAuthApp.appId
output identifierUri string = identifierUri
output servicePrincipalId string = userAuthServicePrincipal.id
output tokenStoreSubject string = botTokenStoreCredential.subject
