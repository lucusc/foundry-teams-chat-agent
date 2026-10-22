@description('Azure region for regional resources.')
param location string = resourceGroup().location

@description('AZD environment name used to derive globally unique resource names.')
param environmentName string = 'foundry-teams-chat-agent-poc'

@description('Foundry project name.')
param projectName string = 'foundry-teams-chat-agent-poc'

@description('Model deployment name and model name.')
param modelDeploymentName string = 'gpt-4.1-mini'

@description('Model version.')
param modelVersion string = '2025-04-14'

@description('Global Standard model capacity, in thousands of tokens per minute. A single request with a large PDF can exceed 10K tokens.')
param modelCapacity int = 100

var nameSuffix = uniqueString(subscription().id, resourceGroup().id)
var accountName = take('foundry${nameSuffix}', 24)
var resourceTags = {
  Environment: environmentName
  Project: 'foundry-teams-chat-agent-poc'
}

resource foundryAccount 'Microsoft.CognitiveServices/accounts@2025-06-01' = {
  name: accountName
  location: location
  kind: 'AIServices'
  identity: {
    type: 'SystemAssigned'
  }
  sku: {
    name: 'S0'
  }
  tags: resourceTags
  properties: {
    allowProjectManagement: true
    customSubDomainName: accountName
    publicNetworkAccess: 'Enabled'
    disableLocalAuth: true
  }
}

resource foundryProject 'Microsoft.CognitiveServices/accounts/projects@2025-06-01' = {
  parent: foundryAccount
  name: projectName
  location: location
  tags: resourceTags
  identity: {
    type: 'SystemAssigned'
  }
  properties: {
    displayName: 'Foundry Teams Chat Agent POC'
    description: 'Microsoft Foundry project for the Teams chat agent proof of concept.'
  }
}

resource modelDeployment 'Microsoft.CognitiveServices/accounts/deployments@2025-06-01' = {
  parent: foundryAccount
  name: modelDeploymentName
  dependsOn: [
    foundryProject
  ]
  sku: {
    name: 'GlobalStandard'
    capacity: modelCapacity
  }
  properties: {
    model: {
      format: 'OpenAI'
      name: modelDeploymentName
      version: modelVersion
    }
  }
}

output AZURE_LOCATION string = location
output AZURE_AI_PROJECT_ENDPOINT string = foundryProject.properties.endpoints['AI Foundry API']
output FOUNDRY_PROJECT_ENDPOINT string = foundryProject.properties.endpoints['AI Foundry API']
output AZURE_AI_PROJECT_ID string = foundryProject.id
output AZURE_AI_FOUNDRY_PROJECT_ID string = foundryProject.id
output AZURE_AI_ACCOUNT_NAME string = foundryAccount.name
output AZURE_AI_PROJECT_NAME string = foundryProject.name
output AZURE_AI_MODEL_DEPLOYMENT_NAME string = modelDeployment.name
