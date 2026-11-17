targetScope = 'subscription'

@minLength(1)
@maxLength(64)
@description('azd environment name. Use a distinct environment for dev, qa and prod.')
param environmentName string

param location string
param resourceGroupName string = ''

@description('Object ID of the azd user or service principal that configures Search.')
@minLength(36)
@maxLength(36)
param principalId string

@allowed(['User', 'ServicePrincipal'])
param principalType string = 'User'

@allowed(['basic', 'standard', 'standard2', 'standard3', 'serverless'])
param searchSku string = 'basic'

@minValue(1)
@maxValue(12)
param searchReplicaCount int = 1

@allowed([1, 2, 3, 4, 6, 12])
param searchPartitionCount int = 1

@allowed(['disabled', 'free', 'standard'])
param searchSemanticSearch string = 'free'

@description('May differ from Search location. Verify model and Vision availability first.')
param foundryLocation string = ''

param embeddingDeploymentName string = 'text-embedding-3-large'
param embeddingModelVersion string = '1'

@description('Deployment SKU must be available for this model in the chosen region.')
param embeddingSku string = 'Standard'

@minValue(1)
@description('Model quota capacity units; not a raw tokens-per-minute count.')
param embeddingCapacity int = 1

@allowed(['no_embedding', 'image_embeddings', 'text_embeddings', 'all_embeddings'])
param deploymentType string = 'text_embeddings'

var tags = {
  'azd-env-name': environmentName
  environment: environmentName
}

resource resourceGroup 'Microsoft.Resources/resourceGroups@2025-04-01' = {
  name: empty(resourceGroupName) ? 'rg-${environmentName}' : resourceGroupName
  location: location
  tags: tags
}

module resources './resources.bicep' = {
  name: 'search-foundry'
  scope: resourceGroup
  params: {
    location: location
    foundryLocation: empty(foundryLocation) ? location : foundryLocation
    resourceToken: uniqueString(subscription().subscriptionId, resourceGroup.id)
    tags: tags
    principalId: principalId
    principalType: principalType
    searchSku: searchSku
    searchReplicaCount: searchReplicaCount
    searchPartitionCount: searchPartitionCount
    searchSemanticSearch: searchSemanticSearch
    embeddingDeploymentName: embeddingDeploymentName
    embeddingModelVersion: embeddingModelVersion
    embeddingSku: embeddingSku
    embeddingCapacity: embeddingCapacity
  }
}

output AZURE_RESOURCE_GROUP string = resourceGroup.name
output AZURE_TENANT_ID string = tenant().tenantId
output AZURE_SEARCH_SERVICE_NAME string = resources.outputs.searchName
output SEARCH_ENDPOINT string = resources.outputs.searchEndpoint
output SEARCH_IDENTITY_PRINCIPAL_ID string = resources.outputs.searchPrincipalId
output AZURE_AI_ACCOUNT_NAME string = resources.outputs.foundryName
output AZURE_OPENAI_ENDPOINT string = resources.outputs.openaiEndpoint
output AZURE_AI_SERVICES_ENDPOINT string = resources.outputs.aiServicesEndpoint
output AZURE_OPENAI_DEPLOYMENT string = embeddingDeploymentName
output DEPLOYMENT_TYPE string = deploymentType
