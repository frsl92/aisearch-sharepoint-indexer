param location string
param foundryLocation string
param resourceToken string
param tags object
param principalId string
param principalType string
param searchSku string
param searchReplicaCount int
param searchPartitionCount int
param searchSemanticSearch string
param embeddingDeploymentName string
param embeddingModelVersion string
param embeddingSku string
param embeddingCapacity int

resource search 'Microsoft.Search/searchServices@2026-09-01-preview' = {
  name: 'search-${resourceToken}'
  location: location
  tags: tags
  sku: {
    name: searchSku
  }
  identity: {
    type: 'SystemAssigned'
  }
  properties: union({
    semanticSearch: searchSemanticSearch
    publicNetworkAccess: 'Enabled'
    disableLocalAuth: true
  }, searchSku == 'serverless' ? {} : {
    replicaCount: searchReplicaCount
    partitionCount: searchPartitionCount
    hostingMode: 'Default'
  })
}

resource foundry 'Microsoft.CognitiveServices/accounts@2025-06-01' = {
  name: 'foundry-${resourceToken}'
  location: foundryLocation
  tags: tags
  // Multi-service account for OpenAI, OCR and Vision (including image embeddings).
  kind: 'AIServices'
  sku: {
    name: 'S0'
  }
  identity: {
    type: 'SystemAssigned'
  }
  properties: {
    customSubDomainName: 'foundry-${resourceToken}'
    allowProjectManagement: true
    publicNetworkAccess: 'Enabled'
    disableLocalAuth: true
  }
}

resource embedding 'Microsoft.CognitiveServices/accounts/deployments@2025-06-01' = {
  parent: foundry
  name: embeddingDeploymentName
  sku: {
    name: embeddingSku
    capacity: embeddingCapacity
  }
  properties: {
    model: {
      format: 'OpenAI'
      name: 'text-embedding-3-large'
      version: embeddingModelVersion
    }
    versionUpgradeOption: 'NoAutoUpgrade'
  }
}

var searchContributorRole = subscriptionResourceId(
  'Microsoft.Authorization/roleDefinitions', '7ca78c08-252a-4471-8644-bb5ff32d4ba0'
)
var openaiUserRole = subscriptionResourceId(
  'Microsoft.Authorization/roleDefinitions', '5e0bd9bd-7b93-4f28-af87-19fc36ad61bd'
)
var cognitiveServicesUserRole = subscriptionResourceId(
  'Microsoft.Authorization/roleDefinitions', 'a97b65f3-24c7-4388-baec-2e87135dc908'
)

resource searchContributor 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(search.id, principalId, searchContributorRole)
  scope: search
  properties: {
    roleDefinitionId: searchContributorRole
    principalId: principalId
    principalType: principalType
  }
}

resource searchOpenaiUser 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(foundry.id, search.id, openaiUserRole)
  scope: foundry
  properties: {
    roleDefinitionId: openaiUserRole
    principalId: search.identity.principalId
    principalType: 'ServicePrincipal'
  }
}

resource searchCognitiveServicesUser 'Microsoft.Authorization/roleAssignments@2022-04-01' = {
  name: guid(foundry.id, search.id, cognitiveServicesUserRole)
  scope: foundry
  properties: {
    roleDefinitionId: cognitiveServicesUserRole
    principalId: search.identity.principalId
    principalType: 'ServicePrincipal'
  }
}

output searchName string = search.name
output searchEndpoint string = search.properties.endpoint
output searchPrincipalId string = search.identity.principalId
output foundryName string = foundry.name
output openaiEndpoint string = 'https://${foundry.name}.openai.azure.com'
output aiServicesEndpoint string = foundry.properties.endpoint
