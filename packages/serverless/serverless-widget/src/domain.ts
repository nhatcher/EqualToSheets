let globalGraphqlEndpoint: string = "https://www.equalto.com/serverless/graphql";

export function setGraphqlEndpoint(endpoint: string) {
  globalGraphqlEndpoint = endpoint;
}

export function getGraphqlEndpoint(): string {
  return globalGraphqlEndpoint;
}

let globalSyncEndpoint: string = "https://www.equalto.com/serverless/get-updated-workbook";

export function setSyncEndpoint(endpoint: string) {
  globalSyncEndpoint = endpoint;
}

export function getSyncEndpoint(): string {
  return globalSyncEndpoint;
}


