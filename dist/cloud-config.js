const CLOUD_ENDPOINTS = {
  global: {
    authority: "https://login.microsoftonline.com",
    graphApi: "https://graph.microsoft.com",
    portal: "https://portal.azure.com"
  },
  china: {
    authority: "https://login.chinacloudapi.cn",
    graphApi: "https://microsoftgraph.chinacloudapi.cn",
    portal: "https://portal.azure.cn"
  }
};
const DEFAULT_CLIENT_IDS = {
  global: "084a3e9f-a9f4-43f7-89f9-d229cf97853e",
  china: "f3e61a6e-bc26-4281-8588-2c7359a02141"
};
function getDefaultClientId(cloudType = "global") {
  return DEFAULT_CLIENT_IDS[cloudType];
}
function getCloudEndpoints(cloudType = "global") {
  const endpoints = CLOUD_ENDPOINTS[cloudType];
  if (!endpoints) {
    throw new Error(
      `Unknown cloud type: ${cloudType}. Valid values: ${Object.keys(CLOUD_ENDPOINTS).join(", ")}`
    );
  }
  return endpoints;
}
function isValidCloudType(value) {
  return value in CLOUD_ENDPOINTS;
}
function parseCloudType(value) {
  if (!value) return "global";
  const normalized = value.toLowerCase().trim();
  if (!isValidCloudType(normalized)) {
    throw new Error(
      `Invalid cloud type: ${value}. Valid values: ${Object.keys(CLOUD_ENDPOINTS).join(", ")}`
    );
  }
  return normalized;
}
export {
  CLOUD_ENDPOINTS,
  DEFAULT_CLIENT_IDS,
  getCloudEndpoints,
  getDefaultClientId,
  isValidCloudType,
  parseCloudType
};
