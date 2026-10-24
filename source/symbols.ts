export const ERR_REFRESH_FAILED = "err/refresh/failed";
export const ERR_UNKNOWN = "err/unknown";

export const GRAPH_ENDPOINT = "https://graph.microsoft.com/v1.0";

/**
 * Hosts a Graph endpoint may use without opting in (global and national
 * clouds)
 */
export const GRAPH_HOSTS = [
    "graph.microsoft.com",
    "graph.microsoft.us",
    "dod-graph.microsoft.us",
    "microsoftgraph.chinacloudapi.cn"
];

/**
 * Stop following `@odata.nextLink` after this many pages
 */
export const MAX_DIRECTORY_PAGES = 1000;

export const MICROSOFT_OAUTH2_BASE_URL = "https://login.microsoftonline.com";
export const MICROSOFT_OAUTH2_DEFAULT_TENANT = "common";

/**
 * Scopes a vault storage integration typically needs: file read/write, plus
 * `offline_access` so a refresh token is issued
 */
export const ONEDRIVE_DEFAULT_SCOPES = ["Files.ReadWrite", "offline_access"];
