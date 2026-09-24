export const ERR_REFRESH_FAILED = "err/refresh/failed";
export const ERR_UNKNOWN = "err/unknown";

export const GRAPH_ENDPOINT = "https://graph.microsoft.com/v1.0";

export const MICROSOFT_OAUTH2_BASE_URL = "https://login.microsoftonline.com";
export const MICROSOFT_OAUTH2_DEFAULT_TENANT = "common";

/**
 * Scopes a vault storage integration typically needs: file read/write, plus
 * `offline_access` so a refresh token is issued
 */
export const ONEDRIVE_DEFAULT_SCOPES = ["Files.ReadWrite", "offline_access"];
