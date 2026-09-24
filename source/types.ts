export type Headers = Record<string, string>;

/**
 * A OneDrive file or folder, referenced either by its drive item ID or by
 * its path relative to the root of the user's drive. `fileID` takes
 * precedence when both are set.
 */
export interface OneDriveTarget {
    fileID?: string;
    path?: string;
}

export interface OneDriveClientOptions {
    /**
     * Microsoft Graph API base URL - only needed for national clouds.
     * Defaults to "https://graph.microsoft.com/v1.0".
     */
    endpoint?: string;
}

export interface OneDriveItem {
    id: string;
    name: string;
    /**
     * Path of the item relative to the drive root, eg. "/Vaults/my.bcup"
     */
    path: string;
    size: number;
    type: "directory" | "file";
    mtime: string | null;
}

export interface MicrosoftToken {
    access_token: string;
    refresh_token?: string;
    expires_in?: number;
    /**
     * Expiry as a JS timestamp, computed from `expires_in` on receipt
     */
    expiry_date?: number;
    scope?: string;
    token_type?: string;
}

export interface MicrosoftTokenResponse {
    res: any;
    tokens: MicrosoftToken;
}

export interface GenerateAuthUrlOptions {
    scope?: Array<string> | string;
    state?: string;
    prompt?: "login" | "none" | "consent" | "select_account";
    /**
     * PKCE code challenge (base64url SHA-256 of the verifier). Recommended
     * for public clients (desktop/mobile/browser) - generate the pair in the
     * application with its platform's crypto.
     */
    codeChallenge?: string;
    codeChallengeMethod?: "S256" | "plain";
}

export interface OneDriveOAuth2ClientOptions {
    /**
     * Only for confidential (server-side) clients - public clients should
     * use PKCE instead
     */
    clientSecret?: string;
    /**
     * Tenant: "common" (default - personal and work/school accounts),
     * "consumers", "organizations" or a tenant ID
     */
    tenant?: string;
}
