import { EventEmitter } from "eventemitter3";
import { Layerr } from "layerr";
import { fetch } from "@buttercup/fetch";
import { encodeForm, handleBadResponse, timeoutSignal } from "./request.js";
import {
    ERR_BAD_TOKEN_RESPONSE,
    ERR_REFRESH_FAILED,
    MICROSOFT_OAUTH2_BASE_URL,
    MICROSOFT_OAUTH2_DEFAULT_TENANT,
    ONEDRIVE_DEFAULT_SCOPES
} from "./symbols.js";
import {
    GenerateAuthUrlOptions,
    MicrosoftToken,
    MicrosoftTokenResponse,
    OneDriveOAuth2ClientOptions
} from "./types.js";

/**
 * OAuth2 client for the Microsoft identity platform (v2.0 endpoints), for
 * obtaining and refreshing OneDrive access tokens. Mirrors the API of
 * `@buttercup/google-oauth2-client`. Emits "tokens" whenever new tokens are
 * received.
 */
export class OAuth2Client extends EventEmitter {
    protected _accessToken: string = null;
    protected _clientID: string;
    protected _clientSecret: string | undefined;
    protected _fetch: typeof fetch;
    protected _redirectURL: string;
    protected _refreshToken: string = null;
    protected _refreshTokenPromises: Map<string, Promise<MicrosoftTokenResponse>>;
    protected _tenant: string;
    protected _timeout: number | undefined;

    constructor(clientID: string, redirectURL: string, options: OneDriveOAuth2ClientOptions = {}) {
        super();
        const { clientSecret, tenant = MICROSOFT_OAUTH2_DEFAULT_TENANT, timeout } = options;
        this._timeout = timeout;
        this._clientID = clientID;
        this._clientSecret = clientSecret;
        this._redirectURL = redirectURL;
        this._tenant = tenant;
        this._refreshTokenPromises = new Map();
        this._fetch = fetch;
    }

    get accessToken() {
        return this._accessToken;
    }

    get refreshToken() {
        return this._refreshToken;
    }

    get authorizeURL() {
        return `${MICROSOFT_OAUTH2_BASE_URL}/${this._tenant}/oauth2/v2.0/authorize`;
    }

    get tokenURL() {
        return `${MICROSOFT_OAUTH2_BASE_URL}/${this._tenant}/oauth2/v2.0/token`;
    }

    /**
     * Generate the URL to send the user to for signing in
     * @param options Scope, state, prompt and PKCE options
     */
    generateAuthUrl(options: GenerateAuthUrlOptions = {}): string {
        const {
            codeChallenge,
            codeChallengeMethod = "S256",
            prompt,
            scope: rawScope = ONEDRIVE_DEFAULT_SCOPES,
            state
        } = options;
        const scope = Array.isArray(rawScope) ? rawScope.join(" ") : rawScope;
        if (!state) {
            throw new Layerr("A state value is required: compare it with the one on the redirect");
        }
        if (!codeChallenge && !this._clientSecret) {
            throw new Layerr("A PKCE code challenge is required for a client without a secret");
        }
        if (codeChallengeMethod !== "S256") {
            throw new Layerr("Only the S256 code challenge method is supported");
        }
        return `${this.authorizeURL}?${encodeForm({
            client_id: this._clientID,
            response_type: "code",
            redirect_uri: this._redirectURL,
            response_mode: "query",
            scope,
            state,
            prompt,
            code_challenge: codeChallenge,
            code_challenge_method: codeChallenge ? codeChallengeMethod : undefined
        })}`;
    }

    /**
     * Exchange the authorisation code from the redirect for tokens
     * @param authCode The code from the redirect URL, already URL-decoded
     *  (as `URLSearchParams.get` returns it)
     * @param codeVerifier The PKCE verifier, if a challenge was used
     */
    async exchangeAuthCodeForToken(
        authCode: string,
        codeVerifier?: string
    ): Promise<MicrosoftTokenResponse> {
        const res = await this._requestToken({
            grant_type: "authorization_code",
            code: authCode,
            redirect_uri: this._redirectURL,
            code_verifier: codeVerifier
        });
        handleBadResponse(res);
        return this._receiveTokens(res);
    }

    /**
     * Refresh the access token. Concurrent calls for the same refresh token
     * share one request.
     * @param refreshToken The refresh token
     */
    refreshAccessToken(refreshToken: string): Promise<MicrosoftTokenResponse> {
        if (this._refreshTokenPromises.has(refreshToken)) {
            return this._refreshTokenPromises.get(refreshToken);
        }
        const promise = this.refreshAccessTokenNoCache(refreshToken).finally(() => {
            this._refreshTokenPromises.delete(refreshToken);
        });
        this._refreshTokenPromises.set(refreshToken, promise);
        return promise;
    }

    /**
     * Refresh the access token, without sharing concurrent requests
     * @param refreshToken The refresh token
     */
    async refreshAccessTokenNoCache(refreshToken: string): Promise<MicrosoftTokenResponse> {
        const res = await this._requestToken({
            grant_type: "refresh_token",
            refresh_token: refreshToken
        });
        if (res.status >= 400 && res.status < 500) {
            // eg. revoked, expired (inactivity) or consent withdrawn - the
            // user needs to sign in again
            throw new Layerr(
                {
                    info: {
                        code: ERR_REFRESH_FAILED,
                        status: res.status,
                        statusText: res.statusText
                    }
                },
                "Bad refresh response"
            );
        }
        handleBadResponse(res);
        return this._receiveTokens(res, refreshToken);
    }

    protected async _requestToken(data: Record<string, string | undefined>) {
        const { signal, clear } = timeoutSignal(this._timeout);
        try {
            return await this._fetch(this.tokenURL, {
                method: "POST",
                body: encodeForm({
                    client_id: this._clientID,
                    client_secret: this._clientSecret,
                    ...data
                }),
                headers: {
                    "Content-Type": "application/x-www-form-urlencoded"
                },
                signal
            });
        } finally {
            clear();
        }
    }

    protected async _receiveTokens(
        res: any,
        previousRefreshToken?: string
    ): Promise<MicrosoftTokenResponse> {
        const tokens = (await res.json()) as MicrosoftToken;
        if (
            !tokens ||
            typeof tokens.access_token !== "string" ||
            tokens.access_token.length === 0
        ) {
            throw new Layerr(
                { info: { code: ERR_BAD_TOKEN_RESPONSE, status: res.status } },
                "Token response has no access token"
            );
        }
        if (tokens.expires_in) {
            tokens.expiry_date = Date.now() + tokens.expires_in * 1000;
            delete tokens.expires_in;
        }
        // Microsoft usually rotates refresh tokens, but keep the old one if a
        // response omits it
        if (!tokens.refresh_token && previousRefreshToken) {
            tokens.refresh_token = previousRefreshToken;
        }
        this._accessToken = tokens.access_token;
        this._refreshToken = tokens.refresh_token || null;
        this.emit("tokens", tokens);
        return { tokens, res };
    }
}
