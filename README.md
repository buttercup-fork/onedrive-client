# OneDrive Client
> Basic client and OAuth2 helper for Microsoft OneDrive

## About

A small client for the parts of the [Microsoft Graph](https://learn.microsoft.com/graph/api/resources/onedrive) OneDrive API that Buttercup needs: **reading**, **writing**, **deleting** and **listing** files. It also includes a minimal **OAuth2 client** for the Microsoft identity platform, for obtaining and refreshing access tokens.

Like Buttercup's [Dropbox](https://github.com/buttercup/dropbox-client) and [Google Drive](https://github.com/buttercup/googledrive-client) clients, it makes plain HTTP requests rather than using Microsoft's SDKs. This keeps it small, easy to follow and working the same in NodeJS, the browser and React Native. Requests use [`@buttercup/fetch`](https://github.com/buttercup/fetch).

Only text files are supported for reading and writing.

## Installation

This fork's packages aren't published to npm. Install from git:

```shell
npm install github:buttercup-fork/onedrive-client --save
```

The package builds itself on install (`prepare`). It requires an [ESM](https://nodejs.org/api/esm.html) environment and NodeJS 22 or newer.

## Usage

### Signing in

Register an application in the [Azure portal](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade) ("App registrations"), add a redirect URI, and grant the delegated Microsoft Graph permission `Files.ReadWrite` (or `Files.ReadWrite.AppFolder` to limit access to an app folder).

```typescript
import { OAuth2Client } from "@buttercup/onedrive-client";

const oauth = new OAuth2Client(CLIENT_ID, "http://localhost:8123/callback");

// 1. Send the user to the sign-in page. Public clients (desktop, mobile,
//    browser) should use PKCE: generate a random verifier and pass its
//    base64url SHA-256 as the challenge.
const url = oauth.generateAuthUrl({ codeChallenge });

// 2. Exchange the `code` from the redirect for tokens
const { tokens } = await oauth.exchangeAuthCodeForToken(code, codeVerifier);
// tokens.access_token, tokens.refresh_token, tokens.expiry_date

// 3. Later, refresh an expired access token
const { tokens: fresh } = await oauth.refreshAccessToken(tokens.refresh_token);
```

`OAuth2Client` emits `"tokens"` whenever it receives new tokens. Options: `tenant` (`"common"` by default, which allows personal and work/school accounts; also `"consumers"`, `"organizations"` or a tenant ID), and `clientSecret` for confidential server-side apps only.

A refresh that fails with a 4xx response (revoked or expired refresh token) throws an error with `code: ERR_REFRESH_FAILED` in its [Layerr](https://github.com/perry-mitchell/layerr) info. The user needs to sign in again.

### Files

```typescript
import { OneDriveClient } from "@buttercup/onedrive-client";

const client = new OneDriveClient(accessToken);

// Items are referenced by drive item ID or by path from the drive root
const contents = await client.getFileContents({ path: "/Vaults/my.bcup" });
const itemID = await client.putFileContents({ path: "/Vaults/my.bcup" }, contents); // creates if missing
await client.getFileContents({ fileID: itemID });
await client.deleteFile({ fileID: itemID });

const items = await client.getDirectoryContents({ path: "/Vaults" }); // root if omitted
// [{ id, name, path, size, type: "file" | "directory", mtime }]
```

Failed requests throw a `Layerr` error whose info includes `status` and `authFailure` (`true` for a 401, meaning the access token needs refreshing).

For national clouds, pass the Graph endpoint: `new OneDriveClient(token, { endpoint: "https://microsoftgraph.chinacloudapi.cn/v1.0" })`.

### Patching requests

All client requests go through a [hot-patcher](https://github.com/perry-mitchell/hot-patcher) method named `"request"`, which can be replaced (for proxying, logging or testing):

```typescript
client.patcher.patch("request", (config) => myFetch(config.url, config));
```

## Testing sign-in against Microsoft

`npm test` uses fakes. `scripts/signin-test.mjs` is a manual check against the real service: it signs in with PKCE through a loopback redirect, writes, reads, lists and deletes one file, then refreshes the token. It needs an Entra app registration (public client, redirect `http://localhost`, delegated `Files.ReadWrite` and `offline_access`) and an account with OneDrive; a work account without a OneDrive licence fails with `403 notAllowed`.

```
npm run build
$env:ONEDRIVE_CLIENT_ID = "<application (client) ID>"
node scripts/signin-test.mjs
```

Optional: `ONEDRIVE_TENANT` (`common`, `consumers` or `organizations`) and `ONEDRIVE_INPRIVATE=1` to open a private Edge window. Tokens are never printed.
