# OneDrive Client changelog

## v0.1.0
_2026-11-26_ (CodeKunde)

 * Initial release, published on npm as `@buttercup-fork/onedrive-client`, staged from GitHub Actions through trusted publishing with provenance (`publish.yml`, started by hand)
 * Requests use the built-in `fetch` through `@buttercup-fork/fetch` 0.4 (installed under its old name `@buttercup/fetch`, so the imports are unchanged); upstream's `@buttercup/fetch` 0.2.1 and its `node-fetch` are no longer dependencies
 * Requires NodeJS 22 or newer; CI runs on branch pushes only (Node 22 and 24)
 * `OneDriveClient`: read, write, delete and list OneDrive items over the Microsoft Graph API, by drive item ID or by path. Requests go through a `hot-patcher` `"request"` hook, as in `@buttercup/googledrive-client`
 * `OAuth2Client`: Microsoft identity platform (v2.0) sign-in - authorisation URL (with optional PKCE), code exchange and token refresh, mirroring `@buttercup/google-oauth2-client`
 * Security review fixes (B-641 to B-648; behaviour changes before the first release):
   * Paths with `.` or `..` segments are refused (B-641)
   * The endpoint must be https on a known Graph host unless `allowCustomEndpoint` is set (B-642)
   * Directory listings refuse page links outside the endpoint and stop after 1000 pages (B-643)
   * Requests and token requests time out after 60 seconds, `timeout` option (B-644)
   * `generateAuthUrl` requires `state`, and a PKCE challenge unless the client has a secret; only `S256` (B-645)
   * The authorisation code is no longer decoded a second time (B-646)
   * A token response without an access token throws `ERR_BAD_TOKEN_RESPONSE` and is not emitted (B-647)
   * The error info `url` has no query string or fragment (B-648)
