# OneDrive Client changelog

## v0.1.0
_Unreleased_

 * Initial release
 * `OneDriveClient`: read, write, delete and list OneDrive items over the Microsoft Graph API, by drive item ID or by path. Requests go through a `hot-patcher` `"request"` hook, as in `@buttercup/googledrive-client`
 * `OAuth2Client`: Microsoft identity platform (v2.0) sign-in - authorisation URL (with optional PKCE), code exchange and token refresh, mirroring `@buttercup/google-oauth2-client`
