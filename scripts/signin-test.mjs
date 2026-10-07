// Manual sign-in test against real Microsoft (B-466). Not part of `npm test`.
//
//   npm run build
//   $env:ONEDRIVE_CLIENT_ID = "<application (client) ID of the Entra app registration>"
//   node scripts/signin-test.mjs
//
// Optional: ONEDRIVE_TENANT (default "common"; "consumers" or "organizations" to narrow it),
// ONEDRIVE_INPRIVATE=1 to open the sign-in in a private Edge window instead of the default browser.
// Opens the browser, receives the code on a loopback port, then checks: code exchange,
// write/read/list/delete of one test file, and a refresh with the refresh token.
// Tokens are never printed or stored.
import { createServer } from "node:http";
import { createHash, randomBytes } from "node:crypto";
import { exec, spawn } from "node:child_process";
import { Layerr } from "layerr";
import { OAuth2Client, OneDriveClient, ONEDRIVE_DEFAULT_SCOPES } from "../dist/index.js";

const clientID = process.env.ONEDRIVE_CLIENT_ID;
if (!clientID) {
    console.error("Set ONEDRIVE_CLIENT_ID first.");
    process.exit(1);
}
const tenant = process.env.ONEDRIVE_TENANT || "common";
const TEST_PATH = "/buttercup-signin-test.txt";

const b64url = buf => buf.toString("base64url");
const verifier = b64url(randomBytes(48));
const challenge = b64url(createHash("sha256").update(verifier).digest());
const state = b64url(randomBytes(16));

const results = [];
function check(name, ok, detail = "") {
    results.push(ok);
    console.log(`${ok ? "PASS" : "FAIL"}  ${name}${detail ? ` - ${detail}` : ""}`);
}

function waitForCode(server) {
    return new Promise((resolve, reject) => {
        const timer = setTimeout(() => reject(new Error("No sign-in within 5 minutes")), 300000);
        server.on("request", (req, res) => {
            const url = new URL(req.url, "http://localhost");
            if (url.pathname !== "/") {
                res.writeHead(404).end();
                return;
            }
            const error = url.searchParams.get("error");
            res.writeHead(200, { "Content-Type": "text/plain" });
            res.end(error ? `Sign-in failed: ${error}` : "Signed in. You can close this tab.");
            clearTimeout(timer);
            if (error) {
                reject(new Error(`${error}: ${url.searchParams.get("error_description")}`));
            } else if (url.searchParams.get("state") !== state) {
                reject(new Error("State mismatch"));
            } else {
                resolve(url.searchParams.get("code"));
            }
        });
    });
}

const server = createServer();
await new Promise(resolve => server.listen(0, "127.0.0.1", resolve));
const redirect = `http://localhost:${server.address().port}`;
const oauth = new OAuth2Client(clientID, redirect, { tenant });
const authUrl = oauth.generateAuthUrl({
    codeChallenge: challenge,
    scope: ONEDRIVE_DEFAULT_SCOPES,
    state,
    prompt: "select_account"
});

console.log(`Redirect: ${redirect}`);
console.log("Opening the browser. If nothing opens, paste this URL:\n" + authUrl + "\n");
if (process.env.ONEDRIVE_INPRIVATE) {
    spawn("cmd", ["/c", "start", "msedge", "--inprivate", authUrl.replace(/&/g, "^&")], {
        detached: true,
        stdio: "ignore"
    }).unref();
} else {
    exec(`start "" "${authUrl.replace(/&/g, "^&")}"`);
}

let cleanup = null;
try {
    const code = await waitForCode(server);
    check("sign-in redirect returned a code", !!code);

    const { tokens } = await oauth.exchangeAuthCodeForToken(code, verifier);
    check("code exchange returned an access token", !!tokens.access_token);
    check("refresh token issued (offline_access)", !!tokens.refresh_token);
    console.log(`      granted scope: ${tokens.scope}`);
    check("expiry set", typeof tokens.expiry_date === "number" && tokens.expiry_date > Date.now());

    let client = new OneDriveClient(tokens.access_token);
    const body = `buttercup sign-in test ${new Date().toISOString()}`;
    cleanup = () => client.deleteFile({ path: TEST_PATH }).catch(() => {});

    await client.putFileContents({ path: TEST_PATH }, body);
    check("write file", true);
    check("read file back", (await client.getFileContents({ path: TEST_PATH })) === body);
    const listing = await client.getDirectoryContents({ path: "/" });
    check(
        "file shows in root listing",
        listing.some(item => item.name === TEST_PATH.slice(1))
    );

    const oldRefresh = tokens.refresh_token;
    const refreshed = await oauth.refreshAccessToken(oldRefresh);
    check("refresh returned a new access token", !!refreshed.tokens.access_token);
    check(
        "refresh token present after refresh",
        !!refreshed.tokens.refresh_token,
        refreshed.tokens.refresh_token === oldRefresh ? "same token kept" : "rotated"
    );

    client = new OneDriveClient(refreshed.tokens.access_token);
    check(
        "read file with the refreshed token",
        (await client.getFileContents({ path: TEST_PATH })) === body
    );

    await client.deleteFile({ path: TEST_PATH });
    cleanup = null;
    check("delete file", true);
} catch (err) {
    check("run completed", false, err.message);
    const info = Layerr.info(err);
    if (info.url) console.log(`      failed URL: ${info.url}`);
    if (oauth.accessToken) {
        // Ask Graph what it thinks of the token: the error code names the cause
        for (const path of ["/me/drive", "/me/drive/root"]) {
            const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
                headers: { Authorization: `Bearer ${oauth.accessToken}` }
            });
            const text = await res.text();
            let detail = text.slice(0, 300);
            try {
                const e = JSON.parse(text).error;
                if (e) detail = `${e.code}: ${e.message}`;
            } catch {}
            console.log(`      GET ${path} -> ${res.status} ${res.ok ? "ok" : detail}`);
        }
    }
    if (cleanup) await cleanup();
} finally {
    server.close();
}
console.log(results.every(Boolean) ? "\nALL PASSED" : "\nFAILED");
process.exit(results.every(Boolean) ? 0 : 1);
