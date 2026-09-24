import sinon from "sinon";
import { expect } from "chai";
import { Layerr } from "layerr";
import { ERR_REFRESH_FAILED, OAuth2Client } from "../../dist/index.js";

function tokenResponse(body, status = 200) {
    return Promise.resolve({
        ok: status < 400,
        status,
        statusText: status < 400 ? "OK" : "Bad Request",
        json: () => Promise.resolve(body)
    });
}

function parseForm(body) {
    return Object.fromEntries(
        body.split("&").map(pair => pair.split("=").map(part => decodeURIComponent(part)))
    );
}

describe("OAuth2Client", function () {
    beforeEach(function () {
        this.client = new OAuth2Client("client-id", "http://localhost:1234/callback");
        this.fetchStub = sinon
            .stub(this.client, "_fetch")
            .returns(tokenResponse({ access_token: "at", refresh_token: "rt", expires_in: 3600 }));
    });

    describe("generateAuthUrl", function () {
        it("uses the common tenant by default", function () {
            const url = this.client.generateAuthUrl();
            expect(
                url.startsWith("https://login.microsoftonline.com/common/oauth2/v2.0/authorize?")
            ).to.be.true;
        });

        it("supports a specific tenant", function () {
            const client = new OAuth2Client("id", "http://x", { tenant: "consumers" });
            expect(client.generateAuthUrl()).to.contain("/consumers/oauth2/v2.0/authorize");
        });

        it("includes the client ID, redirect and default scopes", function () {
            const query = parseForm(this.client.generateAuthUrl().split("?")[1]);
            expect(query).to.include({
                client_id: "client-id",
                redirect_uri: "http://localhost:1234/callback",
                response_type: "code",
                scope: "Files.ReadWrite offline_access"
            });
        });

        it("includes PKCE, state and prompt when given", function () {
            const query = parseForm(
                this.client
                    .generateAuthUrl({
                        codeChallenge: "challenge",
                        state: "st",
                        prompt: "select_account",
                        scope: ["Files.ReadWrite.AppFolder", "offline_access"]
                    })
                    .split("?")[1]
            );
            expect(query).to.include({
                code_challenge: "challenge",
                code_challenge_method: "S256",
                state: "st",
                prompt: "select_account",
                scope: "Files.ReadWrite.AppFolder offline_access"
            });
        });

        it("omits PKCE when no challenge is given", function () {
            expect(this.client.generateAuthUrl()).to.not.contain("code_challenge");
        });
    });

    describe("exchangeAuthCodeForToken", function () {
        it("posts the code and verifier to the token endpoint", async function () {
            await this.client.exchangeAuthCodeForToken("the%20code", "verifier");
            const [url, options] = this.fetchStub.firstCall.args;
            expect(url).to.equal("https://login.microsoftonline.com/common/oauth2/v2.0/token");
            expect(options.method).to.equal("POST");
            expect(parseForm(options.body)).to.deep.equal({
                client_id: "client-id",
                grant_type: "authorization_code",
                code: "the code",
                redirect_uri: "http://localhost:1234/callback",
                code_verifier: "verifier"
            });
        });

        it("includes the client secret for confidential clients", async function () {
            const client = new OAuth2Client("id", "http://x", { clientSecret: "shh" });
            const stub = sinon.stub(client, "_fetch").returns(tokenResponse({ access_token: "a" }));
            await client.exchangeAuthCodeForToken("code");
            expect(parseForm(stub.firstCall.args[1].body)).to.include({ client_secret: "shh" });
        });

        it("stores and emits the tokens, converting expires_in to expiry_date", async function () {
            const emitted = sinon.spy();
            this.client.on("tokens", emitted);
            const before = Date.now();
            const { tokens } = await this.client.exchangeAuthCodeForToken("code");
            expect(this.client.accessToken).to.equal("at");
            expect(this.client.refreshToken).to.equal("rt");
            expect(tokens).to.not.have.property("expires_in");
            expect(tokens.expiry_date).to.be.at.least(before + 3600 * 1000);
            expect(emitted.calledOnceWith(tokens)).to.be.true;
        });

        it("throws on a failed exchange", async function () {
            this.fetchStub.returns(tokenResponse({ error: "invalid_grant" }, 400));
            let error = null;
            try {
                await this.client.exchangeAuthCodeForToken("code");
            } catch (err) {
                error = err;
            }
            expect(error).to.be.an("error");
        });
    });

    describe("refreshAccessToken", function () {
        it("posts the refresh token", async function () {
            await this.client.refreshAccessToken("old-rt");
            expect(parseForm(this.fetchStub.firstCall.args[1].body)).to.deep.equal({
                client_id: "client-id",
                grant_type: "refresh_token",
                refresh_token: "old-rt"
            });
        });

        it("shares concurrent refreshes of the same token", async function () {
            await Promise.all([
                this.client.refreshAccessToken("old-rt"),
                this.client.refreshAccessToken("old-rt")
            ]);
            expect(this.fetchStub.callCount).to.equal(1);
            await this.client.refreshAccessToken("old-rt");
            expect(this.fetchStub.callCount).to.equal(2);
        });

        it("keeps the previous refresh token if none is returned", async function () {
            this.fetchStub.returns(tokenResponse({ access_token: "new-at" }));
            const { tokens } = await this.client.refreshAccessToken("old-rt");
            expect(tokens.refresh_token).to.equal("old-rt");
            expect(this.client.refreshToken).to.equal("old-rt");
        });

        it("throws a refresh-failed error on a 4xx response", async function () {
            this.fetchStub.returns(tokenResponse({ error: "invalid_grant" }, 400));
            let error = null;
            try {
                await this.client.refreshAccessToken("old-rt");
            } catch (err) {
                error = err;
            }
            expect(Layerr.info(error)).to.include({ code: ERR_REFRESH_FAILED, status: 400 });
        });
    });
});
