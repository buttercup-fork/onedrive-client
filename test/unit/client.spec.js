import { createServer } from "node:http";
import sinon from "sinon";
import { expect } from "chai";
import { Layerr } from "layerr";
import { GRAPH_ENDPOINT, OneDriveClient } from "../../dist/index.js";

const FAKE_TOKEN = "aaaaabbbbbbccccccddddddeeeeee";

function okResponse({ text = "", json = {} } = {}) {
    return Promise.resolve({
        ok: true,
        status: 200,
        statusText: "OK",
        text: () => Promise.resolve(text),
        json: () => Promise.resolve(json)
    });
}

describe("OneDriveClient", function () {
    beforeEach(function () {
        this.client = new OneDriveClient(FAKE_TOKEN);
        this.requestStub = sinon.stub();
        this.client.patcher.patch("request", this.requestStub);
    });

    describe("getFileContents", function () {
        beforeEach(function () {
            this.requestStub.returns(okResponse({ text: "vault\ncontents" }));
        });

        it("reads by file ID", async function () {
            const contents = await this.client.getFileContents({ fileID: "ABC!12" });
            expect(contents).to.equal("vault\ncontents");
            const config = this.requestStub.firstCall.args[0];
            expect(config.method).to.equal("GET");
            expect(config.url).to.equal(`${GRAPH_ENDPOINT}/me/drive/items/ABC!12/content`);
            expect(config.headers.Authorization).to.equal(`Bearer ${FAKE_TOKEN}`);
        });

        it("reads by path, encoding each segment", async function () {
            await this.client.getFileContents({ path: "/Vaults/My Vault #1.bcup" });
            expect(this.requestStub.firstCall.args[0].url).to.equal(
                `${GRAPH_ENDPOINT}/me/drive/root:/Vaults/My%20Vault%20%231.bcup:/content`
            );
        });

        it("normalises paths to NFC", async function () {
            // "ü" as u + combining diaeresis (NFD, as macOS produces)
            await this.client.getFileContents({ path: "Tresor ü.bcup" });
            expect(this.requestStub.firstCall.args[0].url).to.contain(
                encodeURIComponent("Tresor ü.bcup")
            );
        });

        it("prefers file ID over path", async function () {
            await this.client.getFileContents({ fileID: "id1", path: "ignored.bcup" });
            expect(this.requestStub.firstCall.args[0].url).to.contain("/items/id1/");
        });

        it("throws without a target", async function () {
            let error = null;
            try {
                await this.client.getFileContents({});
            } catch (err) {
                error = err;
            }
            expect(error).to.be.an("error");
        });

        it("marks 401 responses as auth failures", async function () {
            this.requestStub.returns(
                Promise.resolve({ ok: false, status: 401, statusText: "Unauthorized" })
            );
            let error = null;
            try {
                await this.client.getFileContents({ path: "a.bcup" });
            } catch (err) {
                error = err;
            }
            expect(Layerr.info(error)).to.include({ authFailure: true, status: 401 });
        });

        it("does not mark other failures as auth failures", async function () {
            this.requestStub.returns(
                Promise.resolve({ ok: false, status: 404, statusText: "Not Found" })
            );
            let error = null;
            try {
                await this.client.getFileContents({ path: "a.bcup" });
            } catch (err) {
                error = err;
            }
            expect(Layerr.info(error)).to.include({ authFailure: false, status: 404 });
        });
    });

    describe("putFileContents", function () {
        beforeEach(function () {
            this.requestStub.returns(okResponse({ json: { id: "NEWID" } }));
        });

        it("uploads text and returns the item ID", async function () {
            const id = await this.client.putFileContents({ path: "a/b.bcup" }, "data");
            expect(id).to.equal("NEWID");
            const config = this.requestStub.firstCall.args[0];
            expect(config.method).to.equal("PUT");
            expect(config.url).to.equal(`${GRAPH_ENDPOINT}/me/drive/root:/a/b.bcup:/content`);
            expect(config.body).to.equal("data");
            expect(config.headers["Content-Type"]).to.equal("text/plain");
        });
    });

    describe("deleteFile", function () {
        it("deletes the item", async function () {
            this.requestStub.returns(okResponse());
            await this.client.deleteFile({ fileID: "x1" });
            const config = this.requestStub.firstCall.args[0];
            expect(config.method).to.equal("DELETE");
            expect(config.url).to.equal(`${GRAPH_ENDPOINT}/me/drive/items/x1`);
        });
    });

    describe("getDirectoryContents", function () {
        it("lists the root by default, following pagination", async function () {
            this.requestStub.onFirstCall().returns(
                okResponse({
                    json: {
                        value: [
                            {
                                id: "1",
                                name: "Vaults",
                                folder: { childCount: 1 },
                                parentReference: { path: "/drive/root:" }
                            }
                        ],
                        "@odata.nextLink": "https://graph.microsoft.com/v1.0/next-page"
                    }
                })
            );
            this.requestStub.onSecondCall().returns(
                okResponse({
                    json: {
                        value: [
                            {
                                id: "2",
                                name: "notes.txt",
                                size: 12,
                                lastModifiedDateTime: "2026-01-01T00:00:00Z",
                                parentReference: { path: "/drive/root:" }
                            }
                        ]
                    }
                })
            );
            const items = await this.client.getDirectoryContents();
            expect(this.requestStub.firstCall.args[0].url).to.equal(
                `${GRAPH_ENDPOINT}/me/drive/root/children`
            );
            expect(this.requestStub.secondCall.args[0].url).to.equal(
                "https://graph.microsoft.com/v1.0/next-page"
            );
            expect(items).to.deep.equal([
                {
                    id: "1",
                    name: "Vaults",
                    path: "/Vaults",
                    size: 0,
                    type: "directory",
                    mtime: null
                },
                {
                    id: "2",
                    name: "notes.txt",
                    path: "/notes.txt",
                    size: 12,
                    type: "file",
                    mtime: "2026-01-01T00:00:00Z"
                }
            ]);
        });

        it("lists a sub-folder by path", async function () {
            this.requestStub.returns(
                okResponse({
                    json: {
                        value: [
                            {
                                id: "3",
                                name: "my.bcup",
                                size: 100,
                                parentReference: { path: "/drive/root:/Vaults" }
                            }
                        ]
                    }
                })
            );
            const items = await this.client.getDirectoryContents({ path: "/Vaults" });
            expect(this.requestStub.firstCall.args[0].url).to.equal(
                `${GRAPH_ENDPOINT}/me/drive/root:/Vaults:/children`
            );
            expect(items[0].path).to.equal("/Vaults/my.bcup");
        });
    });

    it("supports a national cloud endpoint", async function () {
        const client = new OneDriveClient(FAKE_TOKEN, {
            endpoint: "https://microsoftgraph.chinacloudapi.cn/v1.0/"
        });
        const stub = sinon.stub().returns(okResponse());
        client.patcher.patch("request", stub);
        await client.getFileContents({ fileID: "x" });
        expect(stub.firstCall.args[0].url).to.equal(
            "https://microsoftgraph.chinacloudapi.cn/v1.0/me/drive/items/x/content"
        );
    });

    describe("endpoint check", function () {
        const BAD = [
            "http://graph.microsoft.com/v1.0",
            "https://evil.example/v1.0",
            "https://graph.microsoft.com.evil.example/v1.0",
            "https://graph.microsoft.com@evil.example/v1.0",
            "https://evil.example@graph.microsoft.com/v1.0",
            "https://graph.microsoft.com:8443/v1.0",
            "https://graph.microsoft.com/v1.0?x=1",
            "not a url",
            ""
        ];
        for (const endpoint of BAD) {
            it(`refuses ${JSON.stringify(endpoint)}`, function () {
                expect(() => new OneDriveClient(FAKE_TOKEN, { endpoint })).to.throw(
                    "must be https on a Microsoft Graph host"
                );
            });
        }

        it("allows another endpoint when asked to", async function () {
            const client = new OneDriveClient(FAKE_TOKEN, {
                allowCustomEndpoint: true,
                endpoint: "http://localhost:8080/v1.0"
            });
            const stub = sinon.stub().returns(okResponse());
            client.patcher.patch("request", stub);
            await client.getFileContents({ fileID: "x" });
            expect(stub.firstCall.args[0].url).to.equal(
                "http://localhost:8080/v1.0/me/drive/items/x/content"
            );
        });
    });

    describe("path segments", function () {
        for (const path of ["/../../items/OTHER", "/a/../b", "/./a", ".."]) {
            it(`refuses ${JSON.stringify(path)}`, async function () {
                let error = null;
                try {
                    await this.client.getFileContents({ path });
                } catch (err) {
                    error = err;
                }
                expect(error).to.be.an("error");
                expect(error.message).to.contain(". or ..");
                expect(this.requestStub.called).to.be.false;
            });
        }

        it("still allows dots inside a name", async function () {
            this.requestStub.returns(okResponse({ text: "x" }));
            await this.client.getFileContents({ path: "/a../.hidden/b.c..d" });
            expect(this.requestStub.firstCall.args[0].url).to.equal(
                `${GRAPH_ENDPOINT}/me/drive/root:/a../.hidden/b.c..d:/content`
            );
        });
    });

    describe("page links", function () {
        it("refuses a link outside the endpoint, before sending the token", async function () {
            this.requestStub.returns(
                okResponse({
                    json: { value: [], "@odata.nextLink": "http://evil.example/steal" }
                })
            );
            let error = null;
            try {
                await this.client.getDirectoryContents();
            } catch (err) {
                error = err;
            }
            expect(error).to.be.an("error");
            expect(error.message).to.contain("outside the OneDrive endpoint");
            expect(this.requestStub.callCount).to.equal(1);
        });

        it("refuses a link on a look-alike host", async function () {
            this.requestStub.returns(
                okResponse({
                    json: {
                        value: [],
                        "@odata.nextLink": "https://graph.microsoft.com.evil.example/v1.0/x"
                    }
                })
            );
            let error = null;
            try {
                await this.client.getDirectoryContents();
            } catch (err) {
                error = err;
            }
            expect(error).to.be.an("error");
            expect(this.requestStub.callCount).to.equal(1);
        });

        it("stops after too many pages", async function () {
            this.requestStub.callsFake(() =>
                okResponse({
                    json: {
                        value: [],
                        "@odata.nextLink": `${GRAPH_ENDPOINT}/me/drive/root/children?page=next`
                    }
                })
            );
            let error = null;
            try {
                await this.client.getDirectoryContents();
            } catch (err) {
                error = err;
            }
            expect(error).to.be.an("error");
            expect(error.message).to.contain("Too many directory pages");
            expect(this.requestStub.callCount).to.equal(1000);
        });
    });

    describe("timeout", function () {
        it("aborts the default request when the response takes too long", async function () {
            // A real server that never answers
            const server = createServer(() => {});
            await new Promise(resolve => server.listen(0, "127.0.0.1", resolve));
            const client = new OneDriveClient(FAKE_TOKEN, {
                allowCustomEndpoint: true,
                endpoint: `http://127.0.0.1:${server.address().port}/v1.0`,
                timeout: 50
            });
            let error = null;
            const started = Date.now();
            try {
                await client.getFileContents({ fileID: "x" });
            } catch (err) {
                error = err;
            }
            server.closeAllConnections();
            server.close();
            expect(error).to.have.property("name", "AbortError");
            expect(Date.now() - started).to.be.below(2000);
        });
    });
});
