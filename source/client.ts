import { HotPatcher } from "hot-patcher";
import { request } from "./request.js";
import { getDirectoryContents } from "./requests/directoryContents.js";
import { deleteFile, getFileContents, putFileContents } from "./requests/fileContents.js";
import { GRAPH_ENDPOINT } from "./symbols.js";
import { OneDriveClientOptions, OneDriveItem, OneDriveTarget } from "./types.js";

export class OneDriveClient {
    public patcher: HotPatcher;
    private __endpoint: string;
    private __token: string;

    constructor(token: string, options: OneDriveClientOptions = {}) {
        const { endpoint = GRAPH_ENDPOINT } = options;
        this.__endpoint = endpoint.replace(/\/+$/, "");
        this.__token = token;
        this.patcher = new HotPatcher();
        this.patcher.patch("request", request);
    }

    /**
     * Delete a file or folder
     * @param target The item to delete
     */
    async deleteFile(target: OneDriveTarget): Promise<void> {
        await deleteFile({
            endpoint: this.__endpoint,
            patcher: this.patcher,
            target,
            token: this.__token
        });
    }

    /**
     * List the contents of a folder
     * @param target The folder - omit (or use `{ path: "/" }`) for the
     *  drive root
     */
    async getDirectoryContents(
        target: OneDriveTarget = { path: "/" }
    ): Promise<Array<OneDriveItem>> {
        return getDirectoryContents({
            endpoint: this.__endpoint,
            patcher: this.patcher,
            target,
            token: this.__token
        });
    }

    /**
     * Get the text contents of a file
     * @param target The file to read
     */
    async getFileContents(target: OneDriveTarget): Promise<string> {
        return getFileContents({
            endpoint: this.__endpoint,
            patcher: this.patcher,
            target,
            token: this.__token
        });
    }

    /**
     * Write the text contents of a file. Writing to a path that doesn't exist
     * creates the file.
     * @param target The file to write
     * @param contents The new contents
     * @returns The drive item ID of the written file
     */
    async putFileContents(target: OneDriveTarget, contents: string): Promise<string> {
        return putFileContents({
            contents,
            endpoint: this.__endpoint,
            patcher: this.patcher,
            target,
            token: this.__token
        });
    }
}
