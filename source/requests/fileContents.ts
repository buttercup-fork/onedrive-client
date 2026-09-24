import { HotPatcher } from "hot-patcher";
import { Response } from "@buttercup/fetch";
import { handleBadResponse, RequestConfig } from "../request.js";
import { itemURL } from "../util/target.js";
import { OneDriveTarget } from "../types.js";

export interface InternalFileRequestOptions {
    endpoint: string;
    patcher: HotPatcher;
    target: OneDriveTarget;
    token: string;
}

export async function getFileContents(options: InternalFileRequestOptions): Promise<string> {
    const { endpoint, patcher, target, token } = options;
    // Graph answers /content with a 302 to a pre-authenticated download URL,
    // which fetch follows
    const config: RequestConfig = {
        url: itemURL(endpoint, target, "content"),
        method: "GET",
        headers: {
            Authorization: `Bearer ${token}`
        }
    };
    const response = await patcher.execute<Promise<Response>>("request", config);
    handleBadResponse(response);
    return response.text();
}

export async function putFileContents(
    options: InternalFileRequestOptions & { contents: string }
): Promise<string> {
    const { contents, endpoint, patcher, target, token } = options;
    // Simple upload - fine for anything up to 250MB, far beyond vault sizes.
    // Writing to a path that doesn't exist creates the file (and any
    // missing parent folders).
    const config: RequestConfig = {
        url: itemURL(endpoint, target, "content"),
        method: "PUT",
        headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "text/plain"
        },
        body: contents
    };
    const response = await patcher.execute<Promise<Response>>("request", config);
    handleBadResponse(response);
    const { id } = (await response.json()) as { id: string };
    return id;
}

export async function deleteFile(options: InternalFileRequestOptions): Promise<void> {
    const { endpoint, patcher, target, token } = options;
    const config: RequestConfig = {
        url: itemURL(endpoint, target),
        method: "DELETE",
        headers: {
            Authorization: `Bearer ${token}`
        }
    };
    const response = await patcher.execute<Promise<Response>>("request", config);
    handleBadResponse(response);
}
