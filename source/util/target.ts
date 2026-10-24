import { Layerr } from "layerr";
import { OneDriveTarget } from "../types.js";

/**
 * Build the Graph URL for a drive item, optionally with a sub-resource
 * (eg. "content" or "children")
 * @param endpoint The Graph API base URL
 * @param target The item, by ID or by path
 * @param resource Optional sub-resource
 */
export function itemURL(endpoint: string, target: OneDriveTarget, resource?: string): string {
    const { fileID, path } = target;
    const suffix = resource ? `/${resource}` : "";
    if (fileID) {
        return `${endpoint}/me/drive/items/${encodeURIComponent(fileID)}${suffix}`;
    }
    if (typeof path === "string") {
        const encodedPath = encodePath(path);
        if (!encodedPath) {
            // The drive root itself
            return `${endpoint}/me/drive/root${suffix}`;
        }
        return `${endpoint}/me/drive/root:/${encodedPath}:${suffix}`;
    }
    throw new Layerr("No OneDrive file ID or path specified");
}

function encodePath(path: string): string {
    return path
        .normalize("NFC")
        .split("/")
        .filter(segment => segment.length > 0)
        .map(segment => {
            // encodeURIComponent leaves dots alone and URL parsing would
            // resolve them, moving the request off the drive path
            if (segment === "." || segment === "..") {
                throw new Layerr("OneDrive path may not contain . or .. segments");
            }
            return encodeURIComponent(segment);
        })
        .join("/");
}
