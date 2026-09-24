import { HotPatcher } from "hot-patcher";
import { Response } from "@buttercup/fetch";
import { handleBadResponse, RequestConfig } from "../request.js";
import { itemURL } from "../util/target.js";
import { OneDriveItem, OneDriveTarget } from "../types.js";

interface GraphDriveItem {
    id: string;
    name: string;
    size?: number;
    folder?: object;
    lastModifiedDateTime?: string;
    parentReference?: {
        path?: string;
    };
}

interface GraphChildrenResponse {
    value: Array<GraphDriveItem>;
    "@odata.nextLink"?: string;
}

export interface InternalDirectoryRequestOptions {
    endpoint: string;
    patcher: HotPatcher;
    target: OneDriveTarget;
    token: string;
}

export async function getDirectoryContents(
    options: InternalDirectoryRequestOptions
): Promise<Array<OneDriveItem>> {
    const { endpoint, patcher, target, token } = options;
    const items: Array<OneDriveItem> = [];
    let url = itemURL(endpoint, target, "children");
    // Results are paged - follow @odata.nextLink until exhausted
    while (url) {
        const config: RequestConfig = {
            url,
            method: "GET",
            headers: {
                Authorization: `Bearer ${token}`
            }
        };
        const response = await patcher.execute<Promise<Response>>("request", config);
        handleBadResponse(response);
        const data = (await response.json()) as GraphChildrenResponse;
        items.push(...data.value.map(toOneDriveItem));
        url = data["@odata.nextLink"] || null;
    }
    return items;
}

function toOneDriveItem(item: GraphDriveItem): OneDriveItem {
    // parentReference.path looks like "/drive/root:/Some/Folder" (or
    // "/drive/root:" at the root)
    const parentPath = (item.parentReference?.path || "").replace(/^\/drive\/root:?/, "");
    return {
        id: item.id,
        name: item.name,
        path: `${parentPath}/${item.name}`,
        size: item.size ?? 0,
        type: item.folder ? "directory" : "file",
        mtime: item.lastModifiedDateTime ?? null
    };
}
