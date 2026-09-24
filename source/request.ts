import { Response, fetch } from "@buttercup/fetch";
import { Layerr } from "layerr";
import { Headers } from "./types.js";

export interface RequestConfig {
    body?: string | Buffer | ArrayBufferLike;
    headers?: Headers;
    method: "DELETE" | "GET" | "HEAD" | "PATCH" | "POST" | "PUT";
    url: string;
}

export function handleBadResponse(response: Response): void {
    if (!response.ok) {
        throw new Layerr(
            {
                info: {
                    // Graph returns 401 for missing, invalid or expired tokens
                    authFailure: response.status === 401,
                    status: response.status,
                    statusText: response.statusText,
                    url: response.url
                }
            },
            `Request failed: ${response.status} ${response.statusText}`
        );
    }
}

export async function request(config: RequestConfig): Promise<Response> {
    return fetch(config.url, {
        method: config.method,
        headers: config.headers,
        // @ts-ignore
        body: config.body
    });
}

/**
 * Encode an object as an application/x-www-form-urlencoded string. Written
 * out rather than using URLSearchParams, which is incomplete in React Native.
 */
export function encodeForm(data: Record<string, string | undefined>): string {
    return Object.keys(data)
        .filter(key => typeof data[key] === "string")
        .map(key => `${encodeURIComponent(key)}=${encodeURIComponent(data[key])}`)
        .join("&");
}
