import { Response, fetch } from "@buttercup/fetch";
import { Layerr } from "layerr";
import { REQUEST_TIMEOUT } from "./symbols.js";
import { Headers } from "./types.js";

export interface RequestConfig {
    body?: string | Buffer | ArrayBufferLike;
    headers?: Headers;
    method: "DELETE" | "GET" | "HEAD" | "PATCH" | "POST" | "PUT";
    /**
     * Milliseconds to wait for the response, 0 for no limit
     */
    timeout?: number;
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
    const { signal, clear } = timeoutSignal(config.timeout);
    try {
        return await fetch(config.url, {
            method: config.method,
            headers: config.headers,
            signal,
            // @ts-ignore
            body: config.body
        });
    } finally {
        clear();
    }
}

/**
 * Create an abort signal that fires after a timeout. It covers the wait for
 * the response, not the reading of its body.
 * @param timeout Milliseconds (default `REQUEST_TIMEOUT`, 0 for no limit)
 */
export function timeoutSignal(timeout: number = REQUEST_TIMEOUT): {
    signal: AbortSignal | undefined;
    clear: () => void;
} {
    if (!timeout || timeout <= 0) return { signal: undefined, clear: () => {} };
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeout);
    return { signal: controller.signal, clear: () => clearTimeout(timer) };
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
