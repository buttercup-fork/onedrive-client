import { Layerr } from "layerr";
import { GRAPH_HOSTS } from "../symbols.js";

/**
 * Check that an endpoint is an https URL on a known Graph host, so the access
 * token can't be sent anywhere else. Done with a pattern rather than `URL`,
 * which is incomplete in React Native.
 * @param endpoint The Graph API base URL
 * @param allowCustom Skip the host check (eg. for a test server)
 */
export function checkEndpoint(endpoint: string, allowCustom: boolean): void {
    if (allowCustom) return;
    const match = /^https:\/\/([^/?#@:]+)(\/[^?#]*)?$/i.exec(endpoint);
    if (!match || GRAPH_HOSTS.indexOf(match[1].toLowerCase()) < 0) {
        throw new Layerr(
            `OneDrive endpoint must be https on a Microsoft Graph host (got "${endpoint}"); set allowCustomEndpoint to override`
        );
    }
}
