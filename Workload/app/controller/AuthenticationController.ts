import { AccessToken, WorkloadClientAPI } from "@ms-fabric/workload-client";

/**
 * Calls acquire frontend access token from the WorkloadClientAPI.
 * @param {WorkloadClientAPI} workloadClient - An instance of the WorkloadClientAPI.
 * @param {string} scopes - The scopes for which the access token is requested.
 * @returns {AccessToken}
 */
export async function callAcquireFrontendAccessToken(
    workloadClient: WorkloadClientAPI, 
    scopes: string): Promise<AccessToken> {
    return workloadClient.auth.acquireFrontendAccessToken({ scopes: scopes?.length ? scopes.split(' ') : [] });
}

/**
 * Acquires a token with automatic consent fallback.
 * Tries silent acquisition first, then falls back to interactive consent if needed.
 */
export async function acquireTokenWithConsent(
    workloadClient: WorkloadClientAPI,
    scope: string): Promise<AccessToken> {
    try {
        return await workloadClient.auth.acquireFrontendAccessToken({ scopes: [scope] });
    } catch {
        // Silent failed — trigger interactive consent popup
        try {
            await workloadClient.auth.acquireAccessToken({
                additionalScopesToConsent: [scope],
            });
            // Retry silent after consent
            return await workloadClient.auth.acquireFrontendAccessToken({ scopes: [scope] });
        } catch (interactiveErr: unknown) {
            const errCode = (interactiveErr as { error?: number })?.error;
            const messages: Record<number, string> = {
                0: 'Authentication is not supported in this environment.',
                1: 'User cancelled the consent dialog.',
                2: `Entra app is not configured for the scope "${scope}". Add the API permission in Azure Portal → App registrations → API permissions.`,
                3: 'Unknown authentication error.',
                4: `Invalid scope: "${scope}".`,
            };
            throw new Error(messages[errCode ?? -1] || `Authentication failed (code: ${errCode}).`);
        }
    }
}
