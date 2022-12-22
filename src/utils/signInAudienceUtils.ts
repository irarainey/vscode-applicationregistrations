// Converts the audience to the correct format as required for the manifest.
export function convertSignInAudience(audience: string): string {
    return audience === "Single Tenant"
        ? "AzureADMyOrg"
        : audience === "Multiple Tenants"
            ? "AzureADMultipleOrgs"
            : "AzureADandPersonalMicrosoftAccount";
}