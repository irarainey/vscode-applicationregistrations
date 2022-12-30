// Converts the audience to the correct format as required.
export function convertSignInAudience(audience: string): string {
    return audience === "Single Tenant"
        ? "AzureADMyOrg"
        : audience === "Multiple Tenants"
            ? "AzureADMultipleOrgs"
            : audience === "Multiple Tenants and Personal Accounts"
                ? "AzureADandPersonalMicrosoftAccount"
                : "PersonalMicrosoftAccount";
}