// The name of the view
export const view = "appRegistrations";

// The sign in command text
export const signInCommandText = "Sign in to Azure CLI...";

// The scope required when authenticating with Azure CLI
export const scope = "https://graph.microsoft.com/.default";

// The URI to the Application Registration blade in the Azure Portal
export const portalAppUri = "https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Overview/appId/";

// The URI to the User blade in the Azure Portal
export const portalUserUri = "https://portal.azure.com/#view/Microsoft_AAD_UsersAndTenants/UserProfileMenuBlade/~/overview/userId/";

// The URI to access directory objects in the Microsoft Graph API
export const directoryObjectsUri = "https://graph.microsoft.com/v1.0/directoryObjects/";

// A list of properties to ignore when updating an application registration
export const propertiesToIgnoreOnUpdate = [
    "appId",
    "publisherDomain",
    "passwordCredentials",
    "deletedDateTime",
    "disabledByMicrosoftStatus",
    "createdDateTime",
    "serviceManagementReference",
    "parentalControlSettings",
    "keyCredentials",
    "groupMembershipClaims",
    "tags",
    "tokenEncryptionKeyId"
];

// A list of sign in audience options
export const signInAudienceOptions = [
    "Single Tenant",
    "Multiple Tenants",
    "Multiple Tenants and Personal Accounts",
    "Personal Accounts"
];

// URI to the documentation for sign in audience
export const signInAudienceDocumentation = "https://learn.microsoft.com/en-gb/azure/active-directory/develop/supported-accounts-validation";

// The maximum number of applications returned from a query
export const maximumQueryApps = 200;

// The properties to return when querying for applications
export const appSelectProperties = "id,displayName,appId,notes,signInAudience,appRoles,oauth2Permissions,web,spa,api,requiredResourceAccess,publicClient,identifierUris,passwordCredentials,keyCredentials";