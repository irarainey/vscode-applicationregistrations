// The name of the view
export const VIEW_NAME = "appRegistrations";

// The sign in command text
export const SIGNIN_COMMAND_TEXT = "Sign in to Azure CLI";

// The scope required when authenticating with Azure CLI
export const SCOPE = "https://graph.microsoft.com/.default";

// The URI to the Application Registration blade in the Azure Portal
export const PORTAL_APP_URI = "https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Overview/appId/";

// The URI to the User blade in the Azure Portal
export const PORTAL_USER_URI = "https://portal.azure.com/#view/Microsoft_AAD_UsersAndTenants/UserProfileMenuBlade/~/overview/userId/";

// The URI to access directory objects in the Microsoft Graph API
export const DIRECTORY_OBJECTS_URI = "https://graph.microsoft.com/v1.0/directoryObjects/";

// A list of properties to ignore when updating an application registration
export const PROPERTIES_TO_IGNORE_ON_UPDATE = [
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
export const SIGNIN_AUDIENCE_OPTIONS = [
    {
        label: "Single Tenant",
        description: "Accounts in this organizational directory only.",
        value: "AzureADMyOrg"
    },
    {
        label: "Multiple Tenants",
        description: "Accounts in any organizational directory.",
        value: "AzureADMultipleOrgs"
    },
    {
        label: "Multiple Tenants and Personal Accounts",
        description: "Azure AD and personal Microsoft accounts (e.g. Skype, Xbox).",
        value: "AzureADandPersonalMicrosoftAccount"
    },
    {
        label: "Personal Accounts",
        description: "Personal Microsoft accounts only.",
        value: "PersonalMicrosoftAccount"
    }
];

// URI to the documentation for sign in audience
export const SIGNIN_AUDIENCE_DOCUMENTATION_URI = "https://learn.microsoft.com/en-gb/azure/active-directory/develop/supported-accounts-validation";

// The properties to return when querying for applications
export const APPLICATION_SELECT_PROPERTIES = "id,displayName,appId,notes,signInAudience,appRoles,oauth2Permissions,web,spa,api,requiredResourceAccess,publicClient,identifierUris,passwordCredentials,keyCredentials";