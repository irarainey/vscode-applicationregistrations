// The name of the view
export const VIEW_NAME = "appRegistrations";

// The sign in command text
export const SIGNIN_COMMAND_TEXT = "Sign in to Azure CLI";

// The scope required when authenticating with Azure CLI
export const SCOPE = "https://graph.microsoft.com/.default";

// URI Roots for the Azure Portal variations
export const AZURE_PORTAL_APP_ROOT = "https://portal.azure.com";
export const AZURE_PORTAL_APP_ROOT_USGOV = "https://portal.azure.us";
export const AZURE_PORTAL_APP_ROOT_CHINA = "https://portal.azure.cn";
// TODO - consider approach to take for Azure portal access in national clouds: https://learn.microsoft.com/graph/deployments#app-registration-and-token-service-root-endpoints

// URI Roots for the Entra Portal variations
export const ENTRA_PORTAL_APP_ROOT = "https://entra.microsoft.com";
// TODO - determine (if available) Entra portal access in national clouds

// The URI path to the Application Registration blade in the Azure & Entra Portals
export const AZURE_AND_ENTRA_PORTAL_APP_PATH = "/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Overview/appId/";

// The URI path to the User blade in the Azure & Entra Portals
export const AZURE_AND_ENTRA_PORTAL_USER_PATH = "/#view/Microsoft_AAD_UsersAndTenants/UserProfileMenuBlade/~/overview/userId/";

// The URI to access directory objects in the Microsoft Graph API
export const DIRECTORY_OBJECTS_URI = "https://graph.microsoft.com/v1.0/directoryObjects/";
export const DIRECTORY_OBJECTS_URI_USGOV = "https://graph.microsoft.us/v1.0/directoryObjects/";
export const DIRECTORY_OBJECTS_URI_USGOV_DOD = "https://dod-graph.microsoft.us/v1.0/directoryObjects/";
export const DIRECTORY_OBJECTS_URI_CHINA = "https://microsoftgraph.chinacloudapi.cn/v1.0/directoryObjects/";
// TODO - consider approach to take for graph access in national clouds: https://learn.microsoft.com/graph/deployments#microsoft-graph-and-graph-explorer-service-root-endpoints

// The base tenant endpoints
export const BASE_ENDPOINT = "https://login.microsoftonline.com/";
export const BASE_ENDPOINT_USGOV = "https://login.microsoftonline.us/";
export const BASE_ENDPOINT_CHINA = "https://login.chinacloudapi.cn/";
// TODO - consider approach to take for login in national clouds: https://learn.microsoft.com/graph/deployments#app-registration-and-token-service-root-endpoints

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

// The Azure CLI command to login a user
export const CLI_LOGIN_CMD = "az login";

// The Azure CLI command to log out the current user
export const CLI_LOGOUT_CMD = "az logout";

// The Azure CLI command to get information about the current account
export const CLI_ACCOUNT_INFO_CMD= "az account show -o json";
