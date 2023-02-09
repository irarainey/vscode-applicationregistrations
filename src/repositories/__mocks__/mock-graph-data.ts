// This file contains mock data for the tests
export const mockApplications: any[] = [
	{
		id: "ab4e6904-6629-41c9-91d7-2ec9c7d3e46c",
		appId: "96a2745c-6920-4149-a8b0-90339bee74f9",
		createdDateTime: "2023-01-01T00:00:00Z",
		displayName: "First Test App",
		description: null,
		identifierUris: ["https://firsttestapp.com"],
		notes: null,
		signInAudience: "AzureADMultipleOrgs",
		api: {
			acceptMappedClaims: null,
			knownClientApplications: [],
			requestedAccessTokenVersion: null,
			oauth2PermissionScopes: [],
			preAuthorizedApplications: []
		},
		appRoles: [],
		keyCredentials: [],
		passwordCredentials: [
			{
				customKeyIdentifier: null,
				displayName: "NTITB",
				endDateTime: "2023-12-26T00:00:00Z",
				hint: "dof",
				keyId: "6e5f5f73-74ec-4ee4-b3dc-3f092f15fc13",
				secretText: null,
				startDateTime: "2023-01-26T16:14:42.0575529Z"
			}
		],
		publicClient: {
			redirectUris: []
		},
		requiredResourceAccess: [
			{
				resourceAppId: "00000003-0000-0000-c000-000000000000",
				resourceAccess: [
					{
						id: "e1fe6dd8-ba31-4d61-89e7-88639da4683d",
						type: "Scope"
					}
				]
			}
		],
		web: {
			homePageUrl: null,
			logoutUrl: "https://sample.com/logout",
			redirectUris: ["https://sample.com/callback"],
			implicitGrantSettings: {
				enableAccessTokenIssuance: false,
				enableIdTokenIssuance: false
			},
			redirectUriSettings: [
				{
					uri: "https://sample.com/callback",
					index: null
				}
			]
		},
		spa: {
			redirectUris: []
		},
		owners: [
			{
				id: "277c79dd-2e4d-40f0-9d9b-eff644ed74e8",
				displayName: "Sample User",
				userPrincipalName: "sample@user.com",
				mail: "sample@user.com"
			}
		]
	},
	{
		id: "0b8ca89e-c44f-4af3-a95a-575d1e758723",
		appId: "af019b40-4421-4734-bd8e-201c60d4e948",
		createdDateTime: "2023-01-01T00:00:00Z",
		displayName: "Second Test App",
		description: null,
		identifierUris: [],
		notes: null,
		signInAudience: "AzureADMultipleOrgs",
		api: {
			acceptMappedClaims: null,
			knownClientApplications: [],
			requestedAccessTokenVersion: null,
			oauth2PermissionScopes: [],
			preAuthorizedApplications: []
		},
		appRoles: [],
		keyCredentials: [],
		passwordCredentials: [],
		publicClient: {
			redirectUris: []
		},
		requiredResourceAccess: [],
		web: {
			homePageUrl: null,
			logoutUrl: null,
			redirectUris: [],
			implicitGrantSettings: {
				enableAccessTokenIssuance: false,
				enableIdTokenIssuance: false
			},
			redirectUriSettings: []
		},
		spa: {
			redirectUris: []
		},
		owners: [
			{
				id: "277c79dd-2e4d-40f0-9d9b-eff644ed74e8",
				displayName: "Sample User",
				userPrincipalName: "sample@user.com",
				mail: "sample@user.com"
			}
		]
	}
];

export const mockUser: any = {
	id: "277c79dd-2e4d-40f0-9d9b-eff644ed74e8",
	displayName: "Sample User",
	givenName: "Sample",
	mail: "sample@user.com",
	surname: "User",
	userPrincipalName: "sample@user.com"
};

export const mockOrganizations: any[] = [
	{
		id: "c7b3da28-01b8-46d3-9523-d1b24cbbde76",
		displayName: "Sample Tenant",
		verifiedDomains: [
			{ capabilities: "Email, OfficeCommunicationsOnline", isDefault: false, isInitial: true, name: "sample.onmicrosoft.com", type: "Managed" },
			{ capabilities: "Email", isDefault: true, isInitial: false, name: "test.com", type: "Managed" }
		]
	}
];

export const mockRoleAssignments: any[] = [
	{
		principalId: "277c79dd-2e4d-40f0-9d9b-eff644ed74e8",
		roleDefinition: {
			id: "62e90394-69f5-4237-9190-012177145e10",
			description: "Can manage all aspects of Azure AD and Microsoft services that use Azure AD identities.",
			displayName: "Global Administrator"
		}
	}
];
