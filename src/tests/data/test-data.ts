// This file contains mock data for the tests
export let mockApplications: any[];
export let mockUser: any;
export let mockUsers: any[];
export let mockOrganizations: any[];
export let mockRoleAssignments: any[];
export let mockRedirectUris: any;

export const seedMockData = () => {
	mockApplications = JSON.parse(JSON.stringify(seedApplications));
	mockUser = JSON.parse(JSON.stringify(seedUser));
	mockUsers = JSON.parse(JSON.stringify(seedUsers));
	mockOrganizations = JSON.parse(JSON.stringify(seedOrganizations));
	mockRoleAssignments = JSON.parse(JSON.stringify(seedRoleAssignments));
	mockRedirectUris = JSON.parse(JSON.stringify(seedRedirectUris));
};

const seedApplications: any[] = [
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
		appRoles: [
			{
				"allowedMemberTypes": [
					"User",
					"Application"
				],
				"description": "Allows the writing of files",
				"displayName": "Write Files",
				"id": "d93e3770-dccc-44d7-9146-f4dc9abb43f9",
				"isEnabled": true,
				"origin": "Application",
				"value": "Files.Write"
			},
			{
				"allowedMemberTypes": [
					"User"
				],
				"description": "Allows the reading of files",
				"displayName": "Read Files",
				"id": "cd9bf494-5cfb-4a26-8597-d3daad89e446",
				"isEnabled": true,
				"origin": "Application",
				"value": "Files.Read"
			}
		],
		keyCredentials: [
			{
				customKeyIdentifier: "F56DC6F6E469B781F8BA664DF842788E5D197F24",
				displayName: "CN=test1.com",
				endDateTime: "2025-01-01T00:00:00Z",
				key: null,
				keyId: "865b1dbb-5c6d-4e9c-81a3-474af6556bc3",
				startDateTime: "2023-01-01T00:00:00Z",
				type: "AsymmetricX509Cert",
				usage: "Verify"
			},
			{
				customKeyIdentifier: "E46DC6F6E469B781F8BA664DF842788E5D197F24",
				displayName: "CN=test2.com",
				endDateTime: "2024-01-01T00:00:00Z",
				key: null,
				keyId: "ec5db659-6c18-4803-a172-77a6e0dbdb6e",
				startDateTime: "2022-01-01T00:00:00Z",
				type: "AsymmetricX509Cert",
				usage: "Verify"
			}
		],
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
			redirectUris: ["https://mobile.com"]
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
			redirectUris: ["https://spa.com"]
		},
		owners: [
			{
				id: "277c79dd-2e4d-40f0-9d9b-eff644ed74e8",
				displayName: "First User",
				userPrincipalName: "first@user.com",
				mail: "first@user.com"
			},
			{
				id: "b204cef1-baa1-4284-8ae1-2285a609ba35",
				displayName: "Second User",
				userPrincipalName: "second@user.com",
				mail: "second@user.com"
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
				displayName: "First User",
				userPrincipalName: "first@user.com",
				mail: "first@user.com"
			}
		]
	}
];

const seedRedirectUris: any = {
	web: {
		redirectUris: Array.from({ length: 100 }, (_, i) => `https://web${i}.com`)
	},
	publicClient: {
		redirectUris: Array.from({ length: 100 }, (_, i) => `https://native${i}.com`)
	},
	spa: {
		redirectUris: Array.from({ length: 56 }, (_, i) => `https://spa${i}.com`)
	}
};

const seedUser: any = {
	id: "277c79dd-2e4d-40f0-9d9b-eff644ed74e8",
	displayName: "First User",
	givenName: "First",
	mail: "first@user.com",
	surname: "User",
	userPrincipalName: "first@user.com"
};

const seedUsers: any[] = [
	{
		id: "277c79dd-2e4d-40f0-9d9b-eff644ed74e8",
		displayName: "First User",
		givenName: "First",
		mail: "first@user.com",
		surname: "User",
		userPrincipalName: "first@user.com"
	},
	{
		id: "b204cef1-baa1-4284-8ae1-2285a609ba35",
		displayName: "Second User",
		givenName: "Second",
		mail: "second@user.com",
		surname: "User",
		userPrincipalName: "second@user.com"
	}
];

const seedOrganizations: any[] = [
	{
		id: "c7b3da28-01b8-46d3-9523-d1b24cbbde76",
		displayName: "Sample Tenant",
		verifiedDomains: [
			{ capabilities: "Email, OfficeCommunicationsOnline", isDefault: false, isInitial: true, name: "sample.onmicrosoft.com", type: "Managed" },
			{ capabilities: "Email", isDefault: true, isInitial: false, name: "test.com", type: "Managed" }
		]
	}
];

const seedRoleAssignments: any[] = [
	{
		principalId: "277c79dd-2e4d-40f0-9d9b-eff644ed74e8",
		roleDefinition: {
			id: "62e90394-69f5-4237-9190-012177145e10",
			description: "Can manage all aspects of Azure AD and Microsoft services that use Azure AD identities.",
			displayName: "Global Administrator"
		}
	}
];

export const mockPemCertificate = `-----BEGIN CERTIFICATE-----
	MIIDkTCCAnmgAwIBAgIUeuJnMCKaW8S9+Kw51mez+s1JAK0wDQYJKoZIhvcNAQEL
	BQAwWDELMAkGA1UEBhMCVUsxEzARBgNVBAgMClNvbWUtU3RhdGUxITAfBgNVBAoM
	GEludGVybmV0IFdpZGdpdHMgUHR5IEx0ZDERMA8GA1UEAwwIdGVzdC5jb20wHhcN
	MjMwMjE0MTU0NzM1WhcNMzMwMjExMTU0NzM1WjBYMQswCQYDVQQGEwJVSzETMBEG
	A1UECAwKU29tZS1TdGF0ZTEhMB8GA1UECgwYSW50ZXJuZXQgV2lkZ2l0cyBQdHkg
	THRkMREwDwYDVQQDDAh0ZXN0LmNvbTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCC
	AQoCggEBAMD921ZyHxy8imjXAjpY0NkgjbUbkDsWfSsCH4jrPF8EUfib4tRBRehA
	82FOjWR8blKMU9nGngGbX/X6CEv9qVM6dzcyieSMHbPYdOKGUo4dJEpEnFY371D6
	9adQ0TUuHJ2p3Y1wmQfK7G70WGnPgrwmTOCs8Es2+PwjSJxX4PR2ipGmw8OKb9dQ
	PfXe+odLs4ek1ucnSi84AQu6yR8ZII7Wf3lGZ3DVmaWKjPZUJO9Di0jjIVe8JxfE
	WuQDpIdHttbjWKABEQdCAP1mwsU2EXJwf+yC3QGkewGVXHdIpF9MfUYmQ2/ivE9X
	XtKp2URnD7RwpWtg37OlRJ4esq4+5JsCAwEAAaNTMFEwHQYDVR0OBBYEFGBkZKMW
	kvxDPCPYzH4aoZGhMEVyMB8GA1UdIwQYMBaAFGBkZKMWkvxDPCPYzH4aoZGhMEVy
	MA8GA1UdEwEB/wQFMAMBAf8wDQYJKoZIhvcNAQELBQADggEBALDXmQ+bpT4IxEF6
	KB6aWPIMLRBfTgyIpFFVckZlygFOdxt0akGYAixSy6NxnDtVwEUjOxxCQzxDbAzk
	HAJs8fJfpdDnjym2FJAMUX1lE5Au9/u6IRuEpoOFAbSwJG7Z/faVUSX0VM/+YWo4
	LPl8uL0LCVSAu4MSo2sxGStJaBPRUNDBCviSs1B1FTejVO01LH0iSJnFOHbVZaZF
	hUxQ8Ym1fB+8Ml5H49eqPutGwj1sybj1wH434uS99TSH0vNlb81QOo2kwRAqiaim
	3fWYhZsV9VHS/7gSsYBOktuSGCsEA8QE4nloJKG1RwlaVDcq6bs8Tj1GNHkqfZqF
	x1sLjRM=
	-----END CERTIFICATE-----`;

export const mockTenantId = "c7b3da28-01b8-46d3-9523-d1b24cbbde76";
export const mockNewPasswordKeyId = "a7da2abf-da93-4bad-bf0b-6d9ee0d3e8ec";
export const mockAppObjectId = seedApplications[0].id;
export const mockSecondAppObjectId = seedApplications[1].id;
export const mockAppId = seedApplications[0].appId;
export const mockSecondAppId = seedApplications[1].appId;
export const mockUserId = seedUsers[0].id;
export const mockSecondUserId = seedUsers[1].id;
export const mockAppRoleId = seedApplications[0].appRoles[0].id;
