// This file contains mock data for the tests
export let mockApplications: any[];
export let mockUser: any;
export let mockUsers: any[];
export let mockOrganizations: any[];
export let mockRoleAssignments: any[];
export let mockRedirectUris: any;
export let mockServicePrincipals: any[];

export const seedMockData = () => {
	mockApplications = JSON.parse(JSON.stringify(seedApplications));
	mockUser = JSON.parse(JSON.stringify(seedUser));
	mockUsers = JSON.parse(JSON.stringify(seedUsers));
	mockOrganizations = JSON.parse(JSON.stringify(seedOrganizations));
	mockRoleAssignments = JSON.parse(JSON.stringify(seedRoleAssignments));
	mockRedirectUris = JSON.parse(JSON.stringify(seedRedirectUris));
	mockServicePrincipals = JSON.parse(JSON.stringify(seedServicePrincipals));
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
			oauth2PermissionScopes: [
				{
					adminConsentDescription: "Sample description one",
					adminConsentDisplayName: "Sample Scope One",
					id: "2af52627-d1b4-408e-b188-ccca2a5cd33c",
					isEnabled: true,
					type: "User",
					userConsentDescription: null,
					userConsentDisplayName: null,
					value: "Sample.One"
				},
				{
					adminConsentDescription: "Sample description two",
					adminConsentDisplayName: "Sample Scope Two",
					id: "97956368-899b-4e0b-b51e-61743cb60f36",
					isEnabled: true,
					type: "Admin",
					userConsentDescription: null,
					userConsentDisplayName: null,
					value: "Sample.Two"
				}
			],
			preAuthorizedApplications: []
		},
		appRoles: [
			{
				allowedMemberTypes: ["User", "Application"],
				description: "Allows the writing of files",
				displayName: "Write Files",
				id: "d93e3770-dccc-44d7-9146-f4dc9abb43f9",
				isEnabled: true,
				origin: "Application",
				value: "Files.Write"
			},
			{
				allowedMemberTypes: ["User"],
				description: "Allows the reading of files",
				displayName: "Read Files",
				id: "cd9bf494-5cfb-4a26-8597-d3daad89e446",
				isEnabled: true,
				origin: "Application",
				value: "Files.Read"
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
						id: "570282fd-fa5c-430d-a7fd-fc8dc98a9dca",
						type: "Scope"
					},
					{
						id: "e383f46e-2787-4529-855e-0e479a3ffac0",
						type: "Scope"
					},
					{
						id: "810c84a8-4a9e-49e6-bf7d-12d183f40d01",
						type: "Role"
					}
				]
			},
			{
				resourceAppId: "1173dc06-edfc-40fc-980c-ff7d7ceb144d",
				resourceAccess: [
					{
						id: "ef080234-1cd3-4945-9462-0d8901fd327a",
						type: "Scope"
					},
					{
						id: "f9ae202f-8d70-47a9-b1f4-78ba9d9d906f",
						type: "Scope"
					},
					{
						id: "8f583345-62f3-4c1e-aac2-a85584d380d3",
						type: "Role"
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
		notes: "Second Test App",
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
	},
	{
		id: "2352b377-784d-4d20-a874-319ca672b637",
		appId: "0515ee31-a64b-4a03-9d15-ff10c7dbfe39",
		createdDateTime: "2023-01-01T00:00:00Z",
		displayName: "Third Test App",
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
				id: "b204cef1-baa1-4284-8ae1-2285a609ba35",
				displayName: "Second User",
				userPrincipalName: "second@user.com",
				mail: "second@user.com"
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
	},
	{
		id: "57f2aafa-0a41-43cf-a9b9-8897b813e1b1",
		displayName: "Third User",
		givenName: "Third One",
		mail: "third.one@user.com",
		surname: "User",
		userPrincipalName: "third.one@user.com"
	},
	{
		id: "7e354ea2-2b9a-4a4f-add7-7d47a4c1c6dd",
		displayName: "Third User",
		givenName: "Third Two",
		mail: "third.two@user.com",
		surname: "User",
		userPrincipalName: "third.two@user.com"
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

export const seedServicePrincipals = [
	{
		id: "2b60636f-bc43-4069-9644-bc4108f43407",
		appDisplayName: "Microsoft Graph",
		appId: "00000003-0000-0000-c000-000000000000",
		displayName: "Microsoft Graph",
		appRoles: [
			{
				allowedMemberTypes: ["Application"],
				description: "Allows the app to read mail in all mailboxes without a signed-in user.",
				displayName: "Read mail in all mailboxes",
				id: "810c84a8-4a9e-49e6-bf7d-12d183f40d01",
				isEnabled: true,
				origin: "Application",
				value: "Mail.Read"
			},
			{
				allowedMemberTypes: ["Application"],
				description: "Allows the app to send mail as any user without a signed-in user.",
				displayName: "Send mail as any user",
				id: "b633e1c5-b582-4048-a93e-9f11b44c7e96",
				isEnabled: true,
				origin: "Application",
				value: "Mail.Send"
			}
		],
		oauth2PermissionScopes: [
			{
				adminConsentDescription: "Allows the app to read the signed-in user's mailbox.",
				adminConsentDisplayName: "Read user mail ",
				id: "570282fd-fa5c-430d-a7fd-fc8dc98a9dca",
				isEnabled: true,
				type: "User",
				userConsentDescription: "Allows the app to read email in your mailbox. ",
				userConsentDisplayName: "Read your mail ",
				value: "Mail.Read"
			},
			{
				adminConsentDescription: "Allows the app to send mail as users in the organization. ",
				adminConsentDisplayName: "Send mail as a user ",
				id: "e383f46e-2787-4529-855e-0e479a3ffac0",
				isEnabled: true,
				type: "User",
				userConsentDescription: "Allows the app to send mail as you. ",
				userConsentDisplayName: "Send mail as you ",
				value: "Mail.Send"
			},
			{
				adminConsentDescription: "Allows the app to read events in user calendars . ",
				adminConsentDisplayName: "Read user calendars ",
				id: "465a38f9-76ea-45b9-9f34-9e8b0d4b0b42",
				isEnabled: true,
				type: "User",
				userConsentDescription: "Allows the app to read events in your calendars. ",
				userConsentDisplayName: "Read your calendars ",
				value: "Calendars.Read"
			}
		]
	},
	{
		id: "f82e47c3-ac80-4662-986b-c64e2204e20d",
		appDisplayName: "Bert's Cookies API",
		appId: "1173dc06-edfc-40fc-980c-ff7d7ceb144d",
		displayName: "Bert's Cookies API",
		appRoles: [
			{
				allowedMemberTypes: ["User"],
				description: "Staff who work in the store",
				displayName: "Store Staff",
				id: "8f583345-62f3-4c1e-aac2-a85584d380d3",
				isEnabled: true,
				origin: "Application",
				value: "Staff.Store"
			}
		],
		oauth2PermissionScopes: [
			{
				adminConsentDescription: "Allows users to sell cookies",
				adminConsentDisplayName: "Sell cookies",
				id: "ef080234-1cd3-4945-9462-0d8901fd327a",
				isEnabled: true,
				type: "User",
				userConsentDescription: null,
				userConsentDisplayName: null,
				value: "Cookies.Sell"
			},
			{
				adminConsentDescription: "Allows users to eat cookies",
				adminConsentDisplayName: "Eat cookies",
				id: "f9ae202f-8d70-47a9-b1f4-78ba9d9d906f",
				isEnabled: true,
				type: "Admin",
				userConsentDescription: null,
				userConsentDisplayName: null,
				value: "Cookies.Eat"
			}
		]
	},
	{
		id: "7f808afc-8ba4-409e-90cd-27a2f7f9a1f8",
		appDisplayName: "Random Api",
		appId: "0cad2264-23c3-4366-8c28-805aaeda257b",
		displayName: "Random Api",
		appRoles: [],
		oauth2PermissionScopes: [
			{
				adminConsentDescription: "Something",
				adminConsentDisplayName: "Something",
				id: "82062f02-7837-45e6-a497-4938246ceb5c",
				isEnabled: true,
				type: "Admin",
				userConsentDescription: null,
				userConsentDisplayName: null,
				value: "Do.Something"
			}
		]
	}
];

export const mockTenantId = "c7b3da28-01b8-46d3-9523-d1b24cbbde76";
export const mockNewPasswordKeyId = "a7da2abf-da93-4bad-bf0b-6d9ee0d3e8ec";
export const mockGraphApiAppId = "00000003-0000-0000-c000-000000000000";
export const mockExposedApiId = seedApplications[0].api.oauth2PermissionScopes[0].id;
export const mockApiScopeId = seedApplications[0].requiredResourceAccess[0].resourceAppId;
export const mockAppObjectId = seedApplications[0].id;
export const mockSecondAppObjectId = seedApplications[1].id;
export const mockAppId = seedApplications[0].appId;
export const mockSecondAppId = seedApplications[1].appId;
export const mockUserId = seedUsers[0].id;
export const mockSecondUserId = seedUsers[1].id;
export const mockAppRoleId = seedApplications[0].appRoles[0].id;

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
