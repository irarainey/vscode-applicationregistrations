import * as vscode from "vscode";
import * as statusBarUtils from "../utils/status-bar";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { mockApplications, mockAppObjectId, seedMockData } from "./data/test-data";

// Create Jest mocks
jest.mock("vscode");
jest.mock("../repositories/graph-api-repository");

// Create the test suite for tree data provider
describe("Tree Data Provider Tests", () => {
	// Create instances of objects used in the tests
	const graphApiRepository = new GraphApiRepository();
	const treeDataProvider = new AppRegTreeDataProvider(graphApiRepository);

	// Create spy variables
	let triggerTreeErrorSpy: jest.SpyInstance<any, unknown[], any>;
    let statusUtilsSpy : jest.SpyInstance<void, [id: string], any>;

	// Create variables used in the tests
	let item: AppRegItem;

	beforeAll(() => {
		// Suppress console output
		console.error = jest.fn();
	});

	beforeEach(() => {
		// Reset mock data
		seedMockData();

		//Restore the default mock implementations
		jest.restoreAllMocks();

		// Define spies on the functions to be tested
        statusUtilsSpy = jest.spyOn(statusBarUtils, "clearStatusBarMessage");
		triggerTreeErrorSpy = jest.spyOn(Object.getPrototypeOf(treeDataProvider), "handleError");
	});

	afterAll(() => {
		treeDataProvider.dispose();
	});

	test("Create class instance", () => {
		// Assert class has been instantiated
		expect(treeDataProvider).toBeDefined();
	});

	test("Get sign in audience for single tenant", () => {
		// Arrange
		const audience = "AzureADMyOrg";

		// Act
		const result = treeDataProvider.getSignInAudienceDescription(audience);

		// Assert
		expect(result).toBe("Single Tenant");
	});

	test("Get sign in audience for multiple tenants", () => {
		// Arrange
		const audience = "AzureADMultipleOrgs";

		// Act
		const result = treeDataProvider.getSignInAudienceDescription(audience);

		// Assert
		expect(result).toBe("Multiple Tenants");
	});

	test("Get sign in audience for multiple tenants and consumers", () => {
		// Arrange
		const audience = "AzureADandPersonalMicrosoftAccount";

		// Act
		const result = treeDataProvider.getSignInAudienceDescription(audience);

		// Assert
		expect(result).toBe("Multiple Tenants and Personal Microsoft Accounts");
	});

	test("Get sign in audience for consumers", () => {
		// Arrange
		const audience = "PersonalMicrosoftAccount";

		// Act
		const result = treeDataProvider.getSignInAudienceDescription(audience);

		// Assert
		expect(result).toBe("Personal Microsoft Accounts");
	});

	test("Get sign in audience for unknown", () => {
		// Arrange
		const audience = "unknown";

		// Act
		const result = treeDataProvider.getSignInAudienceDescription(audience);

		// Assert
		expect(result).toBe("unknown");
	});

	test("Get tree item by context successfully", async () => {
		// Arrange
		await treeDataProvider.render();

		// Act
		const parent = treeDataProvider.getTreeItemApplicationParent({ objectId: mockAppObjectId });
		const result = treeDataProvider.getTreeItemChildByContext(parent, "AUDIENCE-PARENT");

		// Assert
		expect(result).toBeDefined();
		expect(result?.value).toEqual("AzureADMultipleOrgs");
	});

	test("Get tree item by context from child successfully", async () => {
		// Arrange
		await treeDataProvider.render();

		// Act
		const parent = treeDataProvider.getTreeItemApplicationParent({ objectId: mockAppObjectId });
		const result = treeDataProvider.getTreeItemChildByContext(parent, "AUDIENCE");

		// Assert
		expect(result).toBeDefined();
		expect(result?.value).toEqual("AzureADMultipleOrgs");
	});

	test("Get tree item by context unsuccessfully", async () => {
		// Arrange
		await treeDataProvider.render();

		// Act
		const parent = treeDataProvider.getTreeItemApplicationParent({ objectId: mockAppObjectId });
		const result = treeDataProvider.getTreeItemChildByContext(parent, "NOT-FOUND");

		// Assert
		expect(result).toBeUndefined();
	});

	test("Get all tree data for owned applications", async () => {
		// Arrange
		await treeDataProvider.render();

		// Act
		const result = treeDataProvider.getChildren(undefined);

		// Assert
		expect(result).toBeDefined();
		expect(result).toHaveLength(2);
		expect(treeDataProvider.isTreeEmpty).toBeFalsy();
	});

	test("Get all tree data for all applications", async () => {
		// Arrange
		jest.spyOn(vscode.workspace, "getConfiguration").mockImplementation(() => {
			return {
				get: (key: string) => {
					return key === "showOwnedApplicationsOnly" ? false : undefined;
				}
			} as any;
		});

		// Act
		await treeDataProvider.render();

		// Assert
		const result = treeDataProvider.getChildren(undefined);
		expect(result).toBeDefined();
		expect(result).toHaveLength(3);
		expect(treeDataProvider.isTreeEmpty).toBeFalsy();
	});

    test("Get all tree data for all applications with filter", async () => {
		// Arrange
		jest.spyOn(vscode.workspace, "getConfiguration").mockImplementation(() => {
			return {
				get: (key: string) => {
					if (key === "showOwnedApplicationsOnly") {
						return true;
					} else if (key === "useEventualConsistency") {
						return true;
					}
					return undefined;
				}
			} as any;
		});
        jest.spyOn(vscode.window, "showInputBox").mockResolvedValue("First");
        await treeDataProvider.render();

		// Act
		await treeDataProvider.filter();

		// Assert
		const result = treeDataProvider.getChildren(undefined);
		expect(result).toBeDefined();
		expect(result).toHaveLength(1);
		expect(treeDataProvider.isTreeEmpty).toBeFalsy();
	});

    test("Get all tree data for all applications with clear filter", async () => {
		// Arrange
		jest.spyOn(vscode.workspace, "getConfiguration").mockImplementation(() => {
			return {
				get: (key: string) => {
					if (key === "showOwnedApplicationsOnly") {
						return true;
					} else if (key === "useEventualConsistency") {
						return true;
					}
					return undefined;
				}
			} as any;
		});
        jest.spyOn(vscode.window, "showInputBox")
            .mockResolvedValue("")
            .mockResolvedValueOnce("First");
		await treeDataProvider.filter();

		// Act
		await treeDataProvider.filter();

		// Assert
		const result = treeDataProvider.getChildren(undefined);
		expect(result).toBeDefined();
		expect(result).toHaveLength(2);
		expect(treeDataProvider.isTreeEmpty).toBeFalsy();
	});

    test("Get all tree data for all applications with no filter and empty tree", async () => {
		// Arrange
		jest.spyOn(vscode.workspace, "getConfiguration").mockImplementation(() => {
			return {
				get: (key: string) => {
					if (key === "showOwnedApplicationsOnly") {
						return true;
					} else if (key === "useEventualConsistency") {
						return true;
					}
					return undefined;
				}
			} as any;
		});
        await treeDataProvider.render("Status message", "EMPTY");

		// Act
		await treeDataProvider.filter();

		// Assert
		const result = treeDataProvider.getChildren(undefined);
		expect(result).toBeDefined();
		expect(result).toHaveLength(1);
		expect(treeDataProvider.isTreeEmpty).toBeTruthy();
	});

    test("Get all tree data for all applications with filter but hit escape", async () => {
		// Arrange
		jest.spyOn(vscode.workspace, "getConfiguration").mockImplementation(() => {
			return {
				get: (key: string) => {
					if (key === "showOwnedApplicationsOnly") {
						return true;
					} else if (key === "useEventualConsistency") {
						return true;
					}
					return undefined;
				}
			} as any;
		});
        jest.spyOn(vscode.window, "showInputBox").mockResolvedValue(undefined);
        await treeDataProvider.render();

		// Act
		await treeDataProvider.filter();

		// Assert
		const result = treeDataProvider.getChildren(undefined);
		expect(result).toBeDefined();
		expect(result).toHaveLength(2);
		expect(treeDataProvider.isTreeEmpty).toBeFalsy();
	});

    test("Get all tree data for all applications with filter but eventual consistency not enabled", async () => {
		// Arrange
		jest.spyOn(vscode.workspace, "getConfiguration").mockImplementation(() => {
			return {
				get: (key: string) => {
					if (key === "showOwnedApplicationsOnly") {
						return true;
					} else if (key === "useEventualConsistency") {
						return false;
					}
					return undefined;
				}
			} as any;
		});
        const informationSpy = jest.spyOn(vscode.window, "showInformationMessage");
        await treeDataProvider.render();
        
		// Act
		await treeDataProvider.filter();

		// Assert
        expect(informationSpy).toHaveBeenCalledWith("The application list cannot be filtered when not using eventual consistency. Please enable this in user settings first.", "OK");
	});

	test("Get all tree data for all applications but with error", async () => {
		// Arrange
		const error = new Error("Get all tree data for all applications but with error");
		jest.spyOn(vscode.workspace, "getConfiguration").mockImplementation(() => {
			return {
				get: (key: string) => {
					return key === "showOwnedApplicationsOnly" ? false : undefined;
				}
			} as any;
		});
		jest.spyOn(graphApiRepository, "getApplicationListAll").mockImplementation(async (_filter: string | undefined) => ({ success: false, error }));

		// Act
		await treeDataProvider.render();

		// Assert
		expect(triggerTreeErrorSpy).toHaveBeenCalledWith(error);
	});

	test("Get all tree data for owned applications but with application not found error", async () => {
		// Arrange
		jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementation(async (_id: string, _select: string, _expandOwners: boolean | undefined) => {
            // eslint-disable-next-line no-throw-literal
            throw { message: "custom-error", code: "Request_ResourceNotFound" };
		});
		const consoleSpy = jest.spyOn(console, "log").mockImplementation(() => {
			return;
		});

		// Act
		await treeDataProvider.render();

		// Assert
		expect(consoleSpy).toHaveBeenCalledWith("Application no longer exists.");
	});

	test("Get count for owned applications but with credential error", async () => {
		// Arrange
		jest.spyOn(graphApiRepository, "getApplicationCountOwned").mockImplementation(async () => {
            // eslint-disable-next-line no-throw-literal
            throw { message: "custom-error", code: "CredentialUnavailableError" };
		});
		jest.spyOn(vscode.workspace, "getConfiguration").mockImplementation(() => {
			return {
				get: (key: string) => {
					if (key === "showOwnedApplicationsOnly") {
						return true;
					} else if (key === "showApplicationCountWarning") {
						return true;
					}
					return undefined;
				}
			} as any;
		});

		// Act
		await treeDataProvider.render("credential-error");

		// Assert
		expect(statusUtilsSpy).toHaveBeenCalled();
	});

	test("Get all tree data with owned applications but with error getting application count", async () => {
		// Arrange
		const error = new Error("Get all tree data with owned applications but with error getting application count");
		jest.spyOn(graphApiRepository, "getApplicationCountOwned").mockImplementation(async () => ({ success: false, error }));
		jest.spyOn(vscode.workspace, "getConfiguration").mockImplementation(() => {
			return {
				get: (key: string) => {
					if (key === "showOwnedApplicationsOnly") {
						return true;
					} else if (key === "showApplicationCountWarning") {
						return true;
					}
					return undefined;
				}
			} as any;
		});

		// Act
		await treeDataProvider.render();

		// Assert
		expect(triggerTreeErrorSpy).toHaveBeenCalledWith(error);
	});

	test("Get all tree data with all applications but with error getting application count", async () => {
		// Arrange
		const error = new Error("Get all tree data with all applications but with error getting application count");
		jest.spyOn(graphApiRepository, "getApplicationCountAll").mockImplementation(async () => ({ success: false, error }));
		jest.spyOn(vscode.workspace, "getConfiguration").mockImplementation(() => {
			return {
				get: (key: string) => {
					if (key === "showOwnedApplicationsOnly") {
						return false;
					} else if (key === "showApplicationCountWarning") {
						return true;
					}
					return undefined;
				}
			} as any;
		});

		// Act
		await treeDataProvider.render();

		// Assert
		expect(triggerTreeErrorSpy).toHaveBeenCalledWith(error);
	});

	test("Get all tree data and show owned applications with count warning", async () => {
		// Arrange
		const showWarningSpy = jest.spyOn(vscode.window, "showWarningMessage");
		jest.spyOn(graphApiRepository, "getApplicationCountOwned").mockImplementation(async () => ({ success: true, value: 201 }));
		jest.spyOn(vscode.workspace, "getConfiguration").mockImplementation(() => {
			return {
				get: (key: string) => {
					if (key === "showOwnedApplicationsOnly") {
						return true;
					} else if (key === "showApplicationCountWarning") {
						return true;
					} else if (key === "useEventualConsistency") {
						return false;
					}
					return undefined;
				}
			} as any;
		});

		// Act
		await treeDataProvider.render();

		// Assert
		expect(showWarningSpy).toHaveBeenCalledWith(`You do not have enabled eventual consistency enabled for Graph API calls and have 201 applications in your tenant. You would likely benefit from enabling eventual consistency in user settings. Would you like to do this now?`, "Yes", "No", "Disable Warning");
	});

	test("Get all tree data and show all applications with count warning", async () => {
		// Arrange
		const showWarningSpy = jest.spyOn(vscode.window, "showWarningMessage");
		jest.spyOn(graphApiRepository, "getApplicationCountAll").mockImplementation(async () => ({ success: true, value: 201 }));
		jest.spyOn(vscode.workspace, "getConfiguration").mockImplementation(() => {
			return {
				get: (key: string) => {
					if (key === "showOwnedApplicationsOnly") {
						return false;
					} else if (key === "showApplicationCountWarning") {
						return true;
					} else if (key === "useEventualConsistency") {
						return false;
					}
					return undefined;
				}
			} as any;
		});

		// Act
		await treeDataProvider.render();

		// Assert
		expect(showWarningSpy).toHaveBeenCalledWith(`You do not have enabled eventual consistency enabled for Graph API calls and have 201 applications in your tenant. You would likely benefit from enabling eventual consistency in user settings. Would you like to do this now?`, "Yes", "No", "Disable Warning");
	});

	test("Get all tree data and show all applications with count warning", async () => {
		// Arrange
		const showWarningSpy = jest.spyOn(vscode.window, "showWarningMessage");
		jest.spyOn(vscode.workspace, "getConfiguration").mockImplementation(() => {
			return {
				get: (key: string) => {
					if (key === "showOwnedApplicationsOnly") {
						return false;
					} else if (key === "showApplicationCountWarning") {
						return true;
					} else if (key === "useEventualConsistency") {
						return true;
					}
					return undefined;
				}
			} as any;
		});

		// Act
		await treeDataProvider.render();

		// Assert
		expect(showWarningSpy).toHaveBeenCalledWith(`You have enabled eventual consistency for Graph API calls but only have 3 applications in your tenant. You would likely benefit from disabling eventual consistency in user settings. Would you like to do this now?`, "Yes", "No", "Disable Warning");
	});

	test("Get tree when empty calling render EMPTY", async () => {
		// Act
		await treeDataProvider.render("Status message", "EMPTY");

		// Assert
		const result = treeDataProvider.getChildren(undefined);
		expect(result).toBeDefined();
		expect(result).toHaveLength(1);
		expect(treeDataProvider.isTreeEmpty).toBeTruthy();
	});

	test("Get tree when empty calling with no applications", async () => {
		// Arrange
		mockApplications.splice(0, 2);

		// Act
		await treeDataProvider.render();

		// Assert
		const result = treeDataProvider.getChildren(undefined);
		expect(result).toBeDefined();
		expect(result).toHaveLength(1);
		expect(treeDataProvider.isTreeEmpty).toBeTruthy();
	});

	test("Render sign in", async () => {
		// Act
		await treeDataProvider.render("Status message", "SIGN-IN");

		// Assert
		const result = treeDataProvider.getChildren(undefined);
		expect(result).toBeDefined();
		expect(result).toHaveLength(1);
	});

	test("Get children of a specified element", () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "REDIRECT-PARENT", children: [{ objectId: mockAppObjectId, label: "child" }] };

		// Act
		const result = treeDataProvider.getChildren(item);

		// Assert
		expect(result).toBeDefined();
		expect(result).toHaveLength(1);
	});
});
