import * as vscode from "vscode";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { RequiredResourceAccessService } from "../services/required-resource-access";
import { mockApplications, mockAppObjectId, mockGraphApiAppId, seedMockData } from "./data/test-data";
import { getTopLevelTreeItem } from "./test-utils";

// Create Jest mocks
jest.mock("vscode");
jest.mock("../repositories/graph-api-repository");

// Create the test suite for required resource access service
describe("Required Resource Access Service Tests", () => {
	// Create instances of objects used in the tests
	const graphApiRepository = new GraphApiRepository();
	const treeDataProvider = new AppRegTreeDataProvider(graphApiRepository);
	const requiredResourceAccessService = new RequiredResourceAccessService(graphApiRepository, treeDataProvider);

	// Create spy variables
	let triggerCompleteSpy: jest.SpyInstance<any, unknown[], any>;
	let triggerErrorSpy: jest.SpyInstance<any, unknown[], any>;
	let triggerTreeErrorSpy: jest.SpyInstance<any, unknown[], any>;
	let statusBarSpy: jest.SpyInstance<any, [text: string], any>;
	let iconSpy: jest.SpyInstance<any, [id: string, color?: any | undefined], any>;

	// Create variables used in the tests
	let item: AppRegItem;

	beforeAll(async () => {
		// Suppress console output
		console.error = jest.fn();
	});

	beforeEach(() => {
		// Reset mock data
		seedMockData();

		//Restore the default mock implementations
		jest.restoreAllMocks();

		// Define a standard mock implementation for the showWarningMessage function
		vscode.window.showWarningMessage = jest.fn().mockResolvedValue("Yes");

		// Define spies on the functions to be tested
		statusBarSpy = jest.spyOn(vscode.window, "setStatusBarMessage");
		iconSpy = jest.spyOn(vscode, "ThemeIcon");
		triggerCompleteSpy = jest.spyOn(Object.getPrototypeOf(requiredResourceAccessService), "triggerRefresh");
		triggerErrorSpy = jest.spyOn(Object.getPrototypeOf(requiredResourceAccessService), "handleError");
		triggerTreeErrorSpy = jest.spyOn(Object.getPrototypeOf(treeDataProvider), "handleError");

		// The item to be tested
		item = { objectId: mockAppObjectId, contextValue: "API-PERMISSIONS-APP", resourceAppId: mockGraphApiAppId };
	});

	afterAll(() => {
		// Dispose of the application service
		requiredResourceAccessService.dispose();
	});

	test("Create class instance", () => {
		// Assert class has been instantiated
		expect(requiredResourceAccessService).toBeDefined();
	});

	test("Remove all scopes for an api app successfully", async () => {
		// Act
		await requiredResourceAccessService.removeApi(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "API-PERMISSIONS");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children?.length).toEqual(1);
	});

	test("Remove all scopes for an api app with no existing scopes error", async () => {
		// Arrange
		const error = new Error("Remove all scopes for an api app with no existing scopes error");
		const graphSpy = jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementation(async (_id: string) => ({ success: false, error }));

		// Act
		await requiredResourceAccessService.removeApi(item);

		// Assert
		expect(graphSpy).toHaveBeenCalled();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
	});

	test("Remove all scopes for an api app with update error", async () => {
		// Arrange
		const error = new Error("Remove all scopes for an api app with update error");
		const graphSpy = jest.spyOn(graphApiRepository, "updateApplication").mockImplementation(async (_id: string) => ({ success: false, error }));

		// Act
		await requiredResourceAccessService.removeApi(item);

		// Assert
		expect(graphSpy).toHaveBeenCalled();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
	});

	test("Remove one scope for an api app successfully", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "API-PERMISSIONS-SCOPE", resourceAppId: mockGraphApiAppId, resourceScopeId: "570282fd-fa5c-430d-a7fd-fc8dc98a9dca" };

		// Act
		await requiredResourceAccessService.remove(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "API-PERMISSIONS");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children![0].children!.length).toEqual(2);
	});

	test("Remove last scope for an api app successfully", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "API-PERMISSIONS-SCOPE", resourceAppId: "1173dc06-edfc-40fc-980c-ff7d7ceb144d", resourceScopeId: "ef080234-1cd3-4945-9462-0d8901fd327a" };
		mockApplications[0].requiredResourceAccess![1].resourceAccess.splice(mockApplications[0].requiredResourceAccess![1].resourceAccess.length - 1, 1);
		mockApplications[0].requiredResourceAccess![1].resourceAccess.splice(mockApplications[0].requiredResourceAccess![1].resourceAccess.length - 1, 1);

		// Act
		await requiredResourceAccessService.remove(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "API-PERMISSIONS");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children!.length).toEqual(1);
	});

	test("Remove one scope for an api app with no existing scopes error", async () => {
		// Arrange
		const error = new Error("Remove one scope for an api app with no existing scopes error");
		const graphSpy = jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementation(async (_id: string) => ({ success: false, error }));
		item = { objectId: mockAppObjectId, contextValue: "API-PERMISSIONS-SCOPE", resourceAppId: mockGraphApiAppId, resourceScopeId: "570282fd-fa5c-430d-a7fd-fc8dc98a9dca" };

		// Act
		await requiredResourceAccessService.remove(item);

		// Assert
		expect(graphSpy).toHaveBeenCalled();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
	});

	test("Remove one scope for an api app with update error", async () => {
		// Arrange
		const error = new Error("Remove one scope for an api app with update error");
		const graphSpy = jest.spyOn(graphApiRepository, "updateApplication").mockImplementation(async (_id: string) => ({ success: false, error }));
		item = { objectId: mockAppObjectId, contextValue: "API-PERMISSIONS-SCOPE", resourceAppId: mockGraphApiAppId, resourceScopeId: "570282fd-fa5c-430d-a7fd-fc8dc98a9dca" };

		// Act
		await requiredResourceAccessService.remove(item);

		// Assert
		expect(graphSpy).toHaveBeenCalled();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
	});

	test("Add delegated scope successfully to existing api app", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "API-PERMISSIONS-APP", resourceAppId: mockGraphApiAppId };
		const quickPickSpy = jest.spyOn(vscode.window, "showQuickPick")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce({ label: "Delegated permissions", value: "Scope" } as any)
			.mockResolvedValueOnce({ label: "Calendars.Read", value: "465a38f9-76ea-45b9-9f34-9e8b0d4b0b42" } as any);

		// Act
		await requiredResourceAccessService.addToExisting(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "API-PERMISSIONS");
		expect(quickPickSpy).toHaveBeenCalled();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children![0].children!.length).toEqual(4);
		expect(treeItem?.children![0].children![3].label).toEqual("Delegated: Calendars.Read");
	});

	test("Add application role successfully to existing api app", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "API-PERMISSIONS-APP", resourceAppId: mockGraphApiAppId };
		const quickPickSpy = jest.spyOn(vscode.window, "showQuickPick")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce({ label: "Application permissions", value: "Role" } as any)
			.mockResolvedValueOnce({ label: "Mail.Send", value: "b633e1c5-b582-4048-a93e-9f11b44c7e96" } as any);

		// Act
		await requiredResourceAccessService.addToExisting(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "API-PERMISSIONS");
		expect(quickPickSpy).toHaveBeenCalled();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children![0].children!.length).toEqual(4);
		expect(treeItem?.children![0].children![3].label).toEqual("Application: Mail.Send");
	});

	test("Add scope user pressed escape on permission type", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "API-PERMISSIONS-APP", resourceAppId: mockGraphApiAppId };
		const quickPickSpy = jest.spyOn(vscode.window, "showQuickPick")
			.mockResolvedValue(undefined);

		// Act
		await requiredResourceAccessService.addToExisting(item);

		// Assert
		expect(quickPickSpy).toHaveBeenCalled();
	});

	test("Add delegated scope no service principal found error", async () => {
		// Arrange
		const error = new Error("Add delegated scope no service principal found error");
		const graphSpy = jest.spyOn(graphApiRepository, "findServicePrincipalByAppId").mockImplementation(async (_id: string) => ({ success: false, error }));
		item = { objectId: mockAppObjectId, contextValue: "API-PERMISSIONS-APP", resourceAppId: mockGraphApiAppId };
		const quickPickSpy = jest.spyOn(vscode.window, "showQuickPick")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce({ label: "Delegated permissions", value: "Scope" } as any);

		// Act
		await requiredResourceAccessService.addToExisting(item);

		// Assert
		expect(graphSpy).toHaveBeenCalled();
		expect(quickPickSpy).toHaveBeenCalled();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
	});

	test("Add delegated scope no scopes returned error", async () => {
		// Arrange
		const error = new Error("Add delegated scope no scopes returned error");
		const servicePrincipalSpy = jest.spyOn(graphApiRepository, "findServicePrincipalByAppId");
		const graphSpy = jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementation(async (_id: string) => ({ success: false, error }));
		item = { objectId: mockAppObjectId, contextValue: "API-PERMISSIONS-APP", resourceAppId: mockGraphApiAppId };
		const quickPickSpy = jest.spyOn(vscode.window, "showQuickPick")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce({ label: "Delegated permissions", value: "Scope" } as any);

		// Act
		await requiredResourceAccessService.addToExisting(item);

		// Assert
		expect(servicePrincipalSpy).toHaveBeenCalled();
		expect(quickPickSpy).toHaveBeenCalled();
		expect(graphSpy).toHaveBeenCalled();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
	});

	test("Add delegated scope no scopes unassigned scopes warning", async () => {
		// Arrange
		const servicePrincipalSpy = jest.spyOn(graphApiRepository, "findServicePrincipalByAppId");
		const graphSpy = jest.spyOn(graphApiRepository, "getApplicationDetailsPartial");
		const showInformationSpy = jest.spyOn(vscode.window, "showInformationMessage");
		item = { objectId: mockAppObjectId, contextValue: "API-PERMISSIONS-APP", resourceAppId: "1173dc06-edfc-40fc-980c-ff7d7ceb144d" };
		const quickPickSpy = jest.spyOn(vscode.window, "showQuickPick")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce({ label: "Delegated permissions", value: "Scope" } as any);

		// Act
		await requiredResourceAccessService.addToExisting(item);

		// Assert
		expect(servicePrincipalSpy).toHaveBeenCalled();
		expect(quickPickSpy).toHaveBeenCalled();
		expect(graphSpy).toHaveBeenCalled();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(showInformationSpy).toHaveBeenCalledWith("There are no user delegated permissions available to add to this application registration.", "OK");
	});

	test("Add delegated scope user pressed cancel on scope selection", async () => {
		// Arrange
		const servicePrincipalSpy = jest.spyOn(graphApiRepository, "findServicePrincipalByAppId");
		const graphSpy = jest.spyOn(graphApiRepository, "getApplicationDetailsPartial");
		item = { objectId: mockAppObjectId, contextValue: "API-PERMISSIONS-APP", resourceAppId: mockGraphApiAppId };
		const quickPickSpy = jest.spyOn(vscode.window, "showQuickPick")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce({ label: "Delegated permissions", value: "Scope" } as any);

		// Act
		await requiredResourceAccessService.addToExisting(item);

		// Assert
		expect(servicePrincipalSpy).toHaveBeenCalled();
		expect(graphSpy).toHaveBeenCalled();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(quickPickSpy).toHaveBeenCalled();
	});

	test("Add application scope no scopes unassigned scopes warning", async () => {
		// Arrange
		const servicePrincipalSpy = jest.spyOn(graphApiRepository, "findServicePrincipalByAppId");
		const graphSpy = jest.spyOn(graphApiRepository, "getApplicationDetailsPartial");
		const showInformationSpy = jest.spyOn(vscode.window, "showInformationMessage");
		item = { objectId: mockAppObjectId, contextValue: "API-PERMISSIONS-APP", resourceAppId: "1173dc06-edfc-40fc-980c-ff7d7ceb144d" };
		const quickPickSpy = jest.spyOn(vscode.window, "showQuickPick")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce({ label: "Application permissions", value: "Role" } as any);

		// Act
		await requiredResourceAccessService.addToExisting(item);

		// Assert
		expect(servicePrincipalSpy).toHaveBeenCalled();
		expect(quickPickSpy).toHaveBeenCalled();
		expect(graphSpy).toHaveBeenCalled();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(showInformationSpy).toHaveBeenCalledWith("There are no application permissions available to add to this application registration.", "OK");
	});

	test("Add application scope user pressed cancel on scope selection", async () => {
		// Arrange
		const servicePrincipalSpy = jest.spyOn(graphApiRepository, "findServicePrincipalByAppId");
		const graphSpy = jest.spyOn(graphApiRepository, "getApplicationDetailsPartial");
		item = { objectId: mockAppObjectId, contextValue: "API-PERMISSIONS-APP", resourceAppId: mockGraphApiAppId };
		const quickPickSpy = jest.spyOn(vscode.window, "showQuickPick")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce({ label: "Application permissions", value: "Role" } as any);

		// Act
		await requiredResourceAccessService.addToExisting(item);

		// Assert
		expect(servicePrincipalSpy).toHaveBeenCalled();
		expect(graphSpy).toHaveBeenCalled();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(quickPickSpy).toHaveBeenCalled();
	});

	test("Add delegated scope successfully from a new api app", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "API-PERMISSIONS-APP" };
		const inputSpy = jest.spyOn(vscode.window, "showInputBox")
			.mockResolvedValue("Random Api");
		const quickPickSpy = jest.spyOn(vscode.window, "showQuickPick")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce({ label: "Random Api", value: "0cad2264-23c3-4366-8c28-805aaeda257b" } as any)
			.mockResolvedValueOnce({ label: "Delegated permissions", value: "Scope" } as any)
			.mockResolvedValueOnce({ label: "Do.Something", value: "82062f02-7837-45e6-a497-4938246ceb5c" } as any);

		// Act
		await requiredResourceAccessService.add(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "API-PERMISSIONS");
		expect(inputSpy).toHaveBeenCalled();
		expect(quickPickSpy).toHaveBeenCalled();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children!.length).toEqual(3);
		expect(treeItem?.children![2].label).toEqual("Random Api");
		expect(treeItem?.children![2].children![0].label).toEqual("Delegated: Do.Something");
	});

	test("Add delegated scope from a new api app user pressed escape on API search", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "API-PERMISSIONS-APP" };
		const inputSpy = jest.spyOn(vscode.window, "showInputBox")
			.mockResolvedValue(undefined);

		// Act
		await requiredResourceAccessService.add(item);

		// Assert
		expect(inputSpy).toHaveBeenCalled();
	});

	test("Add delegated scope with new api graph error", async () => {
		// Arrange
		const error = new Error("Add delegated scope with new api graph error");
		item = { objectId: mockAppObjectId, contextValue: "API-PERMISSIONS-APP" };
		const graphSpy = jest.spyOn(graphApiRepository, "findServicePrincipalsByDisplayName").mockImplementation(async (_id: string) => ({ success: false, error }));
		const inputSpy = jest.spyOn(vscode.window, "showInputBox")
			.mockResolvedValue("Random Api");

		// Act
		await requiredResourceAccessService.add(item);

		// Assert
		expect(graphSpy).toHaveBeenCalled();
		expect(inputSpy).toHaveBeenCalled();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
	});

	test("Add delegated scope with new api no service principals found", async () => {
		// Arrange
		const error = new Error("Add delegated scope with new api graph error");
		item = { objectId: mockAppObjectId, contextValue: "API-PERMISSIONS-APP" };
		const graphSpy = jest.spyOn(graphApiRepository, "findServicePrincipalsByDisplayName").mockImplementation(async (_id: string) => ({ success: true, value: [] }));
		const inputSpy = jest.spyOn(vscode.window, "showInputBox")
			.mockResolvedValue("Random Api");
		const informationMessageSpy = jest.spyOn(vscode.window, "showInformationMessage");

		// Act
		await requiredResourceAccessService.add(item);

		// Assert
		expect(graphSpy).toHaveBeenCalled();
		expect(inputSpy).toHaveBeenCalled();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(informationMessageSpy).toHaveBeenCalledWith("No API Applications were found that match the search criteria.", "OK");
	});

	test("Add delegated scope from a new api app user pressed escape on API selection", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "API-PERMISSIONS-APP" };
		const inputSpy = jest.spyOn(vscode.window, "showInputBox")
			.mockResolvedValue("Random Api");
		const quickPickSpy = jest.spyOn(vscode.window, "showQuickPick")
			.mockResolvedValue(undefined);

		// Act
		await requiredResourceAccessService.add(item);

		// Assert
		expect(inputSpy).toHaveBeenCalled();
		expect(quickPickSpy).toHaveBeenCalled();
	});

	test("Error getting required resource access children", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "API-PERMISSIONS" };
		const error = new Error("Error getting required resource access children");
		jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementation(async (id: string, select: string) => {
			if (select === "requiredResourceAccess") {
				return { success: false, error };
			}
			return mockApplications.find((app) => app.id === id);
		});

		// Act
		await treeDataProvider.render();
		await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "API-PERMISSIONS");

		// Assert
		expect(triggerTreeErrorSpy).toHaveBeenCalledWith(error);
	});	

	test("Error finding service principal when accessing children", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "API-PERMISSIONS" };
		const error = new Error("Error finding service principal when accessing children");
		jest.spyOn(graphApiRepository, "findServicePrincipalByAppId").mockImplementation(async (id: string) => ({ success: false, error }));

		// Act
		await treeDataProvider.render();
		await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "API-PERMISSIONS");

		// Assert
		expect(triggerTreeErrorSpy).toHaveBeenCalledWith(error);
	});	
});
