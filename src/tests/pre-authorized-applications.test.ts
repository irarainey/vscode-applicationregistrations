import * as vscode from "vscode";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { mockAppId, mockApplications, mockAppObjectId, mockDelegatedPermissionId, mockPreAuthorizedAppId, seedMockData } from "./data/test-data";
import { PreAuthorizedApplicationsService } from "../services/pre-authorized-applications";
import { getTopLevelTreeItem } from "./test-utils";

// Create Jest mocks
jest.mock("vscode");
jest.mock("../repositories/graph-api-repository");

// Create the test suite for pre authorized client application service
describe("Pre Authorized Client Application Service Tests", () => {
	// Create instances of objects used in the tests
	const graphApiRepository = new GraphApiRepository();
	const treeDataProvider = new AppRegTreeDataProvider(graphApiRepository);
	const preAuthorizedApplicationsService = new PreAuthorizedApplicationsService(graphApiRepository, treeDataProvider);

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
		triggerCompleteSpy = jest.spyOn(Object.getPrototypeOf(preAuthorizedApplicationsService), "triggerRefresh");
		triggerErrorSpy = jest.spyOn(Object.getPrototypeOf(preAuthorizedApplicationsService), "handleError");
		triggerTreeErrorSpy = jest.spyOn(Object.getPrototypeOf(treeDataProvider), "handleError");

		// The item to be tested
		item = { objectId: mockAppObjectId, contextValue: "EXPOSED-API-PERMISSIONS" };
	});

	afterAll(() => {
		preAuthorizedApplicationsService.dispose();
		treeDataProvider.dispose();
	});

	test("Create class instance", () => {
		// Assert class has been instantiated
		expect(preAuthorizedApplicationsService).toBeDefined();
	});

	test("Remove authorised client application successfully", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "AUTHORIZED-CLIENT", label: "Second Test App", resourceAppId: mockPreAuthorizedAppId };
		const warningSpy = jest.spyOn(vscode.window, "showWarningMessage");

		// Act
		await preAuthorizedApplicationsService.removeAuthorisedClient(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(warningSpy).toHaveBeenCalledWith(`Do you want to remove the Authorized Client Application ${item.label}?`, "Yes", "No");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children![0].children).toBeUndefined();
	});

	test("Remove authorised client application with update error", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "AUTHORIZED-CLIENT", label: "Second Test App", resourceAppId: mockPreAuthorizedAppId };
		const error = new Error("Remove authorised client application with update error");
		const graphSpy = jest.spyOn(graphApiRepository, "updateApplication").mockImplementation(async (_id: string) => ({ success: false, error }));

		// Act
		await preAuthorizedApplicationsService.removeAuthorisedClient(item);

		// Assert
		expect(graphSpy).toHaveBeenCalled();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalledWith(error, undefined);
	});

	test("Remove authorised client application with get api error", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "AUTHORIZED-CLIENT", label: "Second Test App", resourceAppId: mockPreAuthorizedAppId };
		const error = new Error("Remove authorised client application with get api error");
		const graphSpy = jest
			.spyOn(graphApiRepository, "getApplicationDetailsPartial")
			.mockImplementation(async (_id: string) => ({ success: false, error }));

		// Act
		await preAuthorizedApplicationsService.removeAuthorisedClient(item);

		// Assert
		expect(graphSpy).toHaveBeenCalled();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
	});

	test("Remove one scope from an authorised client application successfully", async () => {
		// Arrange
		item = {
			objectId: mockAppObjectId,
			contextValue: "AUTHORIZED-CLIENT-SCOPE",
			resourceAppId: mockPreAuthorizedAppId,
			value: mockDelegatedPermissionId,
			label: "Sample.Two"
		};

		// Act
		await preAuthorizedApplicationsService.removeAuthorisedClientScope(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children![0].children![0].children!.length).toEqual(1);
	});

	test("Remove last scope from an authorised client application successfully", async () => {
		// Arrange
		item = {
			objectId: mockAppObjectId,
			contextValue: "AUTHORIZED-CLIENT-SCOPE",
			resourceAppId: mockPreAuthorizedAppId,
			value: mockDelegatedPermissionId,
			label: "Sample.Two"
		};
		mockApplications[0].api!.preAuthorizedApplications![0].delegatedPermissionIds.splice(1, 1);

		// Act
		await preAuthorizedApplicationsService.removeAuthorisedClientScope(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children![0].children).toBeUndefined();
	});

	test("Remove one scope from an authorised client application with get api error", async () => {
		// Arrange
		item = {
			objectId: mockAppObjectId,
			contextValue: "AUTHORIZED-CLIENT-SCOPE",
			resourceAppId: mockPreAuthorizedAppId,
			value: mockDelegatedPermissionId,
			label: "Sample.Two"
		};
		const error = new Error("Remove one scope from an authorised client application with get api error");
		const graphSpy = jest
			.spyOn(graphApiRepository, "getApplicationDetailsPartial")
			.mockImplementation(async (_id: string) => ({ success: false, error }));

		// Act
		await preAuthorizedApplicationsService.removeAuthorisedClientScope(item);

		// Assert
		expect(graphSpy).toHaveBeenCalled();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
	});

	test("Add scope successfully to existing authorized client application", async () => {
		// Arrange
		item = {
			objectId: mockAppObjectId,
			contextValue: "AUTHORIZED-CLIENT",
			label: "Second Test App",
			resourceAppId: mockPreAuthorizedAppId,
			value: mockAppId
		};
		const quickPickSpy = jest
			.spyOn(vscode.window, "showQuickPick")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce({ label: "Sample.One", value: "2af52627-d1b4-408e-b188-ccca2a5cd33c" } as any);

		// Act
		await preAuthorizedApplicationsService.addToExistingAuthorisedClient(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(quickPickSpy).toHaveBeenCalled();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children![0].children![0].children!.length).toEqual(3);
		expect(treeItem?.children![0].children![0].children![2].label).toEqual("Sample.One");
	});

	test("Add scope to existing authorized client application with get api error", async () => {
		// Arrange
		item = {
			objectId: mockAppObjectId,
			contextValue: "AUTHORIZED-CLIENT",
			label: "Second Test App",
			resourceAppId: mockPreAuthorizedAppId,
			value: mockAppId
		};
		const error = new Error("Add scope to existing authorized client application with get api error");
		const graphSpy = jest
			.spyOn(graphApiRepository, "getApplicationDetailsPartial")
			.mockImplementation(async (_id: string) => ({ success: false, error }));

		// Act
		await preAuthorizedApplicationsService.addToExistingAuthorisedClient(item);

		// Assert
		expect(graphSpy).toHaveBeenCalled();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
	});

	test("Add scope to existing authorized client application but escape pushed", async () => {
		// Arrange
		item = {
			objectId: mockAppObjectId,
			contextValue: "AUTHORIZED-CLIENT",
			label: "Second Test App",
			resourceAppId: mockPreAuthorizedAppId,
			value: mockAppId
		};
		const graphSpy = jest.spyOn(graphApiRepository, "updateApplication");
		const quickPickSpy = jest.spyOn(vscode.window, "showQuickPick").mockResolvedValue(undefined);

		// Act
		await preAuthorizedApplicationsService.addToExistingAuthorisedClient(item);

		// Assert
		expect(quickPickSpy).toHaveBeenCalled();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(graphSpy).not.toHaveBeenCalled();
	});

	test("Add scope to existing authorized client application but no available scopes", async () => {
		// Arrange
		item = {
			objectId: mockAppObjectId,
			contextValue: "AUTHORIZED-CLIENT",
			label: "Second Test App",
			resourceAppId: mockPreAuthorizedAppId,
			value: mockAppId
		};
		mockApplications[0].api!.oauth2PermissionScopes!.splice(0, 1);
		const informationMessageSpy = jest.spyOn(vscode.window, "showInformationMessage");

		// Act
		await preAuthorizedApplicationsService.addToExistingAuthorisedClient(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(informationMessageSpy).toHaveBeenCalledWith("There are no scopes available to add to this application.", "OK");
	});

	test("Add scope successfully for new authorised client application", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "AUTHORIZED-CLIENTS" };
		const inputSpy = jest.spyOn(vscode.window, "showInputBox").mockResolvedValue("Bert's Cookies API");
		const quickPickSpy = jest
			.spyOn(vscode.window, "showQuickPick")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce({ label: "Bert's Cookies API", value: "1173dc06-edfc-40fc-980c-ff7d7ceb144d" } as any)
			.mockResolvedValueOnce({ label: "Sample.One", value: "2af52627-d1b4-408e-b188-ccca2a5cd33c" } as any);

		// Act
		await preAuthorizedApplicationsService.addAuthorisedClientScope(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(inputSpy).toHaveBeenCalled();
		expect(quickPickSpy).toHaveBeenCalled();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children![0].children!.length).toEqual(2);
	});

	test("Add scope for new authorised client application but user escaped app search", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "AUTHORIZED-CLIENTS" };
		const inputSpy = jest.spyOn(vscode.window, "showInputBox").mockResolvedValue(undefined);

		// Act
		await preAuthorizedApplicationsService.addAuthorisedClientScope(item);

		// Assert
		expect(inputSpy).toHaveBeenCalled();
	});

	test("Add scope for new authorised client application but user escaped app selection", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "AUTHORIZED-CLIENTS" };
		const inputSpy = jest.spyOn(vscode.window, "showInputBox").mockResolvedValue("Bert's Cookies API");
		const quickPickSpy = jest.spyOn(vscode.window, "showQuickPick").mockResolvedValue(undefined);

		// Act
		await preAuthorizedApplicationsService.addAuthorisedClientScope(item);

		// Assert
		expect(inputSpy).toHaveBeenCalled();
		expect(quickPickSpy).toHaveBeenCalled();
	});

	test("Add scope for new authorised client application but find service principal error", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "AUTHORIZED-CLIENTS" };
		const error = new Error("Add scope for new authorised client application but find service principal error");
		const graphSpy = jest
			.spyOn(graphApiRepository, "findServicePrincipalsByDisplayName")
			.mockImplementation(async (_id: string) => ({ success: false, error }));
		const inputSpy = jest.spyOn(vscode.window, "showInputBox").mockResolvedValue("Bert's Cookies API");

		// Act
		await preAuthorizedApplicationsService.addAuthorisedClientScope(item);

		// Assert
		expect(inputSpy).toHaveBeenCalled();
		expect(graphSpy).toHaveBeenCalled();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
	});

	test("Add scope for new authorised client application but no service principals returned", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "AUTHORIZED-CLIENTS" };
		const inputSpy = jest.spyOn(vscode.window, "showInputBox").mockResolvedValue("Unknown App");
		const informationMessageSpy = jest.spyOn(vscode.window, "showInformationMessage");

		// Act
		await preAuthorizedApplicationsService.addAuthorisedClientScope(item);

		// Assert
		expect(inputSpy).toHaveBeenCalled();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(informationMessageSpy).toHaveBeenCalledWith("No Applications were found that match the search criteria.", "OK");
	});
});
