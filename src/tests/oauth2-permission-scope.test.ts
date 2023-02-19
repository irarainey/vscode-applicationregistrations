import * as vscode from "vscode";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { OAuth2PermissionScopeService } from "../services/oauth2-permission-scope";
import { mockApplications, mockAppObjectId, mockExposedApiId, seedMockData } from "./data/test-data";
import { getTopLevelTreeItem } from "./test-utils";
import exp = require("constants");

// Create Jest mocks
jest.mock("vscode");
jest.mock("../repositories/graph-api-repository");

// Create the test suite for oauth2 permission scope service
describe("OAuth2 Permission Scope Service Tests", () => {
	// Create instances of objects used in the tests
	const graphApiRepository = new GraphApiRepository();
	const treeDataProvider = new AppRegTreeDataProvider(graphApiRepository);
	const oauth2PermissionScopeService = new OAuth2PermissionScopeService(graphApiRepository, treeDataProvider);

	// Create spy variables
	let triggerCompleteSpy: jest.SpyInstance<any, unknown[], any>;
	let triggerErrorSpy: jest.SpyInstance<any, unknown[], any>;
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
		triggerCompleteSpy = jest.spyOn(Object.getPrototypeOf(oauth2PermissionScopeService), "triggerRefresh");
		triggerErrorSpy = jest.spyOn(Object.getPrototypeOf(oauth2PermissionScopeService), "handleError");

		// The item to be tested
		item = { objectId: mockAppObjectId, contextValue: "EXPOSED-API-PERMISSIONS" };
	});

	afterAll(() => {
		// Dispose of the application service
		oauth2PermissionScopeService.dispose();
	});

	test("Create class instance", () => {
		// Assert class has been instantiated
		expect(oauth2PermissionScopeService).toBeDefined();
	});

	test("Delete disabled exposed api permissions successfully", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SCOPE-DISABLED", value: mockExposedApiId, state: false };
		const warningSpy = jest.spyOn(vscode.window, "showWarningMessage");

		// Act
		await oauth2PermissionScopeService.delete(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(warningSpy).toHaveBeenCalledWith(`Do you want to delete the Scope ${item.label}?`, "Yes", "No");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children?.length).toEqual(1);
	});

	test("Delete disabled exposed api permissions but decline warning", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SCOPE-DISABLED", value: mockExposedApiId, state: false, label: "Sample Scope One" };
		const warningSpy = jest.spyOn(vscode.window, "showWarningMessage").mockResolvedValue("No" as any);

		// Act
		await oauth2PermissionScopeService.delete(item);

		// Assert
		treeDataProvider.render();
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(warningSpy).toHaveBeenCalledWith(`Do you want to delete the Scope ${item.label}?`, "Yes", "No");
		expect(treeItem?.children?.length).toEqual(2);
	});

	test("Delete disabled exposed api permissions but error getting existing scopes", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SCOPE-DISABLED", value: mockExposedApiId, state: false };
		const error = new Error("Delete disabled exposed api permissions but error getting existing scopes");
		const graphSpy = jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementation(async (_id: string) => ({ success: false, error }));
		const warningSpy = jest.spyOn(vscode.window, "showWarningMessage");

		// Act
		await oauth2PermissionScopeService.delete(item);

		// Assert
		expect(warningSpy).toHaveBeenCalledWith(`Do you want to delete the Scope ${item.label}?`, "Yes", "No");
		expect(graphSpy).toHaveBeenCalled();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
	});

	test("Delete enabled exposed api permissions but decline warning to disable first", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SCOPE-ENABLED", value: mockExposedApiId, state: true, label: "Sample Scope One" };
		const warningSpy = jest.spyOn(vscode.window, "showWarningMessage").mockResolvedValue("No" as any);

		// Act
		await oauth2PermissionScopeService.delete(item);

		// Assert
		expect(warningSpy).toHaveBeenCalledWith(`The Scope ${item.label} cannot be deleted unless it is disabled. Do you want to disable the scope and then delete it?`, "Yes", "No");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
	});

	test("Delete enabled exposed api permissions and accept warning to disable first", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SCOPE-ENABLED", value: mockExposedApiId, state: true, label: "Sample Scope One" };
		const warningSpy = jest.spyOn(vscode.window, "showWarningMessage").mockResolvedValue("Yes" as any);
		const changeStateSpy = jest.spyOn(oauth2PermissionScopeService, "changeState");

		// Act
		await oauth2PermissionScopeService.delete(item);

		// Assert
		expect(warningSpy).toHaveBeenCalledWith(`The Scope ${item.label} cannot be deleted unless it is disabled. Do you want to disable the scope and then delete it?`, "Yes", "No");
		expect(changeStateSpy).toHaveBeenCalledWith(item, false);
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children?.length).toEqual(1);
	});

	test("Disable enabled exposed api permission but no existing roles returned", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SCOPE-ENABLED", value: mockExposedApiId, state: true, label: "Sample Scope One" };
		const error = new Error("Disable enabled exposed api permission but no existing roles returned");
		const graphSpy = jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementationOnce(async () => ({ success: false, error: error }));

		// Act
		await oauth2PermissionScopeService.changeState(item, false);

		// Assert
		expect(graphSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
	});

	test("Disable enabled exposed api permission", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SCOPE-ENABLED", value: mockExposedApiId, state: true, label: "Sample Scope One" };

		// Act
		await oauth2PermissionScopeService.changeState(item, false);

		// Assert
		treeDataProvider.render();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(treeItem?.children?.[0].state).toEqual(false);
	});

	test("Enable disabled exposed api permission", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SCOPE-DISABLED", value: mockExposedApiId, state: false, label: "Sample Scope One" };

		// Act
		await oauth2PermissionScopeService.changeState(item, true);

		// Assert
		treeDataProvider.render();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(treeItem?.children?.[0].state).toEqual(true);
	});

	test("Edit value but no permissions returned", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SCOPE-ENABLED", value: mockExposedApiId, state: true, label: "Sample Scope One" };
		const error = new Error("Edit value but no permissions returned");
		const graphSpy = jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementationOnce(async () => ({ success: false, error: error }));

		// Act
		await oauth2PermissionScopeService.editValue(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(graphSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
	});

	test("Edit value but no app id uri set", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SCOPE-ENABLED", value: mockExposedApiId, state: true, label: "Sample Scope One" };
		const warningSpy = jest.spyOn(vscode.window, "showWarningMessage").mockResolvedValue("OK" as any);
		mockApplications[0].identifierUris = [];

		// Act
		await oauth2PermissionScopeService.editValue(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(warningSpy).toHaveBeenCalledWith("This application does not have an Application ID URI. Please add one before editing scopes.", "OK");
	});

	test("Edit display name", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SCOPE-ENABLED", value: mockExposedApiId, state: true, label: "Sample Scope One" };
		jest.spyOn(vscode.window, "showInputBox").mockResolvedValue("New Scope" as any);

		// Act
		await oauth2PermissionScopeService.editValue(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(treeItem?.children?.[0].label).toEqual("New Scope");
	});

	test("Edit display name but cancel input", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SCOPE-ENABLED", value: mockExposedApiId, state: true, label: "Sample Scope One" };
		jest.spyOn(vscode.window, "showInputBox").mockResolvedValue(undefined);

		// Act
		await oauth2PermissionScopeService.editValue(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(vscode.window.showInputBox).toHaveBeenCalled();
	});

	test("Edit scope value", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SCOPE-VALUE", value: mockExposedApiId, state: true, label: "Sample.One" };
		jest.spyOn(vscode.window, "showInputBox").mockResolvedValue("New.Scope" as any);

		// Act
		await oauth2PermissionScopeService.editValue(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(treeItem?.children![0].children![0].label).toEqual("Scope: New.Scope");
	});

	test("Edit scope value but cancel input", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SCOPE-VALUE", value: mockExposedApiId, state: true, label: "Sample.One" };
		jest.spyOn(vscode.window, "showInputBox").mockResolvedValue(undefined);

		// Act
		await oauth2PermissionScopeService.editValue(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(vscode.window.showInputBox).toHaveBeenCalled();
	});
	
	test("Edit scope description", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SCOPE-DESCRIPTION", value: mockExposedApiId, state: true, label: "Sample description one" };
		jest.spyOn(vscode.window, "showInputBox").mockResolvedValue("New Description" as any);

		// Act
		await oauth2PermissionScopeService.editValue(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(treeItem?.children![0].children![1].label).toEqual("Description: New Description");
	});

	test("Edit scope description but cancel input", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SCOPE-DESCRIPTION", value: mockExposedApiId, state: true, label: "Sample description one" };
		jest.spyOn(vscode.window, "showInputBox").mockResolvedValue(undefined);

		// Act
		await oauth2PermissionScopeService.editValue(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(vscode.window.showInputBox).toHaveBeenCalled();
	});

	test("Edit scope consent", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SCOPE-CONSENT", value: mockExposedApiId, state: true, label: "Consent: Admins and Users" };
		jest.spyOn(vscode.window, "showQuickPick").mockResolvedValue({ label: "Administrators only", value: "Admin" } as any);

		// Act
		await oauth2PermissionScopeService.editValue(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(treeItem?.children![0].children![2].label).toEqual("Consent: Admins Only");
	});

	test("Edit scope consent but cancel input", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SCOPE-CONSENT", value: mockExposedApiId, state: true, label: "Consent: Admins and Users" };
		jest.spyOn(vscode.window, "showQuickPick").mockResolvedValue(undefined);

		// Act
		await oauth2PermissionScopeService.editValue(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(vscode.window.showInputBox).toHaveBeenCalled();
	});

	test("Edit scope successfully", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SCOPE-ENABLED", value: mockExposedApiId, state: true, label: "Sample Scope One" };
		jest.spyOn(vscode.window, "showInputBox")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce("New.Scope" as any)
			.mockResolvedValueOnce("New Scope" as any)
			.mockResolvedValueOnce("New Description" as any)
			.mockResolvedValueOnce("New User Display Name" as any)
			.mockResolvedValueOnce("New User Description" as any);

		jest.spyOn(vscode.window, "showQuickPick")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce({ label: "Administrators only", value: "Admin" } as any)
			.mockResolvedValueOnce({ label: "Enabled", value: true } as any);

		// Act
		await oauth2PermissionScopeService.edit(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(treeItem?.children![0].children![0].label).toEqual("Scope: New.Scope");
		expect(treeItem?.children![0].children![1].label).toEqual("Description: New Description");
		expect(treeItem?.children![0].children![2].label).toEqual("Consent: Admins Only");
		expect(treeItem?.children![0].children![3].label).toEqual("Enabled: Yes");
	});

	test("Edit scope but cancel from scope input", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SCOPE-ENABLED", value: mockExposedApiId, state: true, label: "Sample Scope One" };
		jest.spyOn(vscode.window, "showInputBox")
			.mockResolvedValue(undefined);

		// Act
		await oauth2PermissionScopeService.edit(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(vscode.window.showInputBox).toHaveBeenCalled();
		treeDataProvider.render();
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(treeItem?.children![0].children![0].label).toEqual("Scope: Sample.One");
		expect(treeItem?.children![0].children![1].label).toEqual("Description: Sample description one");
		expect(treeItem?.children![0].children![2].label).toEqual("Consent: Admins and Users");
		expect(treeItem?.children![0].children![3].label).toEqual("Enabled: Yes");
	});

	test("Edit scope but cancel from consent input", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SCOPE-ENABLED", value: mockExposedApiId, state: true, label: "Sample Scope One" };
		jest.spyOn(vscode.window, "showInputBox")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce("New.Scope" as any);

		jest.spyOn(vscode.window, "showQuickPick")
			.mockResolvedValue(undefined);

		// Act
		await oauth2PermissionScopeService.edit(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(vscode.window.showInputBox).toHaveBeenCalled();
		treeDataProvider.render();
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(treeItem?.children![0].children![0].label).toEqual("Scope: Sample.One");
		expect(treeItem?.children![0].children![1].label).toEqual("Description: Sample description one");
		expect(treeItem?.children![0].children![2].label).toEqual("Consent: Admins and Users");
		expect(treeItem?.children![0].children![3].label).toEqual("Enabled: Yes");
	});

	test("Edit scope but cancel from display name input", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SCOPE-ENABLED", value: mockExposedApiId, state: true, label: "Sample Scope One" };
		jest.spyOn(vscode.window, "showInputBox")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce("New.Scope" as any);

		jest.spyOn(vscode.window, "showQuickPick")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce({ label: "Administrators only", value: "Admin" } as any);

		// Act
		await oauth2PermissionScopeService.edit(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(vscode.window.showInputBox).toHaveBeenCalled();
		treeDataProvider.render();
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(treeItem?.children![0].children![0].label).toEqual("Scope: Sample.One");
		expect(treeItem?.children![0].children![1].label).toEqual("Description: Sample description one");
		expect(treeItem?.children![0].children![2].label).toEqual("Consent: Admins and Users");
		expect(treeItem?.children![0].children![3].label).toEqual("Enabled: Yes");
	});

	test("Edit scope but cancel from description input", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SCOPE-ENABLED", value: mockExposedApiId, state: true, label: "Sample Scope One" };
		jest.spyOn(vscode.window, "showInputBox")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce("New.Scope" as any)
			.mockResolvedValueOnce("New Scope" as any);

		jest.spyOn(vscode.window, "showQuickPick")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce({ label: "Administrators only", value: "Admin" } as any);

		// Act
		await oauth2PermissionScopeService.edit(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(vscode.window.showInputBox).toHaveBeenCalled();
		treeDataProvider.render();
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(treeItem?.children![0].children![0].label).toEqual("Scope: Sample.One");
		expect(treeItem?.children![0].children![1].label).toEqual("Description: Sample description one");
		expect(treeItem?.children![0].children![2].label).toEqual("Consent: Admins and Users");
		expect(treeItem?.children![0].children![3].label).toEqual("Enabled: Yes");
	});

	test("Edit scope but cancel from user consent display name input", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SCOPE-ENABLED", value: mockExposedApiId, state: true, label: "Sample Scope One" };
		jest.spyOn(vscode.window, "showInputBox")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce("New.Scope" as any)
			.mockResolvedValueOnce("New Scope" as any)
			.mockResolvedValueOnce("New Description" as any);

		jest.spyOn(vscode.window, "showQuickPick")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce({ label: "Administrators only", value: "Admin" } as any);

		// Act
		await oauth2PermissionScopeService.edit(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(vscode.window.showInputBox).toHaveBeenCalled();
		treeDataProvider.render();
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(treeItem?.children![0].children![0].label).toEqual("Scope: Sample.One");
		expect(treeItem?.children![0].children![1].label).toEqual("Description: Sample description one");
		expect(treeItem?.children![0].children![2].label).toEqual("Consent: Admins and Users");
		expect(treeItem?.children![0].children![3].label).toEqual("Enabled: Yes");
	});

	test("Edit scope but cancel from user consent description input", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SCOPE-ENABLED", value: mockExposedApiId, state: true, label: "Sample Scope One" };
		jest.spyOn(vscode.window, "showInputBox")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce("New.Scope" as any)
			.mockResolvedValueOnce("New Scope" as any)
			.mockResolvedValueOnce("New Description" as any)
			.mockResolvedValueOnce("New User Display Name" as any);

		jest.spyOn(vscode.window, "showQuickPick")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce({ label: "Administrators only", value: "Admin" } as any);

		// Act
		await oauth2PermissionScopeService.edit(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(vscode.window.showInputBox).toHaveBeenCalled();
		treeDataProvider.render();
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(treeItem?.children![0].children![0].label).toEqual("Scope: Sample.One");
		expect(treeItem?.children![0].children![1].label).toEqual("Description: Sample description one");
		expect(treeItem?.children![0].children![2].label).toEqual("Consent: Admins and Users");
		expect(treeItem?.children![0].children![3].label).toEqual("Enabled: Yes");
	});

	test("Edit scope but cancel from state input", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SCOPE-ENABLED", value: mockExposedApiId, state: true, label: "Sample Scope One" };
		jest.spyOn(vscode.window, "showInputBox")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce("New.Scope" as any)
			.mockResolvedValueOnce("New Scope" as any)
			.mockResolvedValueOnce("New Description" as any)
			.mockResolvedValueOnce("New User Display Name" as any)
			.mockResolvedValueOnce("New User Description" as any);

		jest.spyOn(vscode.window, "showQuickPick")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce({ label: "Administrators only", value: "Admin" } as any);

		// Act
		await oauth2PermissionScopeService.edit(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(vscode.window.showInputBox).toHaveBeenCalled();
		treeDataProvider.render();
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(treeItem?.children![0].children![0].label).toEqual("Scope: Sample.One");
		expect(treeItem?.children![0].children![1].label).toEqual("Description: Sample description one");
		expect(treeItem?.children![0].children![2].label).toEqual("Consent: Admins and Users");
		expect(treeItem?.children![0].children![3].label).toEqual("Enabled: Yes");
	});

	test("Edit but no permissions returned", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SCOPE-ENABLED", value: mockExposedApiId, state: true, label: "Sample Scope One" };
		const error = new Error("Edit but no permissions returned");
		const graphSpy = jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementationOnce(async () => ({ success: false, error: error }));

		// Act
		await oauth2PermissionScopeService.edit(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(graphSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
	});

	test("Edit but no app id uri set", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SCOPE-ENABLED", value: mockExposedApiId, state: true, label: "Sample Scope One" };
		const warningSpy = jest.spyOn(vscode.window, "showWarningMessage").mockResolvedValue("OK" as any);
		mockApplications[0].identifierUris = [];

		// Act
		await oauth2PermissionScopeService.edit(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(warningSpy).toHaveBeenCalledWith("This application does not have an Application ID URI. Please add one before editing scopes.", "OK");
	});

	test("Add new scope successfully", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "EXPOSED-API-PERMISSIONS" };
		jest.spyOn(vscode.window, "showInputBox")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce("Add.Scope" as any)
			.mockResolvedValueOnce("Add Scope" as any)
			.mockResolvedValueOnce("New Add Description" as any)
			.mockResolvedValueOnce("New User Add Display Name" as any)
			.mockResolvedValueOnce("New User Add Description" as any);

		jest.spyOn(vscode.window, "showQuickPick")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce({ label: "Administrators only", value: "Admin" } as any)
			.mockResolvedValueOnce({ label: "Enabled", value: true } as any);

		// Act
		await oauth2PermissionScopeService.add(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(treeItem?.children!.length).toEqual(3);
		expect(treeItem?.children![2].label).toEqual("Add Scope");
		expect(treeItem?.children![2].children![0].label).toEqual("Scope: Add.Scope");
		expect(treeItem?.children![2].children![1].label).toEqual("Description: New Add Description");
		expect(treeItem?.children![2].children![2].label).toEqual("Consent: Admins Only");
		expect(treeItem?.children![2].children![3].label).toEqual("Enabled: Yes");
	});

	test("Add new scope but no permissions returned", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "EXPOSED-API-PERMISSIONS" };
		const error = new Error("Add new scope but no permissions returned");
		const graphSpy = jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementationOnce(async () => ({ success: false, error: error }));

		// Act
		await oauth2PermissionScopeService.add(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(graphSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
	});

	test("Add but no app id uri set", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SCOPE-ENABLED", value: mockExposedApiId, state: true, label: "Sample Scope One" };
		const warningSpy = jest.spyOn(vscode.window, "showWarningMessage").mockResolvedValue("OK" as any);
		mockApplications[0].identifierUris = [];

		// Act
		await oauth2PermissionScopeService.add(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(warningSpy).toHaveBeenCalledWith("This application does not have an Application ID URI. Please add one before adding a scope.", "OK");
	});

	test("Add new scope but escape pressed", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "EXPOSED-API-PERMISSIONS" };
		jest.spyOn(vscode.window, "showInputBox")
			.mockResolvedValue(undefined);

		// Act
		await oauth2PermissionScopeService.add(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(treeItem?.children!.length).toEqual(2);
	});
	
});
