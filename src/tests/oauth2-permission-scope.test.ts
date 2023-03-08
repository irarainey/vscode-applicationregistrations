import * as vscode from "vscode";
import * as validation from "../utils/validation";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { OAuth2PermissionScopeService } from "../services/oauth2-permission-scope";
import { mockApplications, mockAppObjectId, mockExposedApiId, seedMockData } from "./data/test-data";
import { getTopLevelTreeItem } from "./test-utils";
import { ApiApplication } from "@microsoft/microsoft-graph-types";

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
		triggerCompleteSpy = jest.spyOn(Object.getPrototypeOf(oauth2PermissionScopeService), "triggerRefresh");
		triggerErrorSpy = jest.spyOn(Object.getPrototypeOf(oauth2PermissionScopeService), "handleError");
		triggerTreeErrorSpy = jest.spyOn(Object.getPrototypeOf(treeDataProvider), "handleError");

		// The item to be tested
		item = { objectId: mockAppObjectId, contextValue: "EXPOSED-API-PERMISSIONS" };
	});

	afterAll(() => {
		oauth2PermissionScopeService.dispose();
		treeDataProvider.dispose();
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
		expect(treeItem?.children?.length).toEqual(2);
	});

	test("Delete disabled exposed api permissions but decline warning", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SCOPE-DISABLED", value: mockExposedApiId, state: false, label: "Sample Scope One" };
		const warningSpy = jest.spyOn(vscode.window, "showWarningMessage").mockResolvedValue("No" as any);

		// Act
		await oauth2PermissionScopeService.delete(item);

		// Assert
		await treeDataProvider.render();
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(warningSpy).toHaveBeenCalledWith(`Do you want to delete the Scope ${item.label}?`, "Yes", "No");
		expect(treeItem?.children?.length).toEqual(3);
	});

	test("Delete disabled exposed api permissions but error getting existing scopes", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SCOPE-DISABLED", value: mockExposedApiId, state: false };
		const error = new Error("Delete disabled exposed api permissions but error getting existing scopes");
		const graphSpy = jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementation(async (_id: string) => ({ success: false, error }));

		// Act
		await oauth2PermissionScopeService.delete(item);

		// Assert
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
		expect(treeItem?.children?.length).toEqual(2);
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
		await treeDataProvider.render();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(treeItem?.children?.[1].state).toEqual(false);
	});

	test("Enable disabled exposed api permission", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SCOPE-DISABLED", value: mockExposedApiId, state: false, label: "Sample Scope One" };

		// Act
		await oauth2PermissionScopeService.changeState(item, true);

		// Assert
		await treeDataProvider.render();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(treeItem?.children?.[1].state).toEqual(true);
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
		item = { objectId: mockAppObjectId, contextValue: "SCOPE-NAME", value: mockExposedApiId, state: true, label: "Sample Scope One" };
		jest.spyOn(vscode.window, "showInputBox").mockResolvedValue("New Scope" as any);

		// Act
		await oauth2PermissionScopeService.editValue(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(treeItem?.children?.[1].children?.[1].label).toEqual("Name: New Scope");
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
		expect(treeItem?.children![1].children![0].label).toEqual("Scope: New.Scope");
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
		expect(treeItem?.children![1].children![2].label).toEqual("Description: New Description");
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
		expect(treeItem?.children![1].children![3].label).toEqual("Consent: Admins Only");
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
		expect(treeItem?.children![1].children![0].label).toEqual("Scope: New.Scope");
		expect(treeItem?.children![1].children![1].label).toEqual("Name: New Scope");
		expect(treeItem?.children![1].children![2].label).toEqual("Description: New Description");
		expect(treeItem?.children![1].children![3].label).toEqual("Consent: Admins Only");
		expect(treeItem?.children![1].children![4].label).toEqual("Enabled: Yes");
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
		await treeDataProvider.render();
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(treeItem?.children![1].children![0].label).toEqual("Scope: Sample.One");
		expect(treeItem?.children![1].children![1].label).toEqual("Name: Sample Scope One");
		expect(treeItem?.children![1].children![2].label).toEqual("Description: Sample description one");
		expect(treeItem?.children![1].children![3].label).toEqual("Consent: Admins and Users");
		expect(treeItem?.children![1].children![4].label).toEqual("Enabled: Yes");
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
		await treeDataProvider.render();
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(treeItem?.children![1].children![0].label).toEqual("Scope: Sample.One");
		expect(treeItem?.children![1].children![1].label).toEqual("Name: Sample Scope One");
		expect(treeItem?.children![1].children![2].label).toEqual("Description: Sample description one");
		expect(treeItem?.children![1].children![3].label).toEqual("Consent: Admins and Users");
		expect(treeItem?.children![1].children![4].label).toEqual("Enabled: Yes");
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
		await treeDataProvider.render();
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(treeItem?.children![1].children![0].label).toEqual("Scope: Sample.One");
		expect(treeItem?.children![1].children![1].label).toEqual("Name: Sample Scope One");
		expect(treeItem?.children![1].children![2].label).toEqual("Description: Sample description one");
		expect(treeItem?.children![1].children![3].label).toEqual("Consent: Admins and Users");
		expect(treeItem?.children![1].children![4].label).toEqual("Enabled: Yes");
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
		await treeDataProvider.render();
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(treeItem?.children![1].children![0].label).toEqual("Scope: Sample.One");
		expect(treeItem?.children![1].children![1].label).toEqual("Name: Sample Scope One");
		expect(treeItem?.children![1].children![2].label).toEqual("Description: Sample description one");
		expect(treeItem?.children![1].children![3].label).toEqual("Consent: Admins and Users");
		expect(treeItem?.children![1].children![4].label).toEqual("Enabled: Yes");
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
		await treeDataProvider.render();
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(treeItem?.children![1].children![0].label).toEqual("Scope: Sample.One");
		expect(treeItem?.children![1].children![1].label).toEqual("Name: Sample Scope One");
		expect(treeItem?.children![1].children![2].label).toEqual("Description: Sample description one");
		expect(treeItem?.children![1].children![3].label).toEqual("Consent: Admins and Users");
		expect(treeItem?.children![1].children![4].label).toEqual("Enabled: Yes");
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
		await treeDataProvider.render();
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(treeItem?.children![1].children![0].label).toEqual("Scope: Sample.One");
		expect(treeItem?.children![1].children![1].label).toEqual("Name: Sample Scope One");
		expect(treeItem?.children![1].children![2].label).toEqual("Description: Sample description one");
		expect(treeItem?.children![1].children![3].label).toEqual("Consent: Admins and Users");
		expect(treeItem?.children![1].children![4].label).toEqual("Enabled: Yes");
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
		await treeDataProvider.render();
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(treeItem?.children![1].children![0].label).toEqual("Scope: Sample.One");
		expect(treeItem?.children![1].children![1].label).toEqual("Name: Sample Scope One");
		expect(treeItem?.children![1].children![2].label).toEqual("Description: Sample description one");
		expect(treeItem?.children![1].children![3].label).toEqual("Consent: Admins and Users");
		expect(treeItem?.children![1].children![4].label).toEqual("Enabled: Yes");
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
		const newScopeValue = "Add.Scope";
		const newScopeAdminDisplayName = "Add Scope";
		const newScopeAdminDescription = "Add Description";
		const newScopeUserDisplayName = "User Add Display Name";
		const newScopeUserDescription = "User Add Description";
		item = { objectId: mockAppObjectId, contextValue: "EXPOSED-API-PERMISSIONS" };
		jest.spyOn(vscode.window, "showInputBox")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce(newScopeUserDescription as any);

		jest.spyOn(vscode.window, "showQuickPick")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce({ label: "Administrators only", value: "Admin" } as any)
			.mockResolvedValueOnce({ label: "Enabled", value: true } as any);

		jest.spyOn(oauth2PermissionScopeService, "inputValue").mockImplementation(async (_title: string, _existingValue: string | undefined, _isEditing: boolean, _signInAudience: string, _scopes: ApiApplication, validate: (value: string, isEditing: boolean, existingValue: string | undefined, signInAudience: string, scopes: ApiApplication) => string | undefined) => {
			const result = validate(newScopeValue, false, undefined, "AzureADMyOrg", mockApplications[0].api);
			expect(result).toBeUndefined();
			return newScopeValue;
		});
	
		jest.spyOn(oauth2PermissionScopeService, "inputAdminConsentDisplayName").mockImplementation(async (_title: string, _existingValue: string | undefined, validate: (displayName: string) => string | undefined) => {
			const result = validate(newScopeAdminDisplayName);
			expect(result).toBeUndefined();
			return newScopeAdminDisplayName;
		});

		jest.spyOn(oauth2PermissionScopeService, "inputAdminConsentDescription").mockImplementation(async (_title: string, _existingValue: string | undefined, validate: (description: string) => string | undefined) => {
			const result = validate(newScopeAdminDescription);
			expect(result).toBeUndefined();
			return newScopeAdminDescription;
		});

		jest.spyOn(oauth2PermissionScopeService, "inputUserConsentDisplayName").mockImplementation(async (_title: string, _existingValue: string | undefined, validate: (displayName: string) => string | undefined) => {
			const result = validate(newScopeUserDisplayName);
			expect(result).toBeUndefined();
			return newScopeUserDisplayName;
		});

		// Act
		await oauth2PermissionScopeService.add(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(treeItem?.children!.length).toEqual(4);
		expect(treeItem?.children![3].label).toEqual(newScopeValue);
		expect(treeItem?.children![3].children![0].label).toEqual(`Scope: ${newScopeValue}`);
		expect(treeItem?.children![3].children![1].label).toEqual(`Name: ${newScopeAdminDisplayName}`);
		expect(treeItem?.children![3].children![2].label).toEqual(`Description: ${newScopeAdminDescription}`);
		expect(treeItem?.children![3].children![3].label).toEqual("Consent: Admins Only");
		expect(treeItem?.children![3].children![4].label).toEqual("Enabled: Yes");
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
		expect(treeItem?.children!.length).toEqual(3);
	});

	test("Add new scope with admin display name too long error", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "EXPOSED-API-PERMISSIONS" };
		const inputSpy = jest.spyOn(vscode.window, "showInputBox")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce("Add.Scope" as any);
		jest.spyOn(vscode.window, "showQuickPick")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce({ label: "Administrators only", value: "Admin" } as any);
	
		const validationSpy = jest.spyOn(validation, "validateScopeAdminDisplayName");
		jest.spyOn(oauth2PermissionScopeService, "inputAdminConsentDisplayName").mockImplementation(async (_title: string, _existingValue: string | undefined, validate: (displayName: string) => string | undefined) => {
			const result = validate("X".padEnd(101, "X"));
			expect(result).toBe("An admin consent display name cannot be longer than 100 characters.");
			return undefined;
		});

		// Act
		await oauth2PermissionScopeService.add(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(inputSpy).toBeCalled();
		expect(validationSpy).toBeCalled();
	});

	test("Add new scope with user display name too long error", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "EXPOSED-API-PERMISSIONS" };
		const inputSpy = jest.spyOn(vscode.window, "showInputBox")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce("New.Scope" as any)
			.mockResolvedValueOnce("New Scope" as any)
			.mockResolvedValueOnce("New Description" as any);
		jest.spyOn(vscode.window, "showQuickPick")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce({ label: "Administrators only", value: "Admin" } as any);
	
		const validationSpy = jest.spyOn(validation, "validateScopeUserDisplayName");
		jest.spyOn(oauth2PermissionScopeService, "inputUserConsentDisplayName").mockImplementation(async (_title: string, _existingValue: string | undefined, validate: (displayName: string) => string | undefined) => {
			const result = validate("X".padEnd(101, "X"));
			expect(result).toBe("An user consent display name cannot be longer than 100 characters.");
			return undefined;
		});

		// Act
		await oauth2PermissionScopeService.add(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(inputSpy).toBeCalled();
		expect(validationSpy).toBeCalled();
	});

	test("Add new scope with admin display name too short error", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "EXPOSED-API-PERMISSIONS" };
		const inputSpy = jest.spyOn(vscode.window, "showInputBox")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce("Add.Scope" as any);
		jest.spyOn(vscode.window, "showQuickPick")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce({ label: "Administrators only", value: "Admin" } as any);
	
		const validationSpy = jest.spyOn(validation, "validateScopeAdminDisplayName");
		jest.spyOn(oauth2PermissionScopeService, "inputAdminConsentDisplayName").mockImplementation(async (_title: string, _existingValue: string | undefined, validate: (displayName: string) => string | undefined) => {
			const result = validate("");
			expect(result).toBe("An admin consent display name cannot be empty.");
			return undefined;
		});

		// Act
		await oauth2PermissionScopeService.add(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(inputSpy).toBeCalled();
		expect(validationSpy).toBeCalled();
	});

	test("Add new scope with admin description too short error", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "EXPOSED-API-PERMISSIONS" };
		const inputSpy = jest.spyOn(vscode.window, "showInputBox")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce("Add.Scope" as any)
			.mockResolvedValueOnce("Add Scope" as any);
		jest.spyOn(vscode.window, "showQuickPick")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce({ label: "Administrators only", value: "Admin" } as any);
	
		const validationSpy = jest.spyOn(validation, "validateScopeAdminDescription");
		jest.spyOn(oauth2PermissionScopeService, "inputAdminConsentDescription").mockImplementation(async (_title: string, _existingValue: string | undefined, validate: (displayName: string) => string | undefined) => {
			const result = validate("");
			expect(result).toBe("An admin consent description cannot be empty.");
			return undefined;
		});

		// Act
		await oauth2PermissionScopeService.add(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(inputSpy).toBeCalled();
		expect(validationSpy).toBeCalled();
	});

	test("Add new scope with scope value too long for AAD audiences error", async () => {
		// Arrange
		const errorMessage = "A value cannot be longer than 120 characters.";
		item = { objectId: mockAppObjectId, contextValue: "EXPOSED-API-PERMISSIONS" };
		const validationSpy = jest.spyOn(validation, "validateScopeValue");
		jest.spyOn(oauth2PermissionScopeService, "inputValue").mockImplementation(async (_title: string, _existingValue: string | undefined, _isEditing: boolean, _signInAudience: string, _scopes: ApiApplication, validate: (value: string, isEditing: boolean, existingValue: string | undefined, signInAudience: string, scopes: ApiApplication) => string | undefined) => {
			const result = validate("X.X".padEnd(121, "X"), false, undefined, "AzureADMyOrg", { oauth2PermissionScopes: [] });
			expect(result).toBe(errorMessage);
			return errorMessage;
		});

		// Act
		await oauth2PermissionScopeService.add(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(validationSpy).toBeCalled();
	});

	test("Add new scope with scope value too long for consumer audiences error", async () => {
		// Arrange
		const errorMessage = "A value cannot be longer than 40 characters.";
		item = { objectId: mockAppObjectId, contextValue: "EXPOSED-API-PERMISSIONS" };
		const validationSpy = jest.spyOn(validation, "validateScopeValue");
		jest.spyOn(oauth2PermissionScopeService, "inputValue").mockImplementation(async (_title: string, _existingValue: string | undefined, _isEditing: boolean, _signInAudience: string, _scopes: ApiApplication, validate: (value: string, isEditing: boolean, existingValue: string | undefined, signInAudience: string, scopes: ApiApplication) => string | undefined) => {
			const result = validate("X.X".padEnd(41, "X"), false, undefined, "AzureADandPersonalMicrosoftAccount", { oauth2PermissionScopes: [] });
			expect(result).toBe(errorMessage);
			return errorMessage;
		});

		// Act
		await oauth2PermissionScopeService.add(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(validationSpy).toBeCalled();
	});

	test("Add new scope with scope value too short error", async () => {
		// Arrange
		const errorMessage = "A scope value cannot be empty.";
		item = { objectId: mockAppObjectId, contextValue: "EXPOSED-API-PERMISSIONS" };
		const validationSpy = jest.spyOn(validation, "validateScopeValue");
		jest.spyOn(oauth2PermissionScopeService, "inputValue").mockImplementation(async (_title: string, _existingValue: string | undefined, _isEditing: boolean, _signInAudience: string, _scopes: ApiApplication, validate: (value: string, isEditing: boolean, existingValue: string | undefined, signInAudience: string, scopes: ApiApplication) => string | undefined) => {
			const result = validate("", false, undefined, "AzureADMyOrg", { oauth2PermissionScopes: [] });
			expect(result).toBe(errorMessage);
			return errorMessage;
		});

		// Act
		await oauth2PermissionScopeService.add(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(validationSpy).toBeCalled();
	});

	test("Add new scope with scope value with space error", async () => {
		// Arrange
		const errorMessage = "A scope value cannot contain spaces.";
		item = { objectId: mockAppObjectId, contextValue: "EXPOSED-API-PERMISSIONS" };
		const validationSpy = jest.spyOn(validation, "validateScopeValue");
		jest.spyOn(oauth2PermissionScopeService, "inputValue").mockImplementation(async (_title: string, _existingValue: string | undefined, _isEditing: boolean, _signInAudience: string, _scopes: ApiApplication, validate: (value: string, isEditing: boolean, existingValue: string | undefined, signInAudience: string, scopes: ApiApplication) => string | undefined) => {
			const result = validate("TEST SCOPE", false, undefined, "AzureADMyOrg", { oauth2PermissionScopes: [] });
			expect(result).toBe(errorMessage);
			return errorMessage;
		});

		// Act
		await oauth2PermissionScopeService.add(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(validationSpy).toBeCalled();
	});

	test("Add new scope with scope value start with full stop error", async () => {
		// Arrange
		const errorMessage = "A scope value cannot start with a full stop.";
		item = { objectId: mockAppObjectId, contextValue: "EXPOSED-API-PERMISSIONS" };
		const validationSpy = jest.spyOn(validation, "validateScopeValue");
		jest.spyOn(oauth2PermissionScopeService, "inputValue").mockImplementation(async (_title: string, _existingValue: string | undefined, _isEditing: boolean, _signInAudience: string, _scopes: ApiApplication, validate: (value: string, isEditing: boolean, existingValue: string | undefined, signInAudience: string, scopes: ApiApplication) => string | undefined) => {
			const result = validate(".TEST.SCOPE", false, undefined, "AzureADMyOrg", { oauth2PermissionScopes: [] });
			expect(result).toBe(errorMessage);
			return errorMessage;
		});

		// Act
		await oauth2PermissionScopeService.add(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(validationSpy).toBeCalled();
	});

	test("Add new scope with scope exist value error", async () => {
		// Arrange
		const errorMessage = "The scope value specified already exists.";
		item = { objectId: mockAppObjectId, contextValue: "EXPOSED-API-PERMISSIONS" };
		const validationSpy = jest.spyOn(validation, "validateScopeValue");
		jest.spyOn(oauth2PermissionScopeService, "inputValue").mockImplementation(async (_title: string, _existingValue: string | undefined, _isEditing: boolean, _signInAudience: string, _scopes: ApiApplication, validate: (value: string, isEditing: boolean, existingValue: string | undefined, signInAudience: string, scopes: ApiApplication) => string | undefined) => {
			const result = validate("Sample.One", true, undefined, "AzureADMyOrg", mockApplications[0].api);
			expect(result).toBe(errorMessage);
			return errorMessage;
		});

		// Act
		await oauth2PermissionScopeService.add(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(validationSpy).toBeCalled();
	});
	
	test("Error getting exposed api permission children", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "EXPOSED-API-PERMISSIONS" };
		const error = new Error("Error getting exposed api permission children");
		jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementation(async (id: string, select: string) => {
			if (select === "api") {
				return { success: false, error };
			}
			return mockApplications.find((app) => app.id === id);
		});

		// Act
		await treeDataProvider.render();
		await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");

		// Assert
		expect(triggerTreeErrorSpy).toHaveBeenCalledWith(error);
	});
});
