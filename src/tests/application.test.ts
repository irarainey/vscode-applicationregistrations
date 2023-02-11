import * as vscode from "vscode";
import * as execShellCmdUtil from "../utils/exec-shell-cmd";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { ApplicationService } from "../services/application";
import { Application } from "@microsoft/microsoft-graph-types";
import { mockAppId, mockAppObjectId, mockTenantId, seedMockData } from "../repositories/__mocks__/test-data";
import { getTopLevelTreeItem } from "./utils";

// Create Jest mocks
jest.mock("vscode");
jest.mock("../repositories/graph-api-repository");
jest.mock("../utils/exec-shell-cmd");

// Create the test suite for application service
describe("Application Service Tests", () => {
	// Create instances of objects used in the tests
	const graphApiRepository = new GraphApiRepository();
	const treeDataProvider = new AppRegTreeDataProvider(graphApiRepository);
	const applicationService = new ApplicationService(graphApiRepository, treeDataProvider);

	// Create spy variables
	let triggerCompleteSpy: jest.SpyInstance<any, unknown[], any>;
	let triggerErrorSpy: jest.SpyInstance<any, unknown[], any>;
	let statusBarSpy: jest.SpyInstance<any, [text: string], any>;
	let openTextDocumentSpy: jest.SpyInstance<any, any, any>;
	let openExternalSpy: jest.SpyInstance<Thenable<boolean>, [target: vscode.Uri], any>;

	// The item to be tested
	let item: AppRegItem = { objectId: mockAppObjectId, appId: mockAppId, contextValue: "APPLICATION" };

	beforeAll(() => {
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

		// Define a mock implementation for the execShellCmd function
		jest.spyOn(execShellCmdUtil, "execShellCmd").mockImplementation(async (_cmd: string) => mockTenantId);

		// Define spies on the functions to be tested
		statusBarSpy = jest.spyOn(vscode.window, "setStatusBarMessage");
		openTextDocumentSpy = jest.spyOn(vscode.workspace, "openTextDocument");
		openExternalSpy = jest.spyOn(vscode.env, "openExternal");
		triggerCompleteSpy = jest.spyOn(Object.getPrototypeOf(applicationService), "triggerRefresh");
		triggerErrorSpy = jest.spyOn(Object.getPrototypeOf(applicationService), "handleError");
	});

	afterAll(() => {
		// Dispose of the application service
		applicationService.dispose();
	});

	test("Create class instance", () => {
		// Assert class has been instantiated
		expect(applicationService).toBeDefined();
	});

	test("Show endpoints for multi-tenant app", async () => {
		// Act
		await applicationService.showEndpoints(item);

		// Assert
		expect(openTextDocumentSpy).toHaveBeenCalled();
	});

	test("Show endpoints for multi-tenant and personal app", async () => {
		// Arrange
		jest.spyOn(graphApiRepository, "getSignInAudience").mockImplementation(async (_id: string) => ({ success: true, value: "AzureADandPersonalMicrosoftAccount" }));

		// Act
		await applicationService.showEndpoints(item);

		// Assert
		expect(openTextDocumentSpy).toHaveBeenCalled();
	});

	test("Show endpoints for consumer app", async () => {
		// Arrange
		jest.spyOn(graphApiRepository, "getSignInAudience").mockImplementation(async (_id: string) => ({ success: true, value: "PersonalMicrosoftAccount" }));

		// Act
		await applicationService.showEndpoints(item);

		// Assert
		expect(openTextDocumentSpy).toHaveBeenCalled();
	});

	test("Show endpoints for consumer app", async () => {
		// Arrange
		jest.spyOn(graphApiRepository, "getSignInAudience").mockImplementation(async (_id: string) => ({ success: true, value: "AzureADMyOrg" }));

		// Act
		await applicationService.showEndpoints(item);

		// Assert
		expect(openTextDocumentSpy).toHaveBeenCalled();
	});

	test("Show endpoints for unknown app type", async () => {
		// Arrange
		jest.spyOn(graphApiRepository, "getSignInAudience").mockImplementation(async (_id: string) => ({ success: true, value: "Unknown" }));

		// Act
		await applicationService.showEndpoints(item);

		// Assert
		expect(openTextDocumentSpy).toHaveBeenCalled();
	});

	test("Show endpoints with sign in audience error", async () => {
		// Arrange
		jest.spyOn(graphApiRepository, "getSignInAudience").mockImplementation(async (_id: string) => ({ success: false, error: new Error("Test Error") }));

		// Act
		await applicationService.showEndpoints(item);

		// Assert
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	test("Show endpoints with Az CLI error", async () => {
		// Arrange
		jest.spyOn(graphApiRepository, "getSignInAudience").mockImplementation(async (_id: string) => ({ success: true, value: "AzureADMyOrg" }));
		jest.spyOn(execShellCmdUtil, "execShellCmd").mockImplementation(async (_cmd: string) => { throw new Error("Test Error"); });

		// Act
		await applicationService.showEndpoints(item);

		// Assert
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	test("View manifest", async () => {
		// Act
		await applicationService.viewManifest(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(openTextDocumentSpy).toHaveBeenCalled();
	});

	test("View manifest error", async () => {
		// Arrange
		jest.spyOn(graphApiRepository, "getApplicationDetailsFull").mockImplementation(async (_id: string) => ({ success: false, error: new Error("Test Error") }));

		// Act
		await applicationService.viewManifest(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	test("Copy client id", async () => {
		// Act
		applicationService.copyClientId(item);

		// Assert
		expect(vscode.env.clipboard.readText()).toEqual(mockAppId);
	});

	test("Open in portal", async () => {
		// Act
		applicationService.openInPortal(item);

		// Assert
		expect(openExternalSpy).toHaveBeenCalled();
	});

	test("Edit Logout URL with unset value", async () => {
		// Arrange
		item = { ...item, value: "Not set", contextValue: "LOGOUT-URL" };
		vscode.window.showInputBox = jest.fn().mockResolvedValue("https://test.com/logout");

		// Act
		await applicationService.editLogoutUrl(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "LOGOUT-URL-PARENT");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem!.children![0].label).toEqual("https://test.com/logout");
	});

	test("Edit Logout URL with existing value", async () => {
		// Arrange
		item = { ...item, value: "https://oldtest.com/logout", contextValue: "LOGOUT-URL" };
		vscode.window.showInputBox = jest.fn().mockResolvedValue("https://test.com/logout");

		// Act
		await applicationService.editLogoutUrl(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "LOGOUT-URL-PARENT");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem!.children![0].label).toEqual("https://test.com/logout");
	});

	test("Edit App Id Uri with unset value", async () => {
		// Arrange
		item = { ...item, value: "Not set", contextValue: "APPID-URI" };
		vscode.window.showInputBox = jest.fn().mockResolvedValue("api://test.com");

		// Act
		await applicationService.editAppIdUri(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "APPID-URI-PARENT");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem!.children![0].label).toEqual("api://test.com");
	});

	test("Edit App Id Uri with existing value", async () => {
		// Arrange
		item = { ...item, value: "api://oldtest.com", contextValue: "APPID-URI" };
		vscode.window.showInputBox = jest.fn().mockResolvedValue("api://test.com");

		// Act
		await applicationService.editAppIdUri(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "APPID-URI-PARENT");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem!.children![0].label).toEqual("api://test.com");
	});

	test("Edit App Id Uri sign in error", async () => {
		// Arrange
		item = { ...item, value: "api://oldtest.com", contextValue: "APPID-URI" };
		vscode.window.showInputBox = jest.fn().mockResolvedValue("api://test.com");
		jest.spyOn(graphApiRepository, "getSignInAudience").mockImplementation(async (_id: string) => ({ success: false, error: new Error("Test Error") }));

		// Act
		await applicationService.editAppIdUri(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	test("Remove Logout URL", async () => {
		// Act
		await applicationService.removeLogoutUrl(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "LOGOUT-URL-PARENT");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem!.children![0].label).toEqual("Not set");
	});

	test("Remove Logout URL update application error", async () => {
		// Arrange
		jest.spyOn(graphApiRepository, "updateApplication").mockImplementation(async (_id: string, _appChange: Application) => ({ success: false, error: new Error("Test Error") }));

		// Act
		await applicationService.removeLogoutUrl(item);

		// Assert
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	test("Remove App Id Uri", async () => {
		// Act
		await applicationService.removeAppIdUri(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "APPID-URI-PARENT");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem!.children![0].label).toEqual("Not set");
	});

	test("Delete application successfully", async () => {
		// Act
		await applicationService.delete(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider);
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem).toBeUndefined();
	});

	test("Delete application with error", async () => {
		// Arrange
		jest.spyOn(graphApiRepository, "deleteApplication").mockImplementation(async (_id: string) => ({ success: false, error: new Error("Test Error") }));

		// Act
		await applicationService.delete(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	test("Rename application successfully", async () => {
		// Arrange
		vscode.window.showInputBox = jest.fn().mockResolvedValue("New Application Name");

		// Act
		await applicationService.rename(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider);
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.label).toEqual("New Application Name");
	});

	test("Rename application with update error", async () => {
		// Arrange
		jest.spyOn(graphApiRepository, "updateApplication").mockImplementation(async (_id: string, _appChange: Application) => ({ success: false, error: new Error("Test Error") }));
		vscode.window.showInputBox = jest.fn().mockResolvedValue("New Application Name");

		// Act
		await applicationService.rename(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	test("Rename application with sign in error", async () => {
		// Arrange
		jest.spyOn(graphApiRepository, "getSignInAudience").mockImplementation(async (_id: string) => ({ success: false, error: new Error("Test Error") }));

		// Act
		await applicationService.rename(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	test("Add application successfully", async () => {
		// Arrange
		vscode.window.showQuickPick = jest.fn().mockResolvedValue({ value: "AzureADMyOrg" });
		vscode.window.showInputBox = jest.fn().mockResolvedValue("Add Application Name");

		// Act
		await applicationService.add();

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
	});

	test("Add application with creation error", async () => {
		// Arrange
		jest.spyOn(graphApiRepository, "createApplication").mockImplementation(async () => ({ success: false, error: new Error("Test Error") }));
		vscode.window.showQuickPick = jest.fn().mockResolvedValue({ value: "AzureADMyOrg" });
		vscode.window.showInputBox = jest.fn().mockResolvedValue("Add Application Name");

		// Act
		await applicationService.add();

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalled();
	});
});
