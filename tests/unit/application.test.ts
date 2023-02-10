import * as vscode from "vscode";
import * as execShellCmdUtil from "../../src/utils/exec-shell-cmd";
import { GraphApiRepository } from "../../src/repositories/graph-api-repository";
import { AppRegTreeDataProvider } from "../../src/data/app-reg-tree-data-provider";
import { AppRegItem } from "../../src/models/app-reg-item";
import { ApplicationService } from "../../src/services/application";
import { Application } from "@microsoft/microsoft-graph-types";
import { mockAppId, mockAppObjectId, mockTenantId } from "./constants";
import { validateApplicationDisplayName } from "../../src/utils/validation";

// Create Jest mocks
jest.mock("vscode");
jest.mock("../../src/repositories/graph-api-repository");
jest.mock("../../src/utils/exec-shell-cmd");

// Create the test suite for sign in audience service
describe("Application Service Tests", () => {
	// Create instances of objects used in the tests
	const graphApiRepository = new GraphApiRepository();
	const treeDataProvider = new AppRegTreeDataProvider(graphApiRepository);
	const applicationService = new ApplicationService(graphApiRepository, treeDataProvider);

	// Create spy variables
	let triggerCompleteSpy: jest.SpyInstance<any, unknown[], any>;
	let triggerErrorSpy: jest.SpyInstance<any, unknown[], any>;
	let statusBarSpy: jest.SpyInstance<any, [text: string], any>;
	let iconSpy: jest.SpyInstance<any, [id: string, color?: any | undefined], any>;
	let openTextDocumentSpy: jest.SpyInstance<any, any, any>;
	let openExternalSpy: jest.SpyInstance<Thenable<boolean>, [target: vscode.Uri], any>;

	// The item to be tested
	let item: AppRegItem = { objectId: mockAppObjectId, appId: mockAppId, contextValue: "APPLICATION" };

	// Create common mock functions for all tests
	beforeAll(() => {
		console.error = jest.fn();
	});

	// Create a generic item to use in each test
	beforeEach(() => {
		jest.restoreAllMocks();
		jest.spyOn(execShellCmdUtil, "execShellCmd").mockImplementation(async (_cmd: string) => mockTenantId);
		vscode.window.showWarningMessage = jest.fn().mockResolvedValue("Yes");
		statusBarSpy = jest.spyOn(vscode.window, "setStatusBarMessage");
		openTextDocumentSpy = jest.spyOn(vscode.workspace, "openTextDocument");
		openExternalSpy = jest.spyOn(vscode.env, "openExternal");
		iconSpy = jest.spyOn(vscode, "ThemeIcon");
		triggerCompleteSpy = jest.spyOn(Object.getPrototypeOf(applicationService), "triggerRefresh");
		triggerErrorSpy = jest.spyOn(Object.getPrototypeOf(applicationService), "handleError");
	});

	afterAll(() => {
		applicationService.dispose();
	});

	test("Create class instance", () => {
		expect(applicationService).toBeDefined();
	});

	test("Show endpoints for multi-tenant app", async () => {
		await applicationService.showEndpoints(item);
		expect(openTextDocumentSpy).toHaveBeenCalled();
	});

	test("Show endpoints for multi-tenant and personal app", async () => {
		jest.spyOn(graphApiRepository, "getSignInAudience").mockImplementation(async (_id: string) => ({ success: true, value: "AzureADandPersonalMicrosoftAccount" }));
		await applicationService.showEndpoints(item);
		expect(openTextDocumentSpy).toHaveBeenCalled();
	});

	test("Show endpoints for consumer app", async () => {
		jest.spyOn(graphApiRepository, "getSignInAudience").mockImplementation(async (_id: string) => ({ success: true, value: "PersonalMicrosoftAccount" }));
		await applicationService.showEndpoints(item);
		expect(openTextDocumentSpy).toHaveBeenCalled();
	});

	test("Show endpoints for consumer app", async () => {
		jest.spyOn(graphApiRepository, "getSignInAudience").mockImplementation(async (_id: string) => ({ success: true, value: "AzureADMyOrg" }));
		await applicationService.showEndpoints(item);
		expect(openTextDocumentSpy).toHaveBeenCalled();
	});

	test("Show endpoints for unknown app type", async () => {
		jest.spyOn(graphApiRepository, "getSignInAudience").mockImplementation(async (_id: string) => ({ success: true, value: "Unknown" }));
		await applicationService.showEndpoints(item);
		expect(openTextDocumentSpy).toHaveBeenCalled();
	});

	test("Show endpoints with sign in audience error", async () => {
		jest.spyOn(graphApiRepository, "getSignInAudience").mockImplementation(async (_id: string) => ({ success: false, error: new Error("Test Error") }));
		await applicationService.showEndpoints(item);
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	test("Show endpoints with Az CLI error", async () => {
		jest.spyOn(graphApiRepository, "getSignInAudience").mockImplementation(async (_id: string) => ({ success: true, value: "AzureADMyOrg" }));
		jest.spyOn(execShellCmdUtil, "execShellCmd").mockImplementation(async (_cmd: string) => {
			throw new Error("Test Error");
		});
		await applicationService.showEndpoints(item);
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	test("View manifest", async () => {
		await applicationService.viewManifest(item);
		expect(statusBarSpy).toHaveBeenCalled();
		expect(openTextDocumentSpy).toHaveBeenCalled();
	});

	test("View manifest error", async () => {
		jest.spyOn(graphApiRepository, "getApplicationDetailsFull").mockImplementation(async (_id: string) => ({ success: false, error: new Error("Test Error") }));
		await applicationService.viewManifest(item);
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	test("Copy client id", async () => {
		applicationService.copyClientId(item);
		expect(vscode.env.clipboard.readText()).toEqual(mockAppId);
	});

	test("Open in portal", async () => {
		applicationService.openInPortal(item);
		expect(openExternalSpy).toHaveBeenCalled();
	});

	test("Edit Logout URL with unset value", async () => {
		item = { ...item, value: "Not set", contextValue: "LOGOUT-URL" };
		vscode.window.showInputBox = jest.fn().mockResolvedValue("https://test.com/logout");
		await applicationService.editLogoutUrl(item);
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, "LOGOUT-URL-PARENT");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem!.children![0].label).toEqual("https://test.com/logout");
	});

	test("Edit Logout URL with existing value", async () => {
		item = { ...item, value: "https://oldtest.com/logout", contextValue: "LOGOUT-URL" };
		vscode.window.showInputBox = jest.fn().mockResolvedValue("https://test.com/logout");
		await applicationService.editLogoutUrl(item);
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, "LOGOUT-URL-PARENT");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem!.children![0].label).toEqual("https://test.com/logout");
	});

	test("Edit App Id Uri with unset value", async () => {
		item = { ...item, value: "Not set", contextValue: "APPID-URI" };
		vscode.window.showInputBox = jest.fn().mockResolvedValue("api://test.com");
		await applicationService.editAppIdUri(item);
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, "APPID-URI-PARENT");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem!.children![0].label).toEqual("api://test.com");
	});

	test("Edit App Id Uri with existing value", async () => {
		item = { ...item, value: "api://oldtest.com", contextValue: "APPID-URI" };
		vscode.window.showInputBox = jest.fn().mockResolvedValue("api://test.com");
		await applicationService.editAppIdUri(item);
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, "APPID-URI-PARENT");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem!.children![0].label).toEqual("api://test.com");
	});

	test("Edit App Id Uri sign in error", async () => {
		item = { ...item, value: "api://oldtest.com", contextValue: "APPID-URI" };
		vscode.window.showInputBox = jest.fn().mockResolvedValue("api://test.com");
		jest.spyOn(graphApiRepository, "getSignInAudience").mockImplementation(async (_id: string) => ({ success: false, error: new Error("Test Error") }));
		await applicationService.editAppIdUri(item);
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, "APPID-URI-PARENT");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	test("Remove Logout URL", async () => {
		await applicationService.removeLogoutUrl(item);
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, "LOGOUT-URL-PARENT");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem!.children![0].label).toEqual("Not set");
	});

	test("Remove Logout URL update application error", async () => {
		jest.spyOn(graphApiRepository, "updateApplication").mockImplementation(async (_id: string, _appChange: Application) => ({ success: false, error: new Error("Test Error") }));
		await applicationService.removeLogoutUrl(item);
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	test("Remove App Id Uri", async () => {
		await applicationService.removeAppIdUri(item);
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, "APPID-URI-PARENT");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem!.children![0].label).toEqual("Not set");
	});

	test("Delete application successfully", async () => {
		await applicationService.delete(item);
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
	});

	test("Delete application with error", async () => {
		jest.spyOn(graphApiRepository, "deleteApplication").mockImplementation(async (_id: string) => ({ success: false, error: new Error("Test Error") }));
		await applicationService.delete(item);
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	test("Rename application successfully", async () => {
		vscode.window.showInputBox = jest.fn().mockResolvedValue("New Application Name");
		await applicationService.rename(item);
		const treeItem = await getTopLevelTreeItem(mockAppObjectId);
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.label).toEqual("New Application Name");
	});

	test("Rename application with update error", async () => {
		jest.spyOn(graphApiRepository, "updateApplication").mockImplementation(async (_id: string, _appChange: Application) => ({ success: false, error: new Error("Test Error") }));
		vscode.window.showInputBox = jest.fn().mockResolvedValue("New Application Name");
		await applicationService.rename(item);
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	test("Rename application with sign in error", async () => {
		jest.spyOn(graphApiRepository, "getSignInAudience").mockImplementation(async (_id: string) => ({ success: false, error: new Error("Test Error") }));
		await applicationService.rename(item);
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	test("Add application successfully", async () => {
		vscode.window.showQuickPick = jest.fn().mockResolvedValue({ value: "AzureADMyOrg" });
		vscode.window.showInputBox = jest.fn().mockResolvedValue("Add Application Name");
		await applicationService.add();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
	});

	test("Add application with creation error", async () => {
		jest.spyOn(graphApiRepository, "createApplication").mockImplementation(async () => ({ success: false, error: new Error("Test Error") }));
		vscode.window.showQuickPick = jest.fn().mockResolvedValue({ value: "AzureADMyOrg" });
		vscode.window.showInputBox = jest.fn().mockResolvedValue("Add Application Name");
		await applicationService.add();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	// Get a specific top level tree item
	const getTopLevelTreeItem = async (objectId: string, contextValue?: string): Promise<AppRegItem | undefined> => {
		const tree = await treeDataProvider.getChildren();
		const app = tree!.find((x) => x.objectId === objectId);
		return contextValue === undefined ? app : app?.children?.find((x) => x.contextValue === contextValue);
	};
});
