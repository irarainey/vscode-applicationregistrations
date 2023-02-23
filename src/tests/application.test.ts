import * as vscode from "vscode";
import * as execShellCmdUtil from "../utils/exec-shell-cmd";
import * as validation from "../utils/validation";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { ApplicationService } from "../services/application";
import { Application } from "@microsoft/microsoft-graph-types";
import { mockAppId, mockAppObjectId, mockTenantId, seedMockData } from "./data/test-data";
import { getTopLevelTreeItem } from "./test-utils";
import { AzureCliAccountProvider } from "../utils/azure-cli-account-provider";
import { AZURE_AND_ENTRA_PORTAL_APP_PATH, AZURE_PORTAL_ROOT, ENTRA_PORTAL_ROOT } from "../constants";

// Create Jest mocks
jest.mock("vscode");
jest.mock("../repositories/graph-api-repository");
jest.mock("../utils/exec-shell-cmd");

// Create the test suite for application service
describe("Application Service Tests", () => {
	// Create instances of objects used in the tests
	const graphApiRepository = new GraphApiRepository();
	const treeDataProvider = new AppRegTreeDataProvider(graphApiRepository);
	const accountProvider = new AzureCliAccountProvider();
	const applicationService = new ApplicationService(graphApiRepository, treeDataProvider, accountProvider);

	// Create spy variables
	let triggerCompleteSpy: jest.SpyInstance<any, unknown[], any>;
	let triggerErrorSpy: jest.SpyInstance<any, unknown[], any>;
	let statusBarSpy: jest.SpyInstance<any, [text: string], any>;
	let openTextDocumentSpy: jest.SpyInstance<any, any, any>;
	let openExternalSpy: jest.SpyInstance<Thenable<boolean>, [target: vscode.Uri], any>;

	// The item to be tested
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

		// The item to be tested
		item = { objectId: mockAppObjectId, appId: mockAppId, contextValue: "APPLICATION", value: "First Test App" };

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

	test("Show endpoints with account command error", async () => {
		// Arrange
		const error = new Error("Show endpoints with account command error");
		await treeDataProvider.render();
		jest.spyOn(treeDataProvider, "getTreeItemChildByContext").mockImplementation((_element: AppRegItem, _context: string) => ({ value: "AzureADMyOrg" }));
		jest.spyOn(accountProvider, "getAccountInformation").mockImplementation(async () => { throw error; });

		// Act
		await applicationService.showEndpoints(item);

		// Assert
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
	});

	test("Show endpoints for multi-tenant app", async () => {
		// Arrange
		await treeDataProvider.render();

		// Act
		await applicationService.showEndpoints(item);

		// Assert
		expect(openTextDocumentSpy).toHaveBeenCalled();
	});

	test("Show endpoints for multi-tenant and personal app", async () => {
		// Arrange
		await treeDataProvider.render();
		jest.spyOn(treeDataProvider, "getTreeItemChildByContext").mockImplementation((_element: AppRegItem, _context: string) => ({ value: "AzureADandPersonalMicrosoftAccount" }));

		// Act
		await applicationService.showEndpoints(item);

		// Assert
		expect(openTextDocumentSpy).toHaveBeenCalled();
	});

	test("Show endpoints for consumer app", async () => {
		// Arrange
		await treeDataProvider.render();
		jest.spyOn(treeDataProvider, "getTreeItemChildByContext").mockImplementation((_element: AppRegItem, _context: string) => ({ value: "PersonalMicrosoftAccount" }));

		// Act
		await applicationService.showEndpoints(item);

		// Assert
		expect(openTextDocumentSpy).toHaveBeenCalled();
	});

	test("Show endpoints for consumer app", async () => {
		// Arrange
		await treeDataProvider.render();
		jest.spyOn(treeDataProvider, "getTreeItemChildByContext").mockImplementation((_element: AppRegItem, _context: string) => ({ value: "AzureADMyOrg" }));

		// Act
		await applicationService.showEndpoints(item);

		// Assert
		expect(openTextDocumentSpy).toHaveBeenCalled();
	});

	test("Show endpoints for unknown app type", async () => {
		// Arrange
		await treeDataProvider.render();
		jest.spyOn(treeDataProvider, "getTreeItemChildByContext").mockImplementation((_element: AppRegItem, _context: string) => ({ value: "Unknown" }));

		// Act
		await applicationService.showEndpoints(item);

		// Assert
		expect(openTextDocumentSpy).toHaveBeenCalled();
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
		const error = new Error("View manifest error");
		jest.spyOn(graphApiRepository, "getApplicationDetailsFull").mockImplementation(async (_id: string) => ({ success: false, error }));

		// Act
		await applicationService.viewManifest(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
	});

	test("Copy client id", async () => {
		// Act
		applicationService.copyClientId(item);

		// Assert
		expect(vscode.env.clipboard.readText()).toEqual(mockAppId);
	});

	test("Open in Azure portal", async () => {
		// Arrange
		jest.spyOn(accountProvider, "getAccountInformation").mockImplementation(async () => ({ tenantId: mockTenantId } as any));

		// Act
		const returnValue = await applicationService.openInAzurePortal(item);

		// Assert
		expect(openExternalSpy).toHaveBeenCalledWith(vscode.Uri.parse(`${AZURE_PORTAL_ROOT}/${mockTenantId}${AZURE_AND_ENTRA_PORTAL_APP_PATH}${mockAppId}`));
		expect(returnValue).toEqual(true);
	});

	test("Open in Azure portal without tenant id", async () => {
		// Arrange
		jest.spyOn(vscode.workspace, "getConfiguration").mockImplementation(() => { return { get: (key: string) => { return key === "omitTenantIdFromPortalRequests" ? true : undefined; } } as any; });

		// Act
		const returnValue = await applicationService.openInAzurePortal(item);

		// Assert
		expect(openExternalSpy).toHaveBeenCalledWith(vscode.Uri.parse(`${AZURE_PORTAL_ROOT}${AZURE_AND_ENTRA_PORTAL_APP_PATH}${mockAppId}`));
		expect(returnValue).toEqual(true);
	});

	test("Open in Entra portal", async () => {
		// Arrange
		jest.spyOn(accountProvider, "getAccountInformation").mockImplementation(async () => ({ tenantId: mockTenantId } as any));

		// Act
		const returnValue = await applicationService.openInEntraPortal(item);

		// Assert
		expect(openExternalSpy).toHaveBeenCalledWith(vscode.Uri.parse(`${ENTRA_PORTAL_ROOT}/${mockTenantId}${AZURE_AND_ENTRA_PORTAL_APP_PATH}${mockAppId}`));
		expect(returnValue).toEqual(true);
	});

	test("Edit Logout URL with unset value", async () => {
		// Arrange
		item = { ...item, value: "Not set", contextValue: "LOGOUT-URL" };
		const newLogoutUrl: string = "https://test.com/logout";
		vscode.window.showInputBox = jest.fn().mockResolvedValue(newLogoutUrl);
		jest.spyOn(applicationService, "inputLogoutUrl");

		// Act
		await applicationService.editLogoutUrl(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "LOGOUT-URL-PARENT");
		expect(applicationService.inputLogoutUrl).toHaveBeenCalled();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem!.children![0].label).toEqual("https://test.com/logout");
	});

	test("Edit Logout URL with unset value with validation", async () => {
		// Arrange
		item = { ...item, value: "Not set", contextValue: "LOGOUT-URL" };
		const newLogoutUrl: string = "https://test.com/logout";
		vscode.window.showInputBox = jest.fn().mockResolvedValue(newLogoutUrl);
		const validationSpy = jest.spyOn(validation, "validateLogoutUrl");
		jest.spyOn(applicationService, "inputLogoutUrl").mockImplementation(async (_item: AppRegItem, validation: (value: string) => string | undefined) => {
			const result = validation(newLogoutUrl);
			return result === undefined ? newLogoutUrl : undefined;
		});

		// Act
		await applicationService.editLogoutUrl(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "LOGOUT-URL-PARENT");
		expect(applicationService.inputLogoutUrl).toHaveBeenCalled();
		expect(validationSpy).toHaveBeenCalled();
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

	test("Edit Logout URL with scheme validation error", async () => {
		// Arrange
		item = { ...item, value: "Not set", contextValue: "LOGOUT-URL" };
		const newLogoutUrl: string = "http://test.com/logout";
		vscode.window.showInputBox = jest.fn().mockResolvedValue(newLogoutUrl);
		const validationSpy = jest.spyOn(validation, "validateLogoutUrl");
		jest.spyOn(applicationService, "inputLogoutUrl").mockImplementation(async (_item: AppRegItem, validation: (value: string) => string | undefined) => {
			const result = validation(newLogoutUrl);
			expect(result).toBe("The Logout URL is not valid. It must start with https://.");
			return result === undefined ? newLogoutUrl : undefined;
		});

		// Act
		await applicationService.editLogoutUrl(item);

		// Assert
		expect(applicationService.inputLogoutUrl).toHaveBeenCalled();
		expect(validationSpy).toHaveBeenCalled();
	});

	test("Edit Logout URL with length validation error", async () => {
		// Arrange
		item = { ...item, value: "Not set", contextValue: "LOGOUT-URL" };
		const newLogoutUrl: string = "https://test.com/".padEnd(257, "X");
		vscode.window.showInputBox = jest.fn().mockResolvedValue(newLogoutUrl);
		const validationSpy = jest.spyOn(validation, "validateLogoutUrl");
		jest.spyOn(applicationService, "inputLogoutUrl").mockImplementation(async (_item: AppRegItem, validation: (value: string) => string | undefined) => {
			const result = validation(newLogoutUrl);
			expect(result).toBe("The Logout URL is not valid. A URL cannot be longer than 256 characters.");
			return result === undefined ? newLogoutUrl : undefined;
		});

		// Act
		await applicationService.editLogoutUrl(item);

		// Assert
		expect(applicationService.inputLogoutUrl).toHaveBeenCalled();
		expect(validationSpy).toHaveBeenCalled();
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
		const newAppIdUri: string = "api://test.com";
		item = { ...item, value: "api://oldtest.com", contextValue: "APPID-URI" };
		vscode.window.showInputBox = jest.fn().mockResolvedValue(newAppIdUri);
		jest.spyOn(applicationService, "inputAppIdUri");

		// Act
		await applicationService.editAppIdUri(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "APPID-URI-PARENT");
		expect(applicationService.inputAppIdUri).toHaveBeenCalled();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem!.children![0].label).toEqual(newAppIdUri);
	});

	test("Edit App Id Uri with existing value with validation", async () => {
		// Arrange
		const newAppIdUri: string = "api://test.com";
		item = { ...item, value: "api://oldtest.com", contextValue: "APPID-URI" };
		vscode.window.showInputBox = jest.fn().mockResolvedValue(newAppIdUri);
		const validationSpy = jest.spyOn(validation, "validateAppIdUri");
		jest.spyOn(applicationService, "inputAppIdUri").mockImplementation(async (_item: AppRegItem, signInAudience: string, validation: (value: string, signInAudience: string) => string | undefined) => {
			const result = validation(newAppIdUri, signInAudience);
			return result === undefined ? newAppIdUri : undefined;
		});

		// Act
		await applicationService.editAppIdUri(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "APPID-URI-PARENT");
		expect(applicationService.inputAppIdUri).toHaveBeenCalled();
		expect(validationSpy).toHaveBeenCalled();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem!.children![0].label).toEqual(newAppIdUri);
	});

	test("Edit App Id Uri with trailing slash error for AAD audience", async () => {
		// Arrange
		const newAppIdUri: string = "api://test.com/";
		item = { ...item, value: "api://oldtest.com", contextValue: "APPID-URI" };
		vscode.window.showInputBox = jest.fn().mockResolvedValue(newAppIdUri);
		const validationSpy = jest.spyOn(validation, "validateAppIdUri");
		jest.spyOn(applicationService, "inputAppIdUri").mockImplementation(async (_item: AppRegItem, signInAudience: string, validation: (value: string, signInAudience: string) => string | undefined) => {
			const result = validation(newAppIdUri, signInAudience);
			expect(result).toBe("The Application ID URI cannot end with a trailing slash.");
			return result;
		});

		// Act
		await applicationService.editAppIdUri(item);

		// Assert
		expect(applicationService.inputAppIdUri).toHaveBeenCalled();
		expect(validationSpy).toHaveBeenCalled();
	});

	test("Edit App Id Uri with scheme error for AAD audience", async () => {
		// Arrange
		const newAppIdUri: string = "test.com";
		item = { ...item, value: "api://oldtest.com", contextValue: "APPID-URI" };
		vscode.window.showInputBox = jest.fn().mockResolvedValue(newAppIdUri);
		const validationSpy = jest.spyOn(validation, "validateAppIdUri");
		jest.spyOn(applicationService, "inputAppIdUri").mockImplementation(async (_item: AppRegItem, signInAudience: string, validation: (value: string, signInAudience: string) => string | undefined) => {
			const result = validation(newAppIdUri, signInAudience);
			expect(result).toBe("The Application ID URI is not valid. It must start with http://, https://, api://, MS-APPX://, or customScheme://.");
			return result;
		});

		// Act
		await applicationService.editAppIdUri(item);

		// Assert
		expect(applicationService.inputAppIdUri).toHaveBeenCalled();
		expect(validationSpy).toHaveBeenCalled();
	});

	test("Edit App Id Uri with wildcard error for AAD audience", async () => {
		// Arrange
		const newAppIdUri: string = "api://test.com/*";
		item = { ...item, value: "api://oldtest.com", contextValue: "APPID-URI" };
		vscode.window.showInputBox = jest.fn().mockResolvedValue(newAppIdUri);
		const validationSpy = jest.spyOn(validation, "validateAppIdUri");
		jest.spyOn(applicationService, "inputAppIdUri").mockImplementation(async (_item: AppRegItem, signInAudience: string, validation: (value: string, signInAudience: string) => string | undefined) => {
			const result = validation(newAppIdUri, signInAudience);
			expect(result).toBe("Wildcards are not supported.");
			return result;
		});

		// Act
		await applicationService.editAppIdUri(item);

		// Assert
		expect(applicationService.inputAppIdUri).toHaveBeenCalled();
		expect(validationSpy).toHaveBeenCalled();
	});

	test("Edit App Id Uri with length error for AAD audience", async () => {
		// Arrange
		const newAppIdUri: string = "api://test.com/".padEnd(256, "X");
		item = { ...item, value: "api://oldtest.com", contextValue: "APPID-URI" };
		vscode.window.showInputBox = jest.fn().mockResolvedValue(newAppIdUri);
		const validationSpy = jest.spyOn(validation, "validateAppIdUri");
		jest.spyOn(applicationService, "inputAppIdUri").mockImplementation(async (_item: AppRegItem, signInAudience: string, validation: (value: string, signInAudience: string) => string | undefined) => {
			const result = validation(newAppIdUri, signInAudience);
			expect(result).toBe("The Application ID URI is not valid. A URI cannot be longer than 255 characters.");
			return result;
		});

		// Act
		await applicationService.editAppIdUri(item);

		// Assert
		expect(applicationService.inputAppIdUri).toHaveBeenCalled();
		expect(validationSpy).toHaveBeenCalled();
	});

	test("Edit App Id Uri with scheme error for AAD and consumer audiences", async () => {
		// Arrange
		await treeDataProvider.render();
		const newAppIdUri: string = "cheese://test.com";
		item = { ...item, value: "api://oldtest.com", contextValue: "APPID-URI" };
		vscode.window.showInputBox = jest.fn().mockResolvedValue(newAppIdUri);
		jest.spyOn(treeDataProvider, "getTreeItemChildByContext").mockImplementation((_element: AppRegItem, _context: string) => ({ value: "PersonalMicrosoftAccount" }));
		const validationSpy = jest.spyOn(validation, "validateAppIdUri");
		jest.spyOn(applicationService, "inputAppIdUri").mockImplementation(async (_item: AppRegItem, signInAudience: string, validation: (value: string, signInAudience: string) => string | undefined) => {
			const result = validation(newAppIdUri, signInAudience);
			expect(result).toBe("The Application ID URI is not valid. It must start with http://, https://, or api://.");
			return result;
		});

		// Act
		await applicationService.editAppIdUri(item);

		// Assert
		expect(applicationService.inputAppIdUri).toHaveBeenCalled();
		expect(validationSpy).toHaveBeenCalled();
	});

	test("Edit App Id Uri with wildcard error for AAD and consumer audiences", async () => {
		// Arrange
		await treeDataProvider.render();
		const newAppIdUri: string = "api://test.com/*";
		item = { ...item, value: "api://oldtest.com", contextValue: "APPID-URI" };
		vscode.window.showInputBox = jest.fn().mockResolvedValue(newAppIdUri);
		jest.spyOn(treeDataProvider, "getTreeItemChildByContext").mockImplementation((_element: AppRegItem, _context: string) => ({ value: "PersonalMicrosoftAccount" }));
		const validationSpy = jest.spyOn(validation, "validateAppIdUri");
		jest.spyOn(applicationService, "inputAppIdUri").mockImplementation(async (_item: AppRegItem, signInAudience: string, validation: (value: string, signInAudience: string) => string | undefined) => {
			const result = validation(newAppIdUri, signInAudience);
			expect(result).toBe("Wildcards, fragments, and query strings are not supported.");
			return result;
		});

		// Act
		await applicationService.editAppIdUri(item);

		// Assert
		expect(applicationService.inputAppIdUri).toHaveBeenCalled();
		expect(validationSpy).toHaveBeenCalled();
	});

	test("Edit App Id Uri with length error for AAD and consumer audiences", async () => {
		// Arrange
		await treeDataProvider.render();
		const newAppIdUri: string = "api://test.com/".padEnd(121, "X");
		item = { ...item, value: "api://oldtest.com", contextValue: "APPID-URI" };
		vscode.window.showInputBox = jest.fn().mockResolvedValue(newAppIdUri);
		jest.spyOn(treeDataProvider, "getTreeItemChildByContext").mockImplementation((_element: AppRegItem, _context: string) => ({ value: "PersonalMicrosoftAccount" }));
		const validationSpy = jest.spyOn(validation, "validateAppIdUri");
		jest.spyOn(applicationService, "inputAppIdUri").mockImplementation(async (_item: AppRegItem, signInAudience: string, validation: (value: string, signInAudience: string) => string | undefined) => {
			const result = validation(newAppIdUri, signInAudience);
			expect(result).toBe("The Application ID URI is not valid. A URI cannot be longer than 120 characters.");
			return result;
		});

		// Act
		await applicationService.editAppIdUri(item);

		// Assert
		expect(applicationService.inputAppIdUri).toHaveBeenCalled();
		expect(validationSpy).toHaveBeenCalled();
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
		const error = new Error("Remove Logout URL update application error");
		jest.spyOn(graphApiRepository, "updateApplication").mockImplementation(async (_id: string, _appChange: Application) => ({ success: false, error }));

		// Act
		await applicationService.removeLogoutUrl(item);

		// Assert
		expect(triggerErrorSpy).toHaveBeenCalledWith(error, undefined);
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
		const error = new Error("Delete application with error");
		jest.spyOn(graphApiRepository, "deleteApplication").mockImplementation(async (_id: string) => ({ success: false, error }));

		// Act
		await applicationService.delete(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
	});

	test("Rename application successfully", async () => {
		// Arrange
		await treeDataProvider.render();
		const newAppName: string = "New Application Name";
		vscode.window.showInputBox = jest.fn().mockResolvedValue(newAppName);
		const validationSpy = jest.spyOn(validation, "validateApplicationDisplayName");
		jest.spyOn(applicationService, "inputDisplayNameForRename").mockImplementation(async (_appName: string, signInAudience: string, validation: (value: string, signInAudience: string) => string | undefined) => {
			const result = validation(newAppName, signInAudience);
			return result === undefined ? newAppName : undefined;
		});

		// Act
		await applicationService.rename(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider);
		expect(applicationService.inputDisplayNameForRename).toHaveBeenCalled();
		expect(validationSpy).toHaveBeenCalled();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.label).toEqual("New Application Name");
	});

	test("Rename application with update error", async () => {
		// Arrange
		const error = new Error("Rename application with update error");
		await treeDataProvider.render();
		jest.spyOn(graphApiRepository, "updateApplication").mockImplementation(async (_id: string, _appChange: Application) => ({ success: false, error }));
		vscode.window.showInputBox = jest.fn().mockResolvedValue("New Application Name");

		// Act
		await applicationService.rename(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalledWith(error, undefined);
	});

	test("Add application successfully", async () => {
		// Arrange
		const newAppName: string = "New Application Name";
		vscode.window.showQuickPick = jest.fn().mockResolvedValue({ value: "AzureADMyOrg" });
		vscode.window.showInputBox = jest.fn().mockResolvedValue(newAppName);
		const validationSpy = jest.spyOn(validation, "validateApplicationDisplayName");
		jest.spyOn(applicationService, "inputDisplayNameForNew").mockImplementation(async (signInAudience: string, validation: (value: string, signInAudience: string) => string | undefined) => {
			const result = validation(newAppName, signInAudience);
			return result === undefined ? newAppName : undefined;
		});

		// Act
		await applicationService.add();

		// Assert
		expect(applicationService.inputDisplayNameForNew).toHaveBeenCalled();
		expect(validationSpy).toHaveBeenCalled();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
	});

	test("Add application with no name", async () => {
		// Arrange
		vscode.window.showQuickPick = jest.fn().mockResolvedValue({ value: "AzureADMyOrg" });
		const validationSpy = jest.spyOn(validation, "validateApplicationDisplayName");
		jest.spyOn(applicationService, "inputDisplayNameForNew").mockImplementation(async (signInAudience: string, validation: (value: string, signInAudience: string) => string | undefined) => {
			const result = validation("", signInAudience);
			expect(result).toBe("An application name must be at least one character.");
			return result;
		});

		// Act
		await applicationService.add();

		// Assert
		expect(applicationService.inputDisplayNameForNew).toHaveBeenCalled();
		expect(validationSpy).toHaveBeenCalled();
	});

	test("Add application with name too long for AAD audience", async () => {
		// Arrange
		vscode.window.showQuickPick = jest.fn().mockResolvedValue({ value: "AzureADMyOrg" });
		const validationSpy = jest.spyOn(validation, "validateApplicationDisplayName");
		jest.spyOn(applicationService, "inputDisplayNameForNew").mockImplementation(async (signInAudience: string, validation: (value: string, signInAudience: string) => string | undefined) => {
			const result = validation("X".padEnd(121, "X"), signInAudience);
			expect(result).toBe("An application name cannot be longer than 120 characters.");
			return result;
		});

		// Act
		await applicationService.add();

		// Assert
		expect(applicationService.inputDisplayNameForNew).toHaveBeenCalled();
		expect(validationSpy).toHaveBeenCalled();
	});

	test("Add application with name too long for AAD and consumer audience", async () => {
		// Arrange
		vscode.window.showQuickPick = jest.fn().mockResolvedValue({ value: "AzureADandPersonalMicrosoftAccount" });
		const validationSpy = jest.spyOn(validation, "validateApplicationDisplayName");
		jest.spyOn(applicationService, "inputDisplayNameForNew").mockImplementation(async (signInAudience: string, validation: (value: string, signInAudience: string) => string | undefined) => {
			const result = validation("X".padEnd(91, "X"), signInAudience);
			expect(result).toBe("An application name cannot be longer than 90 characters.");
			return result;
		});

		// Act
		await applicationService.add();

		// Assert
		expect(applicationService.inputDisplayNameForNew).toHaveBeenCalled();
		expect(validationSpy).toHaveBeenCalled();
	});

	test("Add application with creation error", async () => {
		// Arrange
		const error = new Error("Add application with creation error");
		jest.spyOn(graphApiRepository, "createApplication").mockImplementation(async () => ({ success: false, error }));
		vscode.window.showQuickPick = jest.fn().mockResolvedValue({ value: "AzureADMyOrg" });
		vscode.window.showInputBox = jest.fn().mockResolvedValue("Add Application Name");

		// Act
		await applicationService.add();

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
	});
});
