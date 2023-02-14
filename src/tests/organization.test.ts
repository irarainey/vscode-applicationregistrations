import * as vscode from "vscode";
import * as execShellCmdModule from "../utils/exec-shell-cmd";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import { OrganizationService } from "../services/organization";
import { mockTenantId, seedMockData } from "./data/test-data";
import { TextDocumentContentProvider } from "vscode";
import { AzureCliAccountProvider } from "../utils/azure-cli-account-provider";
import { AccountInformation } from "../models/account-information";

// Create Jest mocks
jest.mock("vscode");
jest.mock("../repositories/graph-api-repository");
jest.mock("../utils/exec-shell-cmd");

// Create the test suite for sign in audience service
describe("Organization Service Tests", () => {
	// Create instances of objects used in the tests
	const graphApiRepository = new GraphApiRepository();
	const accountProvider = new AzureCliAccountProvider();
	const treeDataProvider = new AppRegTreeDataProvider(graphApiRepository);
	const organizationService = new OrganizationService(graphApiRepository, treeDataProvider, accountProvider);

	// Create spies
	let triggerErrorSpy: jest.SpyInstance<any, unknown[], any>;
	let statusBarSpy: jest.SpyInstance<any, [text: string], any>;
	let openTextDocumentSpy: jest.SpyInstance<any, any, any>;

	beforeAll(() => {
		// Suppress console output
		console.error = jest.fn();
	});

	beforeEach(() => {
		// Reset mock data
		seedMockData();

		//Restore the default mock implementations
		jest.restoreAllMocks();

		// Define a mock implementation for the execShellCmd function
		jest.spyOn(execShellCmdModule, "execShellCmd").mockImplementation(async (_cmd: string) => mockTenantId);

		// Define spies on the functions to be tested
		triggerErrorSpy = jest.spyOn(Object.getPrototypeOf(organizationService), "handleError");
		statusBarSpy = jest.spyOn(vscode.window, "setStatusBarMessage");
		openTextDocumentSpy = jest.spyOn(vscode.workspace, "openTextDocument");
	});

	afterAll(() => {
		// Dispose of the service
		organizationService.dispose();
	});

	test("Create class instance", () => {
		// Assert that the class instance is instantiated
		expect(organizationService).toBeDefined();
	});

	test("Check status bar message updated on request and open text document is called", async () => {
		// Arrange
		const registerTextDocumentContentProviderSpy = jest.spyOn(vscode.workspace, "registerTextDocumentContentProvider").mockImplementation((_scheme: string, _provider: TextDocumentContentProvider) => {
			return { dispose: jest.fn() };
		});
		jest.spyOn(accountProvider, "getAccountInformation").mockImplementation(async () => ({ tenantId: mockTenantId } as any));

		// Act
		await organizationService.showTenantInformation();

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(registerTextDocumentContentProviderSpy).toHaveBeenCalled();
		expect(openTextDocumentSpy).toHaveBeenCalled();
	});

	test("CLI returns error", async () => {
		// Arrange
		jest.spyOn(execShellCmdModule, "execShellCmd").mockImplementation(async (_cmd: string) => {
			throw new Error("Test error");
		});

		// Act
		await organizationService.showTenantInformation();

		// Assert
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	test("User return error", async () => {
		// Arrange
		jest.spyOn(graphApiRepository, "getUserInformation").mockImplementation(async () => ({ success: false, error: new Error("Test error") }));

		// Act
		await organizationService.showTenantInformation();

		// Assert
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	test("Roles return error", async () => {
		// Arrange
		jest.spyOn(graphApiRepository, "getRoleAssignments").mockImplementation(async (_id: string) => ({ success: false, error: new Error("Test error") }));

		// Act
		await organizationService.showTenantInformation();

		// Assert
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	test("Tenant information return error", async () => {
		// Arrange
		jest.spyOn(graphApiRepository, "getTenantInformation").mockImplementation(async (_id: string) => ({ success: false, error: new Error("Test error") }));

		// Act
		await organizationService.showTenantInformation();

		// Assert
		expect(triggerErrorSpy).toHaveBeenCalled();
	});
});
