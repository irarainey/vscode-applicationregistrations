import * as vscode from "vscode";
import * as execShellCmdModule from "../../src/utils/exec-shell-cmd";
import { GraphApiRepository } from "../../src/repositories/graph-api-repository";
import { AppRegTreeDataProvider } from "../../src/data/app-reg-tree-data-provider";
import { OrganizationService } from "../../src/services/organization";
import { mockTenantId } from "./constants";

// Create Jest mocks
jest.mock("vscode");
jest.mock("../../src/repositories/graph-api-repository");
jest.mock("../../src/utils/exec-shell-cmd");

// Create the test suite for sign in audience service
describe("Organization Service Tests", () => {
	// Create instances of objects used in the tests
	const graphApiRepository = new GraphApiRepository();
	const treeDataProvider = new AppRegTreeDataProvider(graphApiRepository);
	const organizationService = new OrganizationService(graphApiRepository, treeDataProvider);

	// Create spies
	let triggerErrorSpy: jest.SpyInstance<any, unknown[], any>;
	let statusBarSpy: jest.SpyInstance<any, [text: string], any>;

	// Create common mock functions for all tests
	beforeAll(() => {
		console.error = jest.fn();
	});

	// Clear all mocks before each test
	beforeEach(() => {
		jest.restoreAllMocks();
		jest.spyOn(execShellCmdModule, "execShellCmd").mockImplementation(async (_cmd: string) => mockTenantId);
		triggerErrorSpy = jest.spyOn(Object.getPrototypeOf(organizationService), "handleError");
		statusBarSpy = jest.spyOn(vscode.window, "setStatusBarMessage");
	});

	afterAll(() => {
		organizationService.dispose();
	});

	// Test to see if class can be created
	test("Create class instance", () => {
		expect(organizationService).toBeDefined();
	});

	// Test to see if the status bar message is changed when requesting tenant information
	test("Check status bar message updated on request", async () => {
		await organizationService.showTenantInformation();
		expect(statusBarSpy).toHaveBeenCalled();
	});

	// Test to see if the status bar message is changed when requesting tenant information
	test("CLI returns error", async () => {
		jest.spyOn(execShellCmdModule, "execShellCmd").mockImplementation(async (_cmd: string) => { throw new Error("Test error"); });
		await organizationService.showTenantInformation();
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	// Test to see if an error is handled if no user is returned
	test("User return error", async () => {
		jest.spyOn(graphApiRepository, "getUserInformation").mockImplementation(async () => ({ success: false, error: new Error("Test error") }));
		await organizationService.showTenantInformation();
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	// Test to see if an error is handled if no roles returned
	test("Roles return error", async () => {
		jest.spyOn(graphApiRepository, "getRoleAssignments").mockImplementation(async (_id: string) => ({ success: false, error: new Error("Test error") }));
		await organizationService.showTenantInformation();
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	// Test to see if an error is handled if no roles returned
	test("Tenant information return error", async () => {
		jest.spyOn(graphApiRepository, "getTenantInformation").mockImplementation(async (_id: string) => ({ success: false, error: new Error("Test error") }));
		await organizationService.showTenantInformation();
		expect(triggerErrorSpy).toHaveBeenCalled();
	});
});

