import * as vscode from "vscode";
import * as execShellCmdModule from "../../src/utils/exec-shell-cmd";
import { GraphApiRepository } from "../../src/repositories/graph-api-repository";
import { AppRegTreeDataProvider } from "../../src/data/app-reg-tree-data-provider";
import { OrganizationService } from "../../src/services/organization";
import { mockTenantId, seedMockData } from "../../src/repositories/__mocks__/mock-graph-data";

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
	});

	afterAll(() => {
		// Dispose of the service
		organizationService.dispose();
	});

	test("Create class instance", () => {
		// Assert that the class instance is instantiated
		expect(organizationService).toBeDefined();
	});

	test("Check status bar message updated on request", async () => {
		// Act
		await organizationService.showTenantInformation();

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
	});

	test("CLI returns error", async () => {
		// Arrange
		jest.spyOn(execShellCmdModule, "execShellCmd").mockImplementation(async (_cmd: string) => {	throw new Error("Test error");});

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
