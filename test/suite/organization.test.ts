import * as vscode from "vscode";
import { GraphApiRepository } from "../../src/repositories/graph-api-repository";
import { AppRegTreeDataProvider } from "../../src/data/app-reg-tree-data-provider";
import { OrganizationService } from "../../src/services/organization";

// Create Jest mocks
jest.mock("vscode");
jest.mock("../../src/repositories/graph-api-repository");
jest.mock("../../src/utils/exec-shell-cmd", () => {
	return {
		execShellCmd: jest.fn().mockResolvedValue("c7b3da28-01b8-46d3-9523-d1b24cbbde76")
	};
});

// Create the test suite for sign in audience service
describe("Organization Service Tests", () => {
	// Create instances of objects used in the tests
	const graphApiRepository = new GraphApiRepository();
	const treeDataProvider = new AppRegTreeDataProvider(graphApiRepository);
	const organizationService = new OrganizationService(graphApiRepository, treeDataProvider);

	// Create spies
	const statusBarSpy = jest.spyOn(vscode.window, "setStatusBarMessage");
    const triggerErrorSpy = jest.spyOn(Object.getPrototypeOf(organizationService), "handleError");

	// Create common mock functions for all tests
	beforeAll(async () => {});

	// Create a generic item to use in each test
	beforeEach(() => {});

	// Test to see if class can be created
	test("Create class instance", () => {
		expect(organizationService).toBeDefined();
	});

	// Test to see if the status bar message is changed when requesting tenant information
	test("Check status bar message updated on request", async () => {
		await organizationService.showTenantInformation();
		expect(statusBarSpy).toHaveBeenCalled();
	});

    // Test to see if an error is handled if no roles returned
	test("Test roles return error", async () => {
		graphApiRepository.getRoleAssignments = jest.fn().mockResolvedValue({ success: false });
		await organizationService.showTenantInformation();
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	// Test to see if an error is handled if no user is returned
	test("Test user return error", async () => {
		graphApiRepository.getUserInformation = jest.fn().mockResolvedValue({ success: false });
		await organizationService.showTenantInformation();
		expect(triggerErrorSpy).toHaveBeenCalled();
	});
});
