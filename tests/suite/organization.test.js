"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const vscode = require("vscode");
const graph_api_repository_1 = require("../../src/repositories/graph-api-repository");
const app_reg_tree_data_provider_1 = require("../../src/data/app-reg-tree-data-provider");
const organization_1 = require("../../src/services/organization");
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
    const graphApiRepository = new graph_api_repository_1.GraphApiRepository();
    const treeDataProvider = new app_reg_tree_data_provider_1.AppRegTreeDataProvider(graphApiRepository);
    const organizationService = new organization_1.OrganizationService(graphApiRepository, treeDataProvider);
    // Create spies
    let triggerErrorSpy;
    let statusBarSpy;
    // Create common mock functions for all tests
    beforeAll(() => {
        console.error = jest.fn();
    });
    // Clear all mocks before each test
    beforeEach(() => {
        jest.restoreAllMocks();
        triggerErrorSpy = jest.spyOn(Object.getPrototypeOf(organizationService), "handleError");
        statusBarSpy = jest.spyOn(vscode.window, "setStatusBarMessage");
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
    // Test to see if an error is handled if no user is returned
    test("Test user return error", async () => {
        graphApiRepository.getUserInformation = jest.fn().mockResolvedValue({ success: false, error: new Error("Test error") });
        await organizationService.showTenantInformation();
        expect(triggerErrorSpy).toHaveBeenCalled();
    });
    // Test to see if an error is handled if no roles returned
    test("Test roles return error", async () => {
        graphApiRepository.getRoleAssignments = jest.fn().mockResolvedValue({ success: false, error: new Error("Test error") });
        await organizationService.showTenantInformation();
        expect(triggerErrorSpy).toHaveBeenCalled();
    });
    // Test to see if an error is handled if no roles returned
    test("Test tenant information return error", async () => {
        graphApiRepository.getTenantInformation = jest.fn().mockResolvedValue({ success: false, error: new Error("Test error") });
        await organizationService.showTenantInformation();
        expect(triggerErrorSpy).toHaveBeenCalled();
    });
});
//# sourceMappingURL=organization.test.js.map