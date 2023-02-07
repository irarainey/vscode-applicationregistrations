"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const vscode = require("vscode");
const graph_api_repository_1 = require("../../src/repositories/graph-api-repository");
const app_reg_tree_data_provider_1 = require("../../src/data/app-reg-tree-data-provider");
const sign_in_audience_1 = require("../../src/services/sign-in-audience");
// Create Jest mocks
jest.mock("vscode");
jest.mock("../../src/repositories/graph-api-repository");
// Create the test suite for sign in audience service
describe("Sign In Audience Service Tests", () => {
    // Define the object id of the mock application
    const mockAppObjectId = "ab4e6904-6629-41c9-91d7-2ec9c7d3e46c";
    // Create instances of objects used in the tests
    const graphApiRepository = new graph_api_repository_1.GraphApiRepository();
    const treeDataProvider = new app_reg_tree_data_provider_1.AppRegTreeDataProvider(graphApiRepository);
    const signInAudienceService = new sign_in_audience_1.SignInAudienceService(graphApiRepository, treeDataProvider);
    // Create spy variables
    let triggerCompleteSpy;
    let triggerErrorSpy;
    let statusBarSpy;
    let iconSpy;
    // Create variables used in the tests
    let item;
    // Create common mock functions for all tests
    beforeAll(async () => {
        console.error = jest.fn();
        vscode.window.showQuickPick = jest.fn().mockResolvedValue({
            label: "Single Tenant",
            description: "Accounts in this organizational directory only.",
            value: "AzureADMyOrg"
        });
    });
    // Create a generic item to use in each test
    beforeEach(() => {
        jest.restoreAllMocks();
        item = { objectId: mockAppObjectId, contextValue: "AUDIENCE" };
        triggerCompleteSpy = jest.spyOn(Object.getPrototypeOf(signInAudienceService), "triggerRefresh");
        triggerErrorSpy = jest.spyOn(Object.getPrototypeOf(signInAudienceService), "handleError");
        statusBarSpy = jest.spyOn(vscode.window, "setStatusBarMessage");
        iconSpy = jest.spyOn(vscode, "ThemeIcon");
    });
    afterAll(() => {
        signInAudienceService.dispose();
    });
    test("Create class instance", () => {
        expect(signInAudienceService).toBeDefined();
    });
    test("Check status bar message and icon updated on edit", async () => {
        await signInAudienceService.edit(item);
        expect(statusBarSpy).toHaveBeenCalled();
        expect(iconSpy).toHaveBeenCalled();
    });
    test("Trigger complete on successful item edit", async () => {
        await signInAudienceService.edit(item);
        expect(triggerCompleteSpy).toHaveBeenCalled();
    });
    test("Trigger complete on successful parent item edit", async () => {
        item = { objectId: mockAppObjectId, contextValue: "AUDIENCE-PARENT", children: [{ objectId: mockAppObjectId, contextValue: "AUDIENCE" }] };
        await signInAudienceService.edit(item);
        expect(triggerCompleteSpy).toHaveBeenCalled();
    });
    test("Update Sign In Audience", async () => {
        await signInAudienceService.edit(item);
        const signInAudience = await getTreeItem(item.objectId, "AUDIENCE-PARENT");
        expect(triggerCompleteSpy).toHaveBeenCalled();
        expect(signInAudience.children[0].label).toBe("Single Tenant");
    });
    test("Trigger error on unsuccessful edit with a generic error", async () => {
        graphApiRepository.updateApplication = jest.fn().mockResolvedValue({ success: false, error: new Error("Generic error") });
        await signInAudienceService.edit(item);
        expect(triggerErrorSpy).toHaveBeenCalled();
    });
    test("Trigger error on unsuccessful edit with an authentication error", async () => {
        graphApiRepository.updateApplication = jest.fn().mockResolvedValue({ success: false, error: new Error("az login") });
        await signInAudienceService.edit(item);
        expect(triggerErrorSpy).toHaveBeenCalled();
    });
    test("Trigger error on unsuccessful edit with an authentication error", async () => {
        graphApiRepository.updateApplication = jest.fn().mockResolvedValue({ success: false, error: { name: "Error", message: "az account set" } });
        await signInAudienceService.edit(item);
        expect(triggerErrorSpy).toHaveBeenCalled();
    });
    test("Trigger error on unsuccessful edit with a sign in audience error and open documentation clicked", async () => {
        vscode.window.showErrorMessage = jest.fn().mockResolvedValue("Open Documentation");
        graphApiRepository.updateApplication = jest.fn().mockResolvedValue({ success: false, error: new Error("signInAudience") });
        await signInAudienceService.edit(item);
        expect(triggerErrorSpy).toHaveBeenCalled();
    });
    test("Trigger error on unsuccessful edit with a sign in audience error", async () => {
        vscode.window.showErrorMessage = jest.fn().mockResolvedValue("OK");
        graphApiRepository.updateApplication = jest.fn().mockResolvedValue({ success: false, error: new Error("signInAudience") });
        await signInAudienceService.edit(item);
        expect(triggerErrorSpy).toHaveBeenCalled();
    });
    // Get a specific tree item
    const getTreeItem = async (objectId, contextValue) => {
        const tree = await treeDataProvider.getChildren();
        const app = tree.find(x => x.objectId === objectId);
        return app?.children?.find(x => x.contextValue === contextValue);
    };
});
//# sourceMappingURL=sign-in-audience.test.js.map