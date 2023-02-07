"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const vscode = require("vscode");
const graph_api_repository_1 = require("../../src/repositories/graph-api-repository");
const app_reg_tree_data_provider_1 = require("../../src/data/app-reg-tree-data-provider");
const required_resource_access_1 = require("../../src/services/required-resource-access");
// Create Jest mocks
jest.mock("vscode");
jest.mock("../../src/repositories/graph-api-repository");
// Create the test suite for sign in audience service
describe("Required Resource Access Service Tests", () => {
    // Define the object id of the mock application
    const mockAppObjectId = "ab4e6904-6629-41c9-91d7-2ec9c7d3e46c";
    // Create instances of objects used in the tests
    const graphApiRepository = new graph_api_repository_1.GraphApiRepository();
    const treeDataProvider = new app_reg_tree_data_provider_1.AppRegTreeDataProvider(graphApiRepository);
    const requiredResourceAccessService = new required_resource_access_1.RequiredResourceAccessService(graphApiRepository, treeDataProvider);
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
    });
    // Create a generic item to use in each test
    beforeEach(() => {
        jest.restoreAllMocks();
        statusBarSpy = jest.spyOn(vscode.window, "setStatusBarMessage");
        iconSpy = jest.spyOn(vscode, "ThemeIcon");
        triggerCompleteSpy = jest.spyOn(Object.getPrototypeOf(requiredResourceAccessService), "triggerRefresh");
        triggerErrorSpy = jest.spyOn(Object.getPrototypeOf(requiredResourceAccessService), "handleError");
        item = { objectId: mockAppObjectId, contextValue: "AUDIENCE" };
    });
    // Test to see if class can be created
    test("Create class instance", () => {
        expect(requiredResourceAccessService).toBeDefined();
    });
    // Get a specific tree item
    const getTreeItem = async (objectId, contextValue) => {
        const tree = await treeDataProvider.getChildren();
        const app = tree.find(x => x.objectId === objectId);
        return app?.children?.find(x => x.contextValue === contextValue);
    };
});
//# sourceMappingURL=required-resource-access.test.js.map