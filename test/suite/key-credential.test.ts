import * as vscode from "vscode";
import { GraphApiRepository } from "../../src/repositories/graph-api-repository";
import { AppRegTreeDataProvider } from "../../src/data/app-reg-tree-data-provider";
import { Application } from "@microsoft/microsoft-graph-types";
import { GraphResult } from "../../src/types/graph-result";
import { AppRegItem } from "../../src/models/app-reg-item";
import { KeyCredentialService } from "../../src/services/key-credential";

// Create Jest mocks
jest.mock("vscode");
jest.mock("../../src/repositories/graph-api-repository");

// Create the test suite for sign in audience service
describe("Key Credential Service Tests", () => {

    // Define the object id of the mock application
    const mockAppObjectId = "ab4e6904-6629-41c9-91d7-2ec9c7d3e46c";

    // Create instances of objects used in the tests
    const graphApiRepository = new GraphApiRepository();
    const treeDataProvider = new AppRegTreeDataProvider(graphApiRepository);
    const keyCredentialService = new KeyCredentialService(graphApiRepository, treeDataProvider);

    // Create spies
    const statusBarSpy = jest.spyOn(vscode.window, "setStatusBarMessage");
    const iconSpy = jest.spyOn(vscode, "ThemeIcon");
    const triggerCompleteSpy = jest.spyOn(Object.getPrototypeOf(keyCredentialService), "triggerRefresh");
    const triggerErrorSpy = jest.spyOn(Object.getPrototypeOf(keyCredentialService), "handleError");

    // Create variables used in the tests
    let item: AppRegItem;

    // Create common mock functions for all tests
    beforeAll(async () => {

    });

    // Create a generic item to use in each test
    beforeEach(() => {
        item = { objectId: mockAppObjectId, contextValue: "AUDIENCE" };
    });

    // Test to see if class can be created
    test("Create class instance", () => {
        expect(keyCredentialService).toBeDefined();
    });

    // Get a specific tree item
    const getTreeItem = async (objectId: string, contextValue: string): Promise<AppRegItem | undefined> => {
        const tree = await treeDataProvider.getChildren();
        const app = tree!.find(x => x.objectId === objectId);
        return app?.children?.find(x => x.contextValue === contextValue);
    };    
});

