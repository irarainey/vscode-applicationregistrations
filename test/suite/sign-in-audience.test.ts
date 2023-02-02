import * as vscode from "vscode";
import { GraphApiRepository } from "../../src/repositories/graph-api-repository";
import { AppRegTreeDataProvider } from "../../src/data/app-reg-tree-data-provider";
import { SignInAudienceService } from "../../src/services/sign-in-audience";
import { Application } from "@microsoft/microsoft-graph-types";
import { GraphResult } from "../../src/types/graph-result";
import { AppRegItem } from "../../src/models/app-reg-item";

// Create Jest mocks
jest.mock("vscode");
jest.mock("../../src/repositories/graph-api-repository");

// Create the test suite for sign in audience service
describe("Sign In Audience Service Tests", () => {

    // Define the object id of the mock application
    const mockAppObjectId = "ab4e6904-6629-41c9-91d7-2ec9c7d3e46c";

    // Create instances of objects used in the tests
    const graphApiRepository = new GraphApiRepository();
    const treeDataProvider = new AppRegTreeDataProvider(graphApiRepository);
    const signInAudienceService = new SignInAudienceService(graphApiRepository, treeDataProvider);

    // Create spies
    const statusBarSpy = jest.spyOn(vscode.window, "setStatusBarMessage");
    const iconSpy = jest.spyOn(vscode, "ThemeIcon");
    const triggerCompleteSpy = jest.spyOn(Object.getPrototypeOf(signInAudienceService), "triggerOnComplete");
    const triggerErrorSpy = jest.spyOn(Object.getPrototypeOf(signInAudienceService), "triggerOnError");

    // Create variables used in the tests
    let item: AppRegItem;

    // Create common mock functions for all tests
    beforeAll(async () => {
        await treeDataProvider.render();
    });

    // Create a generic item to use in each test
    beforeEach(() => {
        item = { objectId: mockAppObjectId, contextValue: "AUDIENCE" };
    });

    // Test to see if class can be created
    test("Create class instance", () => {
        expect(signInAudienceService).toBeDefined();
    });

    // Test to see if the status bar message is changed when editing
    test("Check status bar message and icon updated on edit", async () => {
        await signInAudienceService.edit(item);
        expect(statusBarSpy).toHaveBeenCalled();
        expect(iconSpy).toHaveBeenCalled();
    });

    // Test to see if trigger on complete function is called on successful edit after selecting a sign in item
    test("Trigger complete on successful item edit", async () => {
        await signInAudienceService.edit(item);
        expect(triggerCompleteSpy).toHaveBeenCalled();
    });

    // Test to see if trigger on complete function is called on successful edit after selecting a parent item
    test("Trigger complete on successful parent item edit", async () => {
        item = { objectId: mockAppObjectId, contextValue: "AUDIENCE-PARENT", children: [{ objectId: mockAppObjectId, contextValue: "AUDIENCE" }] };
        await signInAudienceService.edit(item);
        expect(triggerCompleteSpy).toHaveBeenCalled();
    });

    // Test to see if trigger on error function is called on unsuccessful edit
    test("Trigger error on unsuccessful edit", async () => {
        graphApiRepository.updateApplication = jest.fn(async (id: string, application: Application): Promise<GraphResult<void>> => {
            return { success: false };
        });
        await signInAudienceService.edit(item);
        expect(triggerErrorSpy).toHaveBeenCalled();
    });

    // Test to see if sign in audience can be changed
    test("Update Sign In Audience", async () => {
        await signInAudienceService.edit(item);
        await treeDataProvider.render();
        const tree = await treeDataProvider.getChildren();
        const app = tree!.find(x => x.objectId === item.objectId);
        const signInAudience = app?.children?.find(x => x.contextValue === "AUDIENCE-PARENT");
        expect(triggerCompleteSpy).toHaveBeenCalled();
        expect(signInAudience!.children![0].label).toBe("Single Tenant");
    });
});