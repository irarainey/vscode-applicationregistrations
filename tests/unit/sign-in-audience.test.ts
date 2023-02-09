import * as vscode from "vscode";
import { GraphApiRepository } from "../../src/repositories/graph-api-repository";
import { AppRegTreeDataProvider } from "../../src/data/app-reg-tree-data-provider";
import { AppRegItem } from "../../src/models/app-reg-item";
import { SignInAudienceService } from "../../src/services/sign-in-audience";
import { Application } from "@microsoft/microsoft-graph-types";
import { mockAppObjectId } from "./constants";
import { MessageItem, MessageOptions } from "vscode";

// Create Jest mocks
jest.mock("vscode");
jest.mock("../../src/repositories/graph-api-repository");

// Create the test suite for sign in audience service
describe("Sign In Audience Service Tests", () => {
	// Create instances of objects used in the tests
	const graphApiRepository = new GraphApiRepository();
	const treeDataProvider = new AppRegTreeDataProvider(graphApiRepository);
	const signInAudienceService = new SignInAudienceService(graphApiRepository, treeDataProvider);

	// Create spy variables
	let triggerCompleteSpy: jest.SpyInstance<any, unknown[], any>;
	let triggerErrorSpy: jest.SpyInstance<any, unknown[], any>;
	let statusBarSpy: jest.SpyInstance<any, [text: string], any>;
	let iconSpy: jest.SpyInstance<any, [id: string, color?: any | undefined], any>;

	// Create variables used in the tests
	let item: AppRegItem;

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
		triggerCompleteSpy = jest.spyOn(Object.getPrototypeOf(signInAudienceService), "triggerRefresh");
		triggerErrorSpy = jest.spyOn(Object.getPrototypeOf(signInAudienceService), "handleError");
		statusBarSpy = jest.spyOn(vscode.window, "setStatusBarMessage");
		iconSpy = jest.spyOn(vscode, "ThemeIcon");
		item = { objectId: mockAppObjectId, contextValue: "AUDIENCE" };
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
		const treeItem = await getTopLevelTreeItem(item.objectId!, "AUDIENCE-PARENT");
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem!.children![0].label).toEqual("Single Tenant");
	});

	test("Trigger error on unsuccessful edit with a generic error", async () => {
		jest.spyOn(graphApiRepository, "updateApplication").mockImplementation(async (_id: string, _appChange: Application) => ({ success: false, error: new Error("Test Error") }));
		await signInAudienceService.edit(item);
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	test("Trigger error on unsuccessful edit with an authentication error", async () => {
		jest.spyOn(graphApiRepository, "updateApplication").mockImplementation(async (_id: string, _appChange: Application) => ({ success: false, error: new Error("az login") }));
		await signInAudienceService.edit(item);
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	test("Trigger error on unsuccessful edit with an authentication error", async () => {
		jest.spyOn(graphApiRepository, "updateApplication").mockImplementation(async (_id: string, _appChange: Application) => ({ success: false, error: new Error("az account set") }));
		await signInAudienceService.edit(item);
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	test("Trigger error on unsuccessful edit with a sign in audience error and open documentation clicked", async () => {
		jest.spyOn(vscode.window, "showErrorMessage").mockImplementation(async (_message: string, _options: MessageOptions, ..._items: MessageItem[]) => ({ title: "Open Documentation" }));
		jest.spyOn(graphApiRepository, "updateApplication").mockImplementation(async (_id: string, _appChange: Application) => ({ success: false, error: new Error("signInAudience") }));
		await signInAudienceService.edit(item);
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	test("Trigger error on unsuccessful edit with a sign in audience error", async () => {
		jest.spyOn(vscode.window, "showErrorMessage").mockImplementation(async (_message: string, _options: MessageOptions, ..._items: MessageItem[]) => ({ title: "OK" }));
		jest.spyOn(graphApiRepository, "updateApplication").mockImplementation(async (_id: string, _appChange: Application) => ({ success: false, error: new Error("signInAudience") }));
		await signInAudienceService.edit(item);
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	// Get a specific top level tree item
	const getTopLevelTreeItem = async (objectId: string, contextValue: string): Promise<AppRegItem | undefined> => {
		const tree = await treeDataProvider.getChildren();
		const app = tree!.find((x) => x.objectId === objectId);
		return app?.children?.find((x) => x.contextValue === contextValue);
	};
});
