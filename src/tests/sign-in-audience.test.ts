import * as vscode from "vscode";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { SignInAudienceService } from "../services/sign-in-audience";
import { Application } from "@microsoft/microsoft-graph-types";
import { mockAppObjectId, seedMockData } from "../repositories/__mocks__/test-data";
import { getTopLevelTreeItem } from "./utils";

// Create Jest mocks
jest.mock("vscode");
jest.mock("../repositories/graph-api-repository");

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

	beforeAll(async () => {
		// Suppress console output
		console.error = jest.fn();

		// Mock the showQuickPick function
		vscode.window.showQuickPick = jest.fn().mockResolvedValue({
			label: "Single Tenant",
			description: "Accounts in this organizational directory only.",
			value: "AzureADMyOrg"
		});
	});

	beforeEach(() => {
		// Reset mock data
		seedMockData();

		//Restore the default mock implementations
		jest.restoreAllMocks();

		// Define spies on the functions to be tested
		triggerCompleteSpy = jest.spyOn(Object.getPrototypeOf(signInAudienceService), "triggerRefresh");
		triggerErrorSpy = jest.spyOn(Object.getPrototypeOf(signInAudienceService), "handleError");
		statusBarSpy = jest.spyOn(vscode.window, "setStatusBarMessage");
		iconSpy = jest.spyOn(vscode, "ThemeIcon");

		// Define the base item to be used in the tests
		item = { objectId: mockAppObjectId, contextValue: "AUDIENCE" };
	});

	afterAll(() => {
		// Dispose of the service
		signInAudienceService.dispose();
	});

	test("Create class instance", () => {
		// Assert that the service was instantiated
		expect(signInAudienceService).toBeDefined();
	});

	test("Check status bar message and icon updated on edit", async () => {
		// Act
		await signInAudienceService.edit(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
	});

	test("Trigger complete on successful item edit", async () => {
		// Act
		await signInAudienceService.edit(item);

		// Assert
		expect(triggerCompleteSpy).toHaveBeenCalled();
	});

	test("Trigger complete on successful parent item edit", async () => {
		// Arrange
		item = { ...item, contextValue: "AUDIENCE-PARENT", children: [{ objectId: mockAppObjectId, contextValue: "AUDIENCE" }] };

		// Act
		await signInAudienceService.edit(item);

		// Assert
		expect(triggerCompleteSpy).toHaveBeenCalled();
	});

	test("Update Sign In Audience", async () => {
		// Act
		await signInAudienceService.edit(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(item.objectId!, treeDataProvider, "AUDIENCE-PARENT");
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem!.children![0].label).toEqual("Single Tenant");
	});

	test("Trigger error on unsuccessful edit with a generic error", async () => {
		// Arrange
		jest.spyOn(graphApiRepository, "updateApplication").mockImplementation(async (_id: string, _appChange: Application) => ({ success: false, error: new Error("Test Error") }));

		// Act
		await signInAudienceService.edit(item);

		// Assert
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	test("Trigger error on unsuccessful edit with an authentication error", async () => {
		// Arrange
		jest.spyOn(graphApiRepository, "updateApplication").mockImplementation(async (_id: string, _appChange: Application) => ({ success: false, error: new Error("az login") }));

		// Act
		await signInAudienceService.edit(item);

		// Assert
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	test("Trigger error on unsuccessful edit with an authentication error", async () => {
		// Arrange
		jest.spyOn(graphApiRepository, "updateApplication").mockImplementation(async (_id: string, _appChange: Application) => ({ success: false, error: new Error("az account set") }));

		// Act
		await signInAudienceService.edit(item);

		// Assert
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	test("Trigger error on unsuccessful edit with a sign in audience error and open documentation clicked", async () => {
		// Arrange
		jest.spyOn(vscode.window, "showErrorMessage").mockReturnValue({ then: (callback: any) => callback("Open Documentation") });
		jest.spyOn(graphApiRepository, "updateApplication").mockImplementation(async (_id: string, _appChange: Application) => ({ success: false, error: new Error("signInAudience") }));

		// Act
		await signInAudienceService.edit(item);

		// Assert
		expect(triggerErrorSpy).toHaveBeenCalled();
	});

	test("Trigger error on unsuccessful edit with a sign in audience error", async () => {
		// Arrange
		jest.spyOn(vscode.window, "showErrorMessage").mockReturnValue({ then: (callback: any) => callback("OK") });
		jest.spyOn(graphApiRepository, "updateApplication").mockImplementation(async (_id: string, _appChange: Application) => ({ success: false, error: new Error("signInAudience") }));

		// Act
		await signInAudienceService.edit(item);

		// Assert
		expect(triggerErrorSpy).toHaveBeenCalled();
	});
});
