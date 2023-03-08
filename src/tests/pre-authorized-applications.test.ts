import * as vscode from "vscode";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { mockAppObjectId, mockExposedApiId, seedMockData } from "./data/test-data";
import { PreAuthorizedApplicationsService } from "../services/pre-authorized-applications";
import { getTopLevelTreeItem } from "./test-utils";

// Create Jest mocks
jest.mock("vscode");
jest.mock("../repositories/graph-api-repository");

// Create the test suite for pre authorized client application service
describe("Pre Authorized Client Application Service Tests", () => {
	// Create instances of objects used in the tests
	const graphApiRepository = new GraphApiRepository();
	const treeDataProvider = new AppRegTreeDataProvider(graphApiRepository);
	const preAuthorizedApplicationsService = new PreAuthorizedApplicationsService(graphApiRepository, treeDataProvider);

	// Create spy variables
	let triggerCompleteSpy: jest.SpyInstance<any, unknown[], any>;
	let triggerErrorSpy: jest.SpyInstance<any, unknown[], any>;
	let triggerTreeErrorSpy: jest.SpyInstance<any, unknown[], any>;
	let statusBarSpy: jest.SpyInstance<any, [text: string], any>;
	let iconSpy: jest.SpyInstance<any, [id: string, color?: any | undefined], any>;

	// Create variables used in the tests
	let item: AppRegItem;

	beforeAll(async () => {
		// Suppress console output
		console.error = jest.fn();
	});

	beforeEach(() => {
		// Reset mock data
		seedMockData();

		//Restore the default mock implementations
		jest.restoreAllMocks();

		// Define a standard mock implementation for the showWarningMessage function
		vscode.window.showWarningMessage = jest.fn().mockResolvedValue("Yes");

		// Define spies on the functions to be tested
		statusBarSpy = jest.spyOn(vscode.window, "setStatusBarMessage");
		iconSpy = jest.spyOn(vscode, "ThemeIcon");
		triggerCompleteSpy = jest.spyOn(Object.getPrototypeOf(preAuthorizedApplicationsService), "triggerRefresh");
		triggerErrorSpy = jest.spyOn(Object.getPrototypeOf(preAuthorizedApplicationsService), "handleError");
		triggerTreeErrorSpy = jest.spyOn(Object.getPrototypeOf(treeDataProvider), "handleError");

		// The item to be tested
		item = { objectId: mockAppObjectId, contextValue: "EXPOSED-API-PERMISSIONS" };
	});

	afterAll(() => {
		preAuthorizedApplicationsService.dispose();
		treeDataProvider.dispose();
	});

	test("Create class instance", () => {
		// Assert class has been instantiated
		expect(preAuthorizedApplicationsService).toBeDefined();
	});

	test("Remove authorised client application successfully", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "AUTHORIZED-CLIENT", label: "Test App" };
		const warningSpy = jest.spyOn(vscode.window, "showWarningMessage");

		// Act
		await preAuthorizedApplicationsService.removeAuthorisedClient(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "EXPOSED-API-PERMISSIONS");
		expect(warningSpy).toHaveBeenCalledWith(`Do you want to remove the Authorized Client Application ${item.label}?`, "Yes", "No");
        expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children?.[0].children).toBeUndefined();
	});

});