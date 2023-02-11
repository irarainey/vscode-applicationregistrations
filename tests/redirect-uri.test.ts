import * as vscode from "vscode";
import { GraphApiRepository } from "../src/repositories/graph-api-repository";
import { AppRegTreeDataProvider } from "../src/data/tree-data-provider";
import { AppRegItem } from "../src/models/app-reg-item";
import { RedirectUriService } from "../src/services/redirect-uri";
import { mockAppObjectId, seedMockData } from "../src/repositories/__mocks__/test-data";

// Create Jest mocks
jest.mock("vscode");
jest.mock("../src/repositories/graph-api-repository");

// Create the test suite for redirect uri service
describe("Redirect URI Service Tests", () => {
	// Create instances of objects used in the tests
	const graphApiRepository = new GraphApiRepository();
	const treeDataProvider = new AppRegTreeDataProvider(graphApiRepository);
	const redirectUriService = new RedirectUriService(graphApiRepository, treeDataProvider);

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
	});

	beforeEach(() => {
		// Reset mock data
		seedMockData();

		//Restore the default mock implementations
		jest.restoreAllMocks();

		// Define spies on the functions to be tested
		statusBarSpy = jest.spyOn(vscode.window, "setStatusBarMessage");
		iconSpy = jest.spyOn(vscode, "ThemeIcon");
		triggerCompleteSpy = jest.spyOn(Object.getPrototypeOf(redirectUriService), "triggerRefresh");
		triggerErrorSpy = jest.spyOn(Object.getPrototypeOf(redirectUriService), "handleError");

		// The item to be tested
		item = { objectId: mockAppObjectId, contextValue: "WEB-REDIRECT" };
	});

	afterAll(() => {
		// Dispose of the application service
		redirectUriService.dispose();
	});

	test("Create class instance", () => {
		// Assert class has been instantiated
		expect(redirectUriService).toBeDefined();
	});
});
