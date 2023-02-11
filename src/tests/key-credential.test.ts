import * as vscode from "vscode";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { KeyCredentialService } from "../services/key-credential";
import { mockAppObjectId, seedMockData } from "./test-data";

// Create Jest mocks
jest.mock("vscode");
jest.mock("../repositories/graph-api-repository");

// Create the test suite for key credential service
describe("Key Credential Service Tests", () => {
	// Create instances of objects used in the tests
	const graphApiRepository = new GraphApiRepository();
	const treeDataProvider = new AppRegTreeDataProvider(graphApiRepository);
	const keyCredentialService = new KeyCredentialService(graphApiRepository, treeDataProvider);

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
		triggerCompleteSpy = jest.spyOn(Object.getPrototypeOf(keyCredentialService), "triggerRefresh");
		triggerErrorSpy = jest.spyOn(Object.getPrototypeOf(keyCredentialService), "handleError");

		// The item to be tested
		item = { objectId: mockAppObjectId, contextValue: "PASSWORD-CREDENTIALS" };
	});

	afterAll(() => {
		// Dispose of the key credential service
		keyCredentialService.dispose();
	});

	test("Create class instance", () => {
		// Assert class has been instantiated
		expect(keyCredentialService).toBeDefined();
	});
});
