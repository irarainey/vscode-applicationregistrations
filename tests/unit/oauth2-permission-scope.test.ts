import * as vscode from "vscode";
import { GraphApiRepository } from "../../src/repositories/graph-api-repository";
import { AppRegTreeDataProvider } from "../../src/data/app-reg-tree-data-provider";
import { AppRegItem } from "../../src/models/app-reg-item";
import { OAuth2PermissionScopeService } from "../../src/services/oauth2-permission-scope";
import { mockAppObjectId, seedMockData } from "../../src/repositories/__mocks__/mock-graph-data";

// Create Jest mocks
jest.mock("vscode");
jest.mock("../../src/repositories/graph-api-repository");

// Create the test suite for sign in audience service
describe("OAuth2 Permission Scope Service Tests", () => {
	// Create instances of objects used in the tests
	const graphApiRepository = new GraphApiRepository();
	const treeDataProvider = new AppRegTreeDataProvider(graphApiRepository);
	const oauth2PermissionScopeService = new OAuth2PermissionScopeService(graphApiRepository, treeDataProvider);

	// Create spy variables
	let triggerCompleteSpy: jest.SpyInstance<any, unknown[], any>;
	let triggerErrorSpy: jest.SpyInstance<any, unknown[], any>;
	let statusBarSpy: jest.SpyInstance<any, [text: string], any>;
	let iconSpy: jest.SpyInstance<any, [id: string, color?: any | undefined], any>;

	// Create variables used in the tests
	let item: AppRegItem;

	beforeAll(async () => {
		console.error = jest.fn();
	});

	beforeEach(() => {
		seedMockData();
		jest.restoreAllMocks();
		statusBarSpy = jest.spyOn(vscode.window, "setStatusBarMessage");
		iconSpy = jest.spyOn(vscode, "ThemeIcon");
		triggerCompleteSpy = jest.spyOn(Object.getPrototypeOf(oauth2PermissionScopeService), "triggerRefresh");
		triggerErrorSpy = jest.spyOn(Object.getPrototypeOf(oauth2PermissionScopeService), "handleError");
		item = { objectId: mockAppObjectId, contextValue: "AUDIENCE" };
	});

	// Test to see if class can be created
	test("Create class instance", () => {
		expect(oauth2PermissionScopeService).toBeDefined();
	});

	// Get a specific top level tree item
	const getTopLevelTreeItem = async (objectId: string, contextValue: string): Promise<AppRegItem | undefined> => {
		const tree = await treeDataProvider.getChildren();
		const app = tree!.find((x) => x.objectId === objectId);
		return app?.children?.find((x) => x.contextValue === contextValue);
	};
});
