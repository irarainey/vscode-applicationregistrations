import * as vscode from "vscode";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { OwnerService } from "../services/owner";
import { mockAppId, mockAppObjectId, mockSecondAppObjectId, mockSecondUserId, mockTenantId, mockUserId, seedMockData } from "./data/test-data";
import { getTopLevelTreeItem } from "./test-utils";
import { AzureCliAccountProvider } from "../utils/azure-cli-account-provider";
import { AZURE_AND_ENTRA_PORTAL_USER_PATH, AZURE_PORTAL_ROOT, ENTRA_PORTAL_ROOT } from "../constants";

// Create Jest mocks
jest.mock("vscode");
jest.mock("../repositories/graph-api-repository");

// Create the test suite for sign in audience service
describe("Owner Service Tests", () => {
	// Create instances of objects used in the tests
	const graphApiRepository = new GraphApiRepository();
	const treeDataProvider = new AppRegTreeDataProvider(graphApiRepository);
	const accountProvider = new AzureCliAccountProvider();
	const ownerService = new OwnerService(graphApiRepository, treeDataProvider, accountProvider);

	// Create spy variables
	let triggerCompleteSpy: jest.SpyInstance<any, unknown[], any>;
	let triggerErrorSpy: jest.SpyInstance<any, unknown[], any>;
	let triggerTreeErrorSpy: jest.SpyInstance<any, unknown[], any>;
	let statusBarSpy: jest.SpyInstance<any, [text: string], any>;
	let iconSpy: jest.SpyInstance<any, [id: string, color?: any | undefined], any>;
	let openExternalSpy: jest.SpyInstance<Thenable<boolean>, [target: vscode.Uri], any>;

	// The item to be tested
	let item: AppRegItem = { objectId: mockAppObjectId, appId: mockAppId, userId: mockUserId, contextValue: "OWNER" };

	beforeAll(async () => {
		// Suppress console output
		console.error = jest.fn();
	});

	beforeEach(() => {
		// Reset mock data
		seedMockData();

		// Restore the default mock implementations
		jest.restoreAllMocks();

		// Define a standard mock implementation for the showWarningMessage function
		vscode.window.showWarningMessage = jest.fn().mockResolvedValue("Yes");

		// Define spies on the functions to be tested
		statusBarSpy = jest.spyOn(vscode.window, "setStatusBarMessage");
		iconSpy = jest.spyOn(vscode, "ThemeIcon");
		triggerCompleteSpy = jest.spyOn(Object.getPrototypeOf(ownerService), "triggerRefresh");
		triggerErrorSpy = jest.spyOn(Object.getPrototypeOf(ownerService), "handleError");
		triggerTreeErrorSpy = jest.spyOn(Object.getPrototypeOf(treeDataProvider), "handleError");
		openExternalSpy = jest.spyOn(vscode.env, "openExternal");
	});

	afterAll(() => {
		ownerService.dispose();
		treeDataProvider.dispose();
	});

	test("Create class instance", () => {
		// Assert that the owner service is instantiated
		expect(ownerService).toBeDefined();
	});

	test("Open in Azure portal", async () => {
		// Arrange
		jest.spyOn(accountProvider, "getAccountInformation").mockImplementation(async () => ({ tenantId: mockTenantId } as any));

		// Act
		const returnValue = await ownerService.openInAzurePortal(item);

		// Assert
		expect(openExternalSpy).toHaveBeenCalledWith(vscode.Uri.parse(`${AZURE_PORTAL_ROOT}/${mockTenantId}${AZURE_AND_ENTRA_PORTAL_USER_PATH}${mockUserId}`));
		expect(returnValue).toEqual(true);
	});

	test("Open in Entra portal", async () => {
		// Arrange
		jest.spyOn(accountProvider, "getAccountInformation").mockImplementation(async () => ({ tenantId: mockTenantId } as any));

		// Act
		const returnValue = await ownerService.openInEntraPortal(item);

		// Assert
		expect(openExternalSpy).toHaveBeenCalledWith(vscode.Uri.parse(`${ENTRA_PORTAL_ROOT}/${mockTenantId}${AZURE_AND_ENTRA_PORTAL_USER_PATH}${mockUserId}`));
		expect(returnValue).toEqual(true);
	});

	test("Remove owner", async () => {
		// Act
		await ownerService.remove(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "OWNERS");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem!.children!.length).toEqual(1);
	});

	test("Remove owner with error", async () => {
		// Arrange
		const error = new Error("Remove owner with error");
		jest.spyOn(graphApiRepository, "removeApplicationOwner").mockImplementation(async (_id: string, _userId: string) => ({ success: false, error }));

		// Act
		await ownerService.remove(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
	});

	test("Add owner successfully", async () => {
		// Arrange
		item = { objectId: mockSecondAppObjectId, contextValue: "OWNERS" };
		vscode.window.showInputBox = jest.fn().mockResolvedValue("Second User");
		ownerService.setUserList([{ id: mockSecondUserId }]);

		// Act
		await ownerService.add(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockSecondAppObjectId, treeDataProvider, "OWNERS");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem!.children!.length).toEqual(2);
		expect(treeItem!.children![1].userId).toEqual(mockSecondUserId);
	});

	test("Add owner with add error", async () => {
		// Arrange
		const error = new Error("Add owner with add error");
		item = { objectId: mockSecondAppObjectId, contextValue: "OWNERS" };
		vscode.window.showInputBox = jest.fn().mockResolvedValue("Second User");
		jest.spyOn(graphApiRepository, "addApplicationOwner").mockImplementation(async (_id: string, _userId: string) => ({ success: false, error }));
		ownerService.setUserList([{ id: mockSecondUserId }]);

		// Act
		await ownerService.add(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
	});

	test("Add owner with error getting existing owners", async () => {
		// Arrange
		const error = new Error("Add owner with error getting existing owners");
		jest.spyOn(graphApiRepository, "getApplicationOwners").mockImplementation(async (_id: string) => ({ success: false, error }));

		// Act
		await ownerService.add(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
	});

	test("Error getting owner children", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "OWNERS" };
		const error = new Error("Error getting owner children");
		jest.spyOn(graphApiRepository, "getApplicationOwners").mockImplementation(async (_id: string) => ({ success: false, error }));

		// Act
		await treeDataProvider.render();
		await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "OWNERS");

		// Assert
		expect(triggerTreeErrorSpy).toHaveBeenCalledWith(error);
	});
});
