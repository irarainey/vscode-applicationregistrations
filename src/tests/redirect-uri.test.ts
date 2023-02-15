import * as vscode from "vscode";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { RedirectUriService } from "../services/redirect-uri";
import { mockAppObjectId, seedMockData } from "./data/test-data";
import { getTopLevelTreeItem } from "./test-utils";

// Create Jest mocks
jest.mock("vscode");
jest.mock("../repositories/graph-api-repository");

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
	let treeSpy: jest.SpyInstance<Promise<void>, [status?: string | undefined, type?: string | undefined], any>;

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

		// Define a standard mock implementation for the dialog functions
		vscode.window.showWarningMessage = jest.fn().mockResolvedValue("Yes");
		vscode.window.showInputBox = jest.fn().mockResolvedValue("https://test.com/callback");

		// Define spies on the functions to be tested
		statusBarSpy = jest.spyOn(vscode.window, "setStatusBarMessage");
		iconSpy = jest.spyOn(vscode, "ThemeIcon");
		triggerCompleteSpy = jest.spyOn(Object.getPrototypeOf(redirectUriService), "triggerRefresh");
		triggerErrorSpy = jest.spyOn(Object.getPrototypeOf(redirectUriService), "handleError");
		treeSpy = jest.spyOn(treeDataProvider, "render");

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

	test("Delete web redirect successfully", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "WEB-REDIRECT-URI", label: "https://sample.com/callback" };

		// Act
		await redirectUriService.delete(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "WEB-REDIRECT");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children?.length).toEqual(0);
	});

	test("Delete all web redirects successfully", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "WEB-REDIRECT" };

		// Act
		await redirectUriService.deleteAll(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "WEB-REDIRECT");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children?.length).toEqual(0);
	});

	test("Delete web redirect with get existing error", async () => {
		// Arrange
		const error = new Error("Delete web redirect with get existing error");
		jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementation(async (_id: string) => ({ success: false, error }));
		item = { objectId: mockAppObjectId, contextValue: "WEB-REDIRECT-URI", label: "https://sample.com/callback" };

		// Act
		await redirectUriService.delete(item);

		// Assert
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
		expect(treeSpy).toHaveBeenCalledTimes(0);
	});

	test("Delete spa redirect successfully", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SPA-REDIRECT-URI", label: "https://spa.com" };

		// Act
		await redirectUriService.delete(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "SPA-REDIRECT");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children?.length).toEqual(0);
	});

	test("Delete all spa redirects successfully", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SPA-REDIRECT" };

		// Act
		await redirectUriService.deleteAll(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "SPA-REDIRECT");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children?.length).toEqual(0);
	});

	test("Delete spa redirect with get existing error", async () => {
		// Arrange
		const error = new Error("Delete spa redirect with get existing error");
		jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementation(async (_id: string) => ({ success: false, error }));
		item = { objectId: mockAppObjectId, contextValue: "SPA-REDIRECT-URI", label: "https://spa.com" };

		// Act
		await redirectUriService.delete(item);

		// Assert
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
		expect(treeSpy).toHaveBeenCalledTimes(0);
	});

	test("Delete native redirect successfully", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "NATIVE-REDIRECT-URI", label: "https://mobile.com" };

		// Act
		await redirectUriService.delete(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "NATIVE-REDIRECT");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children?.length).toEqual(0);
	});

	test("Delete all native redirects successfully", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "NATIVE-REDIRECT" };

		// Act
		await redirectUriService.deleteAll(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "NATIVE-REDIRECT");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children?.length).toEqual(0);
	});

	test("Delete native redirect with get existing error", async () => {
		// Arrange
		const error = new Error("Delete native redirect with get existing error");
		jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementation(async (_id: string) => ({ success: false, error }));
		item = { objectId: mockAppObjectId, contextValue: "NATIVE-REDIRECT-URI", label: "https://mobile.com" };

		// Act
		await redirectUriService.delete(item);

		// Assert
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
		expect(treeSpy).toHaveBeenCalledTimes(0);
	});

	test("Edit web redirect successfully", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "WEB-REDIRECT-URI", label: "https://sample.com/callback" };

		// Act
		await redirectUriService.edit(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "WEB-REDIRECT");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children?.length).toEqual(1);
		expect(treeItem?.children?.[0].label).toEqual("https://test.com/callback");
	});

	test("Edit web redirect no existing uris error", async () => {
		// Arrange
		const error = new Error("Edit web redirect no existing uris error");
		jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementation(async (_id: string) => ({ success: false, error }));
		item = { objectId: mockAppObjectId, contextValue: "WEB-REDIRECT-URI", label: "https://sample.com/callback" };

		// Act
		await redirectUriService.edit(item);

		// Assert
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
		expect(treeSpy).toHaveBeenCalledTimes(0);
	});

	test("Edit web redirect update error", async () => {
		// Arrange
		const error = new Error("Edit web redirect update error");
		jest.spyOn(graphApiRepository, "updateApplication").mockImplementation(async (_id: string) => ({ success: false, error }));
		item = { objectId: mockAppObjectId, contextValue: "WEB-REDIRECT-URI", label: "https://sample.com/callback" };

		// Act
		await redirectUriService.edit(item);

		// Assert
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
		expect(treeSpy).toHaveBeenCalledTimes(0);
	});

	test("Edit spa redirect successfully", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SPA-REDIRECT-URI", label: "https://spa.com" };
		vscode.window.showInputBox = jest.fn().mockResolvedValue("https://newspa.com");

		// Act
		await redirectUriService.edit(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "SPA-REDIRECT");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children?.length).toEqual(1);
		expect(treeItem?.children?.[0].label).toEqual("https://newspa.com");
	});

	test("Edit spa redirect no existing uris error", async () => {
		// Arrange
		const error = new Error("Edit spa redirect no existing uris error");
		jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementation(async (_id: string) => ({ success: false, error }));
		vscode.window.showInputBox = jest.fn().mockResolvedValue("https://newspa.com");
		item = { objectId: mockAppObjectId, contextValue: "SPA-REDIRECT-URI", label: "https://spa.com" };

		// Act
		await redirectUriService.edit(item);

		// Assert
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
		expect(treeSpy).toHaveBeenCalledTimes(0);
	});

	test("Edit spa redirect update error", async () => {
		// Arrange
		const error = new Error("Edit spa redirect update error");
		jest.spyOn(graphApiRepository, "updateApplication").mockImplementation(async (_id: string) => ({ success: false, error }));
		vscode.window.showInputBox = jest.fn().mockResolvedValue("https://newspa.com");
		item = { objectId: mockAppObjectId, contextValue: "SPA-REDIRECT-URI", label: "https://spa.com" };

		// Act
		await redirectUriService.edit(item);

		// Assert
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
		expect(treeSpy).toHaveBeenCalledTimes(0);
	});

	test("Edit native redirect successfully", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "NATIVE-REDIRECT-URI", label: "https://mobile.com" };
		vscode.window.showInputBox = jest.fn().mockResolvedValue("https://newmobile.com");

		// Act
		await redirectUriService.edit(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "NATIVE-REDIRECT");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children?.length).toEqual(1);
		expect(treeItem?.children?.[0].label).toEqual("https://newmobile.com");
	});

	test("Edit native redirect no existing uris error", async () => {
		// Arrange
		const error = new Error("Edit native redirect no existing uris error");
		jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementation(async (_id: string) => ({ success: false, error }));
		item = { objectId: mockAppObjectId, contextValue: "NATIVE-REDIRECT-URI", label: "https://mobile.com" };
		vscode.window.showInputBox = jest.fn().mockResolvedValue("https://newmobile.com");

		// Act
		await redirectUriService.edit(item);

		// Assert
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
		expect(treeSpy).toHaveBeenCalledTimes(0);
	});

	test("Edit native redirect update error", async () => {
		// Arrange
		const error = new Error("Edit native redirect update error");
		jest.spyOn(graphApiRepository, "updateApplication").mockImplementation(async (_id: string) => ({ success: false, error }));
		item = { objectId: mockAppObjectId, contextValue: "NATIVE-REDIRECT-URI", label: "https://mobile.com" };
		vscode.window.showInputBox = jest.fn().mockResolvedValue("https://newmobile.com");

		// Act
		await redirectUriService.edit(item);

		// Assert
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
		expect(treeSpy).toHaveBeenCalledTimes(0);
	});

	test("Add web redirect successfully for AAD audiences", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "WEB-REDIRECT" };
		jest.spyOn(graphApiRepository, "getSignInAudience").mockImplementation(async (_id: string) => ({ success: true, value: "AzureADMyOrg" }));

		// Act
		await redirectUriService.add(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "WEB-REDIRECT");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children?.length).toEqual(2);
		expect(treeItem?.children?.[1].label).toEqual("https://test.com/callback");
	});

	test("Add web redirect successfully for AAD and consumer audiences", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "WEB-REDIRECT" };
		jest.spyOn(graphApiRepository, "getSignInAudience").mockImplementation(async (_id: string) => ({ success: true, value: "AzureADandPersonalMicrosoftAccount" }));

		// Act
		await redirectUriService.add(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "WEB-REDIRECT");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children?.length).toEqual(2);
		expect(treeItem?.children?.[1].label).toEqual("https://test.com/callback");
	});

	test("Add web redirect with no existing uris error", async () => {
		// Arrange
		const error = new Error("Add web redirect with no existing uris error");
		jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementation(async (_id: string) => ({ success: false, error }));
		item = { objectId: mockAppObjectId, contextValue: "WEB-REDIRECT" };

		// Act
		await redirectUriService.add(item);

		// Assert
		// Assert
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
		expect(treeSpy).toHaveBeenCalledTimes(0);
	});

	test("Add web redirect with sign in audience error", async () => {
		// Arrange
		const error = new Error("Add web redirect with sign in audience error");
		jest.spyOn(graphApiRepository, "getSignInAudience").mockImplementation(async (_id: string) => ({ success: false, error }));
		item = { objectId: mockAppObjectId, contextValue: "WEB-REDIRECT" };

		// Act
		await redirectUriService.add(item);

		// Assert
		// Assert
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
		expect(treeSpy).toHaveBeenCalledTimes(0);
	});

	test("Add spa redirect successfully", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "SPA-REDIRECT" };
		vscode.window.showInputBox = jest.fn().mockResolvedValue("https://newspa.com");

		// Act
		await redirectUriService.add(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "SPA-REDIRECT");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children?.length).toEqual(2);
		expect(treeItem?.children?.[1].label).toEqual("https://newspa.com");
	});

	test("Add native redirect successfully", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "NATIVE-REDIRECT" };
		vscode.window.showInputBox = jest.fn().mockResolvedValue("https://newmobile.com");
		// Act
		await redirectUriService.add(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "NATIVE-REDIRECT");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children?.length).toEqual(2);
		expect(treeItem?.children?.[1].label).toEqual("https://newmobile.com");
	});
});
