import * as vscode from "vscode";
import * as validation from "../utils/validation";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { RedirectUriService } from "../services/redirect-uri";
import { mockApplications, mockAppObjectId, mockRedirectUris, seedMockData } from "./data/test-data";
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
		const graphSpy = jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementation(async (_id: string) => ({ success: false, error }));
		item = { objectId: mockAppObjectId, contextValue: "WEB-REDIRECT-URI", label: "https://sample.com/callback" };

		// Act
		await redirectUriService.delete(item);

		// Assert
		expect(graphSpy).toHaveBeenCalled();
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
		const graphSpy = jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementation(async (_id: string) => ({ success: false, error }));
		item = { objectId: mockAppObjectId, contextValue: "SPA-REDIRECT-URI", label: "https://spa.com" };

		// Act
		await redirectUriService.delete(item);

		// Assert
		expect(graphSpy).toHaveBeenCalled();
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
		const graphSpy = jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementation(async (_id: string) => ({ success: false, error }));
		item = { objectId: mockAppObjectId, contextValue: "NATIVE-REDIRECT-URI", label: "https://mobile.com" };

		// Act
		await redirectUriService.delete(item);

		// Assert
		expect(graphSpy).toHaveBeenCalled();
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
		const graphSpy = jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementation(async (_id: string) => ({ success: false, error }));
		item = { objectId: mockAppObjectId, contextValue: "WEB-REDIRECT-URI", label: "https://sample.com/callback" };

		// Act
		await redirectUriService.edit(item);

		// Assert
		expect(graphSpy).toHaveBeenCalled();
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
		const graphSpy = jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementation(async (_id: string) => ({ success: false, error }));
		vscode.window.showInputBox = jest.fn().mockResolvedValue("https://newspa.com");
		item = { objectId: mockAppObjectId, contextValue: "SPA-REDIRECT-URI", label: "https://spa.com" };

		// Act
		await redirectUriService.edit(item);

		// Assert
		expect(graphSpy).toHaveBeenCalled();
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
		const graphSpy = jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementation(async (_id: string) => ({ success: false, error }));
		item = { objectId: mockAppObjectId, contextValue: "NATIVE-REDIRECT-URI", label: "https://mobile.com" };
		vscode.window.showInputBox = jest.fn().mockResolvedValue("https://newmobile.com");

		// Act
		await redirectUriService.edit(item);

		// Assert
		expect(graphSpy).toHaveBeenCalled();
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

	test("Add web redirect with successfully with localhost url", async () => {
		// Arrange
		const newRedirectUri: string = "http://localhost/callback";
		item = { objectId: mockAppObjectId, contextValue: "WEB-REDIRECT" };
		const validationSpy = jest.spyOn(validation, "validateRedirectUri");
		jest.spyOn(redirectUriService, "inputRedirectUri").mockImplementation(async (item: AppRegItem, existingRedirectUris: string[], validation: (uri: string, context: string, existingRedirectUris: string[], isEditing: boolean, oldValue: string | undefined) => string | undefined) => {
			const result = validation(newRedirectUri, item.contextValue!, existingRedirectUris, false, undefined);
			return result === undefined ? newRedirectUri : undefined;
		});

		// Act
		await redirectUriService.add(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "WEB-REDIRECT");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(validationSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children?.length).toEqual(2);
		expect(treeItem?.children?.[1].label).toEqual(newRedirectUri);
	});

	test("Edit web redirect with successfully with localhost url", async () => {
		// Arrange
		const newRedirectUri: string = "http://localhost/callback";
		item = { objectId: mockAppObjectId, contextValue: "WEB-REDIRECT", label: mockApplications[0].web.redirectUris[0] };
		const validationSpy = jest.spyOn(validation, "validateRedirectUri");
		jest.spyOn(redirectUriService, "inputRedirectUri").mockImplementation(async (item: AppRegItem, existingRedirectUris: string[], validation: (uri: string, context: string, existingRedirectUris: string[], isEditing: boolean, oldValue: string | undefined) => string | undefined) => {
			const result = validation(newRedirectUri, item.contextValue!, existingRedirectUris, true, item.label!.toString());
			return result === undefined ? newRedirectUri : undefined;
		});

		// Act
		await redirectUriService.edit(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "WEB-REDIRECT");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(validationSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children?.length).toEqual(1);
		expect(treeItem?.children?.[0].label).toEqual(newRedirectUri);
	});

	test("Add spa redirect with successfully with localhost url", async () => {
		// Arrange
		const newRedirectUri: string = "http://localhost/callback";
		item = { objectId: mockAppObjectId, contextValue: "SPA-REDIRECT" };
		const validationSpy = jest.spyOn(validation, "validateRedirectUri");
		jest.spyOn(redirectUriService, "inputRedirectUri").mockImplementation(async (item: AppRegItem, existingRedirectUris: string[], validation: (uri: string, context: string, existingRedirectUris: string[], isEditing: boolean, oldValue: string | undefined) => string | undefined) => {
			const result = validation(newRedirectUri, item.contextValue!, existingRedirectUris, false, undefined);
			return result === undefined ? newRedirectUri : undefined;
		});

		// Act
		await redirectUriService.add(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "SPA-REDIRECT");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(validationSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children?.length).toEqual(2);
		expect(treeItem?.children?.[1].label).toEqual(newRedirectUri);
	});

	test("Add web redirect with duplicate validation error", async () => {
		// Arrange
		const newRedirectUri: string = "https://sample.com/callback";
		item = { objectId: mockAppObjectId, contextValue: "WEB-REDIRECT-URI" };
		const validationSpy = jest.spyOn(validation, "validateRedirectUri");
		jest.spyOn(redirectUriService, "inputRedirectUri").mockImplementation(async (item: AppRegItem, existingRedirectUris: string[], validation: (uri: string, context: string, existingRedirectUris: string[], isEditing: boolean, oldValue: string | undefined) => string | undefined) => {
			const result = validation(newRedirectUri, item.contextValue!, existingRedirectUris, false, undefined);
			return result === undefined ? newRedirectUri : undefined;
		});

		// Act
		await redirectUriService.add(item);

		// Assert
		expect(redirectUriService.inputRedirectUri).toHaveBeenCalled();
		expect(validationSpy).toHaveReturnedWith("The Redirect URI specified already exists in this application.");
		expect(statusBarSpy).toHaveBeenCalled();
	});

	test("Add web redirect with length validation error", async () => {
		// Arrange
		const newRedirectUri: string = "https://sample.com/callback".padEnd(257, "X");
		item = { objectId: mockAppObjectId, contextValue: "WEB-REDIRECT-URI" };
		const validationSpy = jest.spyOn(validation, "validateRedirectUri");
		jest.spyOn(redirectUriService, "inputRedirectUri").mockImplementation(async (item: AppRegItem, existingRedirectUris: string[], validation: (uri: string, context: string, existingRedirectUris: string[], isEditing: boolean, oldValue: string | undefined) => string | undefined) => {
			const result = validation(newRedirectUri, item.contextValue!, existingRedirectUris, false, undefined);
			return result === undefined ? newRedirectUri : undefined;
		});

		// Act
		await redirectUriService.add(item);

		// Assert
		expect(redirectUriService.inputRedirectUri).toHaveBeenCalled();
		expect(validationSpy).toHaveReturnedWith("The Redirect URI is not valid. A Redirect URI cannot be longer than 256 characters.");
		expect(statusBarSpy).toHaveBeenCalled();
	});

	test("Add web redirect with incorrect scheme validation error", async () => {
		// Arrange
		const newRedirectUri: string = "http://sample.com/callback";
		item = { objectId: mockAppObjectId, contextValue: "WEB-REDIRECT-URI" };
		const validationSpy = jest.spyOn(validation, "validateRedirectUri");
		jest.spyOn(redirectUriService, "inputRedirectUri").mockImplementation(async (item: AppRegItem, existingRedirectUris: string[], validation: (uri: string, context: string, existingRedirectUris: string[], isEditing: boolean, oldValue: string | undefined) => string | undefined) => {
			const result = validation(newRedirectUri, item.contextValue!, existingRedirectUris, false, undefined);
			return result === undefined ? newRedirectUri : undefined;
		});

		// Act
		await redirectUriService.add(item);

		// Assert
		expect(redirectUriService.inputRedirectUri).toHaveBeenCalled();
		expect(validationSpy).toHaveReturnedWith("The Redirect URI is not valid. A Redirect URI must start with https:// unless it is using http://localhost.");
		expect(statusBarSpy).toHaveBeenCalled();
	});

	test("Add native redirect with incorrect scheme validation error", async () => {
		// Arrange
		const newRedirectUri: string = "not/a/url";
		item = { objectId: mockAppObjectId, contextValue: "NATIVE-REDIRECT" };
		const validationSpy = jest.spyOn(validation, "validateRedirectUri");
		jest.spyOn(redirectUriService, "inputRedirectUri").mockImplementation(async (item: AppRegItem, existingRedirectUris: string[], validation: (uri: string, context: string, existingRedirectUris: string[], isEditing: boolean, oldValue: string | undefined) => string | undefined) => {
			const result = validation(newRedirectUri, item.contextValue!, existingRedirectUris, false, undefined);
			return result === undefined ? newRedirectUri : undefined;
		});

		// Act
		await redirectUriService.add(item);

		// Assert
		expect(redirectUriService.inputRedirectUri).toHaveBeenCalled();
		expect(validationSpy).toHaveReturnedWith("The Redirect URI is not valid. A Redirect URI must start with https, http, or customScheme://.");
		expect(statusBarSpy).toHaveBeenCalled();
	});

	test("Add web redirect for AAD audiences with total error", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "WEB-REDIRECT" };
		const warningSpy = jest.spyOn(vscode.window, "showWarningMessage");
		const graphSpy = jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementation(async (_id: string, _select: string, _expandOwners?: boolean | undefined) => ({ success: true, value: mockRedirectUris }));
		jest.spyOn(graphApiRepository, "getSignInAudience").mockImplementation(async (_id: string) => ({ success: true, value: "AzureADMyOrg" }));

		// Act
		await redirectUriService.add(item);

		// Assert
		expect(graphSpy).toHaveBeenCalled();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(warningSpy).toHaveBeenCalledWith("You cannot add any more Redirect URIs. The maximum for this application type is 256.", "OK");
	});

	test("Add web redirect for AAD and consumer audiences with total error", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "WEB-REDIRECT" };
		const warningSpy = jest.spyOn(vscode.window, "showWarningMessage");
		const graphSpy = jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementation(async (_id: string, _select: string, _expandOwners?: boolean | undefined) => ({ success: true, value: mockRedirectUris }));
		jest.spyOn(graphApiRepository, "getSignInAudience").mockImplementation(async (_id: string) => ({ success: true, value: "AzureADandPersonalMicrosoftAccount" }));

		// Act
		await redirectUriService.add(item);

		// Assert
		expect(graphSpy).toHaveBeenCalled();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(warningSpy).toHaveBeenCalledWith("You cannot add any more Redirect URIs. The maximum for this application type is 100.", "OK");
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
		const graphSpy = jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementation(async (_id: string) => ({ success: false, error }));
		item = { objectId: mockAppObjectId, contextValue: "WEB-REDIRECT" };

		// Act
		await redirectUriService.add(item);

		// Assert
		expect(graphSpy).toHaveBeenCalled();
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
