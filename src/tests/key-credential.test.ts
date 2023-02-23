import * as vscode from "vscode";
import * as fs from "fs";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { KeyCredentialService } from "../services/key-credential";
import { mockApplications, mockAppObjectId, mockPemCertificate, seedMockData } from "./data/test-data";
import { getTopLevelTreeItem } from "./test-utils";
import { TextDocument, Uri } from "vscode";

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

		// Define a standard mock implementation for the dialog functions
		vscode.window.showWarningMessage = jest.fn().mockResolvedValue("Yes");
		vscode.window.showInputBox = jest.fn().mockResolvedValue("New Certificate");

		// Define spies on the functions to be tested
		statusBarSpy = jest.spyOn(vscode.window, "setStatusBarMessage");
		iconSpy = jest.spyOn(vscode, "ThemeIcon");
		triggerCompleteSpy = jest.spyOn(Object.getPrototypeOf(keyCredentialService), "triggerRefresh");
		triggerErrorSpy = jest.spyOn(Object.getPrototypeOf(keyCredentialService), "handleError");
		triggerTreeErrorSpy = jest.spyOn(Object.getPrototypeOf(treeDataProvider), "handleError");

		// The item to be tested
		item = { objectId: mockAppObjectId, contextValue: "CERTIFICATE", keyId: "865b1dbb-5c6d-4e9c-81a3-474af6556bc3" };
	});

	afterAll(() => {
		// Dispose of the key credential service
		keyCredentialService.dispose();
	});

	test("Create class instance", () => {
		// Assert class has been instantiated
		expect(keyCredentialService).toBeDefined();
	});

	test("Delete key credential successfully", async () => {
		// Act
		await keyCredentialService.delete(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "CERTIFICATE-CREDENTIALS");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children?.length).toEqual(1);
	});

	test("Delete key credential with get key error", async () => {
		// Arrange
		const error = new Error("Delete key credential with get key error");
		const graphSpy = jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementation(async () => ({ success: false, error }));

		// Act
		await keyCredentialService.delete(item);

		// Assert
		expect(graphSpy).toHaveBeenCalled();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
	});

	test("Delete key credential with update key error", async () => {
		// Arrange
		const error = new Error("Delete key credential with update key error");
		jest.spyOn(graphApiRepository, "updateKeyCredentials").mockImplementation(async () => ({ success: false, error }));

		// Act
		await keyCredentialService.delete(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
	});

	test("Upload pem certificate and add key credential successfully", async () => {
		// Arrange
		jest.spyOn(vscode.window, "showOpenDialog").mockResolvedValue([Uri.file("testfilepath.pem")]);
		jest.spyOn(vscode.workspace, "openTextDocument").mockResolvedValue({ getText: () => mockPemCertificate } as TextDocument);

		// Act
		await keyCredentialService.upload(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "CERTIFICATE-CREDENTIALS");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children?.length).toEqual(3);
		const certificateItem = treeItem?.children?.[2];
		expect(certificateItem).toBeDefined();
		expect(certificateItem?.label).toEqual("New Certificate");
		expect(certificateItem?.children?.length).toEqual(4);
	});

	test("Upload crt certificate and add key credential successfully", async () => {
		// Arrange
		const crtFile = "./src/tests/data/test.crt";
		jest.spyOn(vscode.window, "showOpenDialog").mockResolvedValue([Uri.file(crtFile)]);
		const mockFile = [{ fsPath: crtFile }];
		const realReadFile = fs.promises.readFile;
		const spyReadFile = jest.spyOn(fs.promises, "readFile");
		spyReadFile.mockImplementation((path) => {
			if (path === mockFile[0].fsPath) {
				return realReadFile(path);
			} else {
				return Promise.reject(new Error("File not found"));
			}
		});

		// Act
		await keyCredentialService.upload(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "CERTIFICATE-CREDENTIALS");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children?.length).toEqual(3);
		const certificateItem = treeItem?.children?.[2];
		expect(certificateItem).toBeDefined();
		expect(certificateItem?.label).toEqual("New Certificate");
		expect(certificateItem?.children?.length).toEqual(4);
	});

	test("Upload cer certificate and add key credential successfully", async () => {
		// Arrange
		const crtFile = "./src/tests/data/test.cer";
		jest.spyOn(vscode.window, "showOpenDialog").mockResolvedValue([Uri.file(crtFile)]);
		const mockFile = [{ fsPath: crtFile }];
		const realReadFile = fs.promises.readFile;
		const spyReadFile = jest.spyOn(fs.promises, "readFile");
		spyReadFile.mockImplementation((path) => {
			if (path === mockFile[0].fsPath) {
				return realReadFile(path);
			} else {
				return Promise.reject(new Error("File not found"));
			}
		});

		// Act
		await keyCredentialService.upload(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "CERTIFICATE-CREDENTIALS");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children?.length).toEqual(3);
		const certificateItem = treeItem?.children?.[2];
		expect(certificateItem).toBeDefined();
		expect(certificateItem?.label).toEqual("New Certificate");
		expect(certificateItem?.children?.length).toEqual(4);
	});

	test("Upload pem certificate and add key credential successfully with empty description", async () => {
		// Arrange
		vscode.window.showInputBox = jest.fn().mockResolvedValue("");
		jest.spyOn(vscode.window, "showOpenDialog").mockResolvedValue([Uri.file("testfilepath.pem")]);
		jest.spyOn(vscode.workspace, "openTextDocument").mockResolvedValue({ getText: () => mockPemCertificate } as TextDocument);

		// Act
		await keyCredentialService.upload(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "CERTIFICATE-CREDENTIALS");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children?.length).toEqual(3);
		const certificateItem = treeItem?.children?.[2];
		expect(certificateItem).toBeDefined();
		expect(certificateItem?.label).toEqual("CN=test.com");
		expect(certificateItem?.children?.length).toEqual(4);
	});

	test("Upload certificate with no existing credentials found error", async () => {
		// Arrange
		jest.spyOn(vscode.window, "showOpenDialog").mockResolvedValue([Uri.file("testfilepath.pem")]);
		jest.spyOn(vscode.workspace, "openTextDocument").mockResolvedValue({ getText: () => mockPemCertificate } as TextDocument);
		const getDetailSpy = jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementation(async () => ({ success: false, error: new Error("Test Error") }));

		// Act
		await keyCredentialService.upload(item);

		// Assert
		expect(getDetailSpy).toHaveBeenCalled();
	});

	test("Upload certificate with no file error", async () => {
		// Arrange
		const fileSpy = jest.spyOn(vscode.window, "showOpenDialog").mockResolvedValue(undefined);
		const inputSpy = jest.spyOn(vscode.window, "showInputBox");

		// Act
		await keyCredentialService.upload(item);

		// Assert
		expect(inputSpy).toHaveBeenCalled();
		expect(fileSpy).toHaveBeenCalled();
	});

	test("Upload certificate with no description error", async () => {
		// Arrange
		const inputSpy = jest.spyOn(vscode.window, "showInputBox").mockResolvedValue(undefined);

		// Act
		await keyCredentialService.upload(item);

		// Assert
		expect(inputSpy).toHaveBeenCalled();
	});

	test("Error getting certificate credential children", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "CERTIFICATE-CREDENTIALS" };
		const error = new Error("Error getting certificate credential children");
		jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementation(async (id: string, select: string) => {
			if (select === "keyCredentials") {
				return { success: false, error };
			}
			return mockApplications.find((app) => app.id === id);
		});

		// Act
		await treeDataProvider.render();
		await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "CERTIFICATE-CREDENTIALS");

		// Assert
		expect(triggerTreeErrorSpy).toHaveBeenCalledWith(error);
	});
});
