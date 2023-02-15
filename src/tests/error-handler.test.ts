import * as vscode from "vscode";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import * as errorHandlerModule from "../error-handler";
import { AppRegItem } from "../models/app-reg-item";

// Create Jest mocks
jest.mock("vscode");

// Create the test suite for error handler
describe("Error Handler Tests", () => {
	// Create instances of objects used in the tests
	const graphApiRepository = new GraphApiRepository();
	const treeDataProvider = new AppRegTreeDataProvider(graphApiRepository);

	// Create spies
	let triggerOnDidChangeTreeDataSpy: jest.SpyInstance<void, [item?: AppRegItem | undefined], any>;
	let treeViewRenderSpy: jest.SpyInstance<Promise<void>, [status?: string | undefined, type?: string | undefined], any>;
	let showErrorMessageSpy: jest.SpyInstance<Thenable<vscode.MessageItem | undefined>, [message: string, options: vscode.MessageOptions, ...items: vscode.MessageItem[]], any>;

	beforeAll(() => {
		// Suppress console output
		console.error = jest.fn();
	});

	beforeEach(() => {
		//Restore the default mock implementations
		jest.restoreAllMocks();

		// Define spies on the functions to be tested
		triggerOnDidChangeTreeDataSpy = jest.spyOn(treeDataProvider, "triggerOnDidChangeTreeData");
		showErrorMessageSpy = jest.spyOn(vscode.window, "showErrorMessage");
		treeViewRenderSpy = jest.spyOn(treeDataProvider, "render");
	});

	test("Reset icons on error", async () => {
		// Act
		await errorHandlerModule.errorHandler({ error: new Error("Test Error"), item: { iconPath: "spinner", baseIcon: "base" }, treeDataProvider: treeDataProvider });

		// Assert
		expect(triggerOnDidChangeTreeDataSpy).toHaveBeenCalled();
	});

	test("AZ CLI error", async () => {
		// Act
		await errorHandlerModule.errorHandler({ error: new Error("az login"), item: { iconPath: "spinner", baseIcon: "base" }, treeDataProvider: treeDataProvider });

		// Assert
		const tree = await treeDataProvider.getChildren();
		expect(showErrorMessageSpy).toHaveBeenCalled();
		expect(treeViewRenderSpy).toHaveBeenCalled();
		expect(tree![0].contextValue).toEqual("SIGN-IN");
	});

	test("Azure subscription error", async () => {
		// Act
		await errorHandlerModule.errorHandler({ error: new Error("az account set"), item: { iconPath: "spinner", baseIcon: "base" }, treeDataProvider: treeDataProvider });

		// Assert
		const tree = await treeDataProvider.getChildren();
		expect(showErrorMessageSpy).toHaveBeenCalled();
		expect(treeViewRenderSpy).toHaveBeenCalled();
		expect(tree![0].contextValue).toEqual("SIGN-IN");
	});

	test("Azure subscription error without tree data provider", async () => {
		// Act
		await errorHandlerModule.errorHandler({ error: new Error("az account set"), item: { iconPath: "spinner", baseIcon: "base" }, treeDataProvider: undefined });

		// Assert
		expect(showErrorMessageSpy).toHaveBeenCalled();
	});

	test("Sign in audience error", async () => {
		// Arrange
		jest.spyOn(vscode.window, "showErrorMessage").mockReturnValue({ then: (callback: any) => callback("Open Documentation") });

		// Act
		await errorHandlerModule.errorHandler({ error: new Error("signInAudience"), item: { iconPath: "spinner", baseIcon: "base" }, treeDataProvider: treeDataProvider });

		// Assert
		expect(showErrorMessageSpy).toHaveBeenCalled();
	});
});
