import * as vscode from "vscode";
import * as execShellCmdUtil from "../utils/exec-shell-cmd";
import * as authenticationUtil from "../utils/authentication";
import * as statusBarUtils from "../utils/status-bar";
import * as errorHandlerUtils from "../error-handler";
import { AzureCliAccountProvider } from "../utils/azure-cli-account-provider";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import { mockTenantId } from "./data/test-data";
import { CLI_LOGOUT_CMD } from "../constants";
import { ErrorResult } from "../types/error-result";

// Create Jest mocks
jest.mock("../utils/exec-shell-cmd");

// Create the test suite for app role service
describe("Authentication Tests", () => {
	// Create instances of objects used in the tests
	const graphApiRepository = new GraphApiRepository();
	const accountProvider = new AzureCliAccountProvider();
	const treeDataProvider = new AppRegTreeDataProvider(graphApiRepository);

	// Create spy variables
	let triggerErrorSpy: jest.SpyInstance<Promise<void>, [result: ErrorResult], any>;
	let statusBarSpy: jest.SpyInstance<void, [], any>;

	beforeAll(async () => {
		// Suppress console output
		console.error = jest.fn();
	});

	beforeEach(() => {
		//Restore the default mock implementations
		jest.restoreAllMocks();

		statusBarSpy = jest.spyOn(statusBarUtils, "clearAllStatusBarMessages");
		triggerErrorSpy = jest.spyOn(errorHandlerUtils, "errorHandler");
	});

	afterAll(() => {
		treeDataProvider.dispose();
	});

	test("Authenticate user successfully without tenant", async () => {
		// Assert
		const treeSpy = jest.spyOn(treeDataProvider, "render");
		jest.spyOn(execShellCmdUtil, "execShellCmd").mockImplementation(async (_cmd: string) => mockTenantId);
		jest.spyOn(vscode.window, "showInputBox").mockResolvedValue(undefined);
		jest.spyOn(treeDataProvider.graphRepository, "initialise").mockResolvedValue(true);
		jest.spyOn(statusBarUtils, "setStatusBarMessage").mockImplementation((_message: string) => "1d399f60-5249-4232-80d3-7a3546eeb62b");

		// Act
		await authenticationUtil.signInUser(treeDataProvider, accountProvider);

		// Assert
		expect(treeSpy).toBeCalledWith(undefined, "AUTHENTICATING");
		expect(treeSpy).toBeCalledWith(undefined, "INITIALISING");
		expect(treeSpy).toBeCalledWith(statusBarUtils.setStatusBarMessage("1d399f60-5249-4232-80d3-7a3546eeb62b"));
	});

	test("Authenticate user unsuccessfully without tenant and initialise error", async () => {
		// Assert
		const treeSpy = jest.spyOn(treeDataProvider, "render");
		jest.spyOn(execShellCmdUtil, "execShellCmd").mockImplementation(async (_cmd: string) => mockTenantId);
		jest.spyOn(vscode.window, "showInputBox").mockResolvedValue(undefined);
		jest.spyOn(treeDataProvider.graphRepository, "initialise").mockResolvedValue(false);

		// Act
		await authenticationUtil.signInUser(treeDataProvider, accountProvider);

		// Assert
		expect(treeSpy).toBeCalledWith(undefined, "AUTHENTICATING");
		expect(treeSpy).toBeCalledWith(undefined, "INITIALISING");
		expect(treeSpy).toBeCalledWith(undefined, "SIGN-IN");
	});

	test("Authenticate user successfully with tenant", async () => {
		// Assert
		const treeSpy = jest.spyOn(treeDataProvider, "render");
		jest.spyOn(execShellCmdUtil, "execShellCmd").mockImplementation(async (_cmd: string) => mockTenantId);
		jest.spyOn(vscode.window, "showInputBox").mockResolvedValue(mockTenantId);
		jest.spyOn(treeDataProvider.graphRepository, "initialise").mockResolvedValue(true);
		jest.spyOn(statusBarUtils, "setStatusBarMessage").mockImplementation((_message: string) => "1d399f60-5249-4232-80d3-7a3546eeb62b");

		// Act
		await authenticationUtil.signInUser(treeDataProvider, accountProvider);

		// Assert
		expect(treeSpy).toBeCalledWith(undefined, "AUTHENTICATING");
		expect(treeSpy).toBeCalledWith(statusBarUtils.setStatusBarMessage("1d399f60-5249-4232-80d3-7a3546eeb62b"), "AUTHENTICATED");
	});

	test("Authenticated user with tenant", async () => {
		// Assert
		const treeSpy = jest.spyOn(treeDataProvider, "render");
		jest.spyOn(execShellCmdUtil, "execShellCmd").mockImplementation(async (_cmd: string) => mockTenantId);
		jest.spyOn(vscode.window, "showInputBox").mockResolvedValue(mockTenantId);
		jest.spyOn(statusBarUtils, "setStatusBarMessage").mockImplementation((_message: string) => "1d399f60-5249-4232-80d3-7a3546eeb62b");

		// Act
		await authenticationUtil.signInUser(treeDataProvider, accountProvider);

		// Assert
		expect(treeSpy).toBeCalledWith(undefined, "AUTHENTICATING");
		expect(treeSpy).toBeCalledWith(statusBarUtils.setStatusBarMessage("1d399f60-5249-4232-80d3-7a3546eeb62b"), "AUTHENTICATED");
	});

	test("Error authenticating user", async () => {
		// Assert
		const treeSpy = jest.spyOn(treeDataProvider, "render");
		jest.spyOn(execShellCmdUtil, "execShellCmd").mockImplementation(async (_cmd: string) => {
			throw new Error("Error authenticating user");
		});
		jest.spyOn(vscode.window, "showInputBox").mockResolvedValue(mockTenantId);
		jest.spyOn(statusBarUtils, "setStatusBarMessage").mockImplementation((_message: string) => "1d399f60-5249-4232-80d3-7a3546eeb62b");

		// Act
		await authenticationUtil.signInUser(treeDataProvider, accountProvider);

		// Assert
		expect(treeSpy).toBeCalledWith(undefined, "AUTHENTICATING");
		expect(treeSpy).toBeCalledWith(undefined, "SIGN-IN");
	});

	test("Sign user out successfully", async () => {
		// Assert
		const treeSpy = jest.spyOn(treeDataProvider, "render");
		const execSpy = jest.spyOn(execShellCmdUtil, "execShellCmd").mockImplementation(async (_cmd: string) => "");

		// Act
		await authenticationUtil.signOutUser(treeDataProvider, accountProvider);

		// Assert
		expect(execSpy).toBeCalledWith(CLI_LOGOUT_CMD);
		expect(treeSpy).toBeCalledWith(undefined, "SIGN-IN");
	});

	test("Sign user out with error", async () => {
		// Assert
		const error = new Error("Sign user out with error");
		const execSpy = jest.spyOn(execShellCmdUtil, "execShellCmd").mockImplementation(async (_cmd: string) => {
			throw error;
		});

		// Act
		await authenticationUtil.signOutUser(treeDataProvider, accountProvider);

		// Assert
		expect(execSpy).toBeCalledWith(CLI_LOGOUT_CMD);
		expect(triggerErrorSpy).toHaveBeenCalledWith(error);
		expect(statusBarSpy).toHaveBeenCalled();
	});
});
