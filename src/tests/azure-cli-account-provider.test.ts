import * as execShellCmdUtil from "../utils/exec-shell-cmd";
import { AzureCliAccountProvider } from "../utils/azure-cli-account-provider";
import { mockTenantId } from "./data/test-data";
import { CLI_ACCOUNT_INFO_CMD, CLI_LOGIN_CMD, CLI_LOGOUT_CMD } from "../constants";
import { AccountInformation } from "../models/account-information";

// Create Jest mocks
jest.mock("../utils/exec-shell-cmd");

// Create the test suite for app role service
describe("Azure CLI Account Provider Tests", () => {
	// Create instances of objects used in the tests
	const accountProvider = new AzureCliAccountProvider();

	beforeAll(async () => {
		// Suppress console output
		console.error = jest.fn();
	});

	beforeEach(() => {
		//Restore the default mock implementations
		jest.restoreAllMocks();
	});

	test("Create class instance", () => {
		// Assert class has been instantiated
		expect(accountProvider).toBeDefined();
	});

	test("Login user successfully with tenant", async () => {
		// Assert
		const execSpy = jest.spyOn(execShellCmdUtil, "execShellCmd").mockImplementation(async (_cmd: string) => mockTenantId);

		// Act
		const result = await accountProvider.loginUser(mockTenantId);

		// Assert
		expect(execSpy).toBeCalledWith(`${CLI_LOGIN_CMD} --tenant ${mockTenantId}`);
		expect(result).toBe(true);
	});

	test("Login user successfully without tenant", async () => {
		// Assert
		const execSpy = jest.spyOn(execShellCmdUtil, "execShellCmd").mockImplementation(async (_cmd: string) => mockTenantId);

		// Act
		const result = await accountProvider.loginUser("");

		// Assert
		expect(execSpy).toBeCalledWith(CLI_LOGIN_CMD);
		expect(result).toBe(true);
	});

	test("Login user unsuccessfully", async () => {
		// Assert
		const execSpy = jest.spyOn(execShellCmdUtil, "execShellCmd").mockImplementation(async (_cmd: string) => {
			throw new Error("Error");
		});

		// Act
		const result = await accountProvider.loginUser("");

		// Assert
		expect(execSpy).toBeCalledWith(CLI_LOGIN_CMD);
		expect(result).toBe(false);
	});

	test("Logout user", async () => {
		// Assert
		const execSpy = jest.spyOn(execShellCmdUtil, "execShellCmd").mockImplementation(async (_cmd: string) => mockTenantId);

		// Act
		await accountProvider.logoutUser();

		// Assert
		expect(execSpy).toBeCalledWith(CLI_LOGOUT_CMD);
	});

	test("Get account information", async () => {
		// Assert
        const mockAccountInfo = new AccountInformation(mockTenantId, mockTenantId, "AzureCloud", "Test Subscription", "18dcff1a-33d9-4809-9f5c-cf3548c6cab6");
		const execSpy = jest.spyOn(execShellCmdUtil, "execShellCmd").mockImplementation(async (_cmd: string) =>	(JSON.stringify({...mockAccountInfo, name: mockAccountInfo.subscriptionName, id: mockAccountInfo.subscriptionId})));

		// Act
		const result = await accountProvider.getAccountInformation();

		// Assert
		expect(execSpy).toBeCalledWith(CLI_ACCOUNT_INFO_CMD);
		expect(JSON.stringify(result)).toBe(JSON.stringify(mockAccountInfo));
	});
});
