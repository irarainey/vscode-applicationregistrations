import * as vscode from "vscode";
import * as validation from "../utils/validation";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { AppRoleService } from "../services/app-role";
import { mockApplications, mockAppObjectId, mockAppRoleId, seedMockData } from "./data/test-data";
import { getTopLevelTreeItem } from "./test-utils";
import { AppRole } from "@microsoft/microsoft-graph-types";

// Create Jest mocks
jest.mock("vscode");
jest.mock("../repositories/graph-api-repository");

// Create the test suite for app role service
describe("App Role Service Tests", () => {
	// Create instances of objects used in the tests
	const graphApiRepository = new GraphApiRepository();
	const treeDataProvider = new AppRegTreeDataProvider(graphApiRepository);
	const appRoleService = new AppRoleService(graphApiRepository, treeDataProvider);

	// Create spy variables
	let triggerCompleteSpy: jest.SpyInstance<any, unknown[], any>;
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

		// Define spies on the functions to be tested
		statusBarSpy = jest.spyOn(vscode.window, "setStatusBarMessage");
		iconSpy = jest.spyOn(vscode, "ThemeIcon");
		triggerCompleteSpy = jest.spyOn(Object.getPrototypeOf(appRoleService), "triggerRefresh");
		triggerTreeErrorSpy = jest.spyOn(Object.getPrototypeOf(treeDataProvider), "handleError");

		// Define a standard mock implementation for the showWarningMessage function
		vscode.window.showWarningMessage = jest.fn().mockResolvedValue("Yes");

		// Define the base item to be used in the tests
		item = { objectId: mockAppObjectId, contextValue: "ROLE-ENABLED", value: mockAppRoleId, state: true };
	});

	afterAll(() => {
		// Dispose of the application service
		appRoleService.dispose();
	});

	test("Create class instance", () => {
		// Assert class has been instantiated
		expect(appRoleService).toBeDefined();
	});

	test("Delete enabled app role but do not disable first", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "ROLE-ENABLED", value: mockAppRoleId, state: true, label: "Test Role" };
		const warningSpy = jest.spyOn(vscode.window, "showWarningMessage").mockResolvedValue("No" as any);

		// Act
		await appRoleService.delete(item, false);

		// Assert
		expect(warningSpy).toHaveBeenCalledWith(`The App Role ${item.label} cannot be deleted unless it is disabled. Do you want to disable the role and then delete it?`, "Yes", "No");
	});

	test("Delete enabled app role and disable first", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "ROLE-ENABLED", value: mockAppRoleId, state: true, label: "Test Role" };
		const changeStateSpy = jest.spyOn(appRoleService, "changeState");
		const warningSpy = jest.spyOn(vscode.window, "showWarningMessage").mockResolvedValue("Yes" as any);

		// Act
		await appRoleService.delete(item, false);

		// Assert
		expect(warningSpy).toHaveBeenCalledWith(`The App Role ${item.label} cannot be deleted unless it is disabled. Do you want to disable the role and then delete it?`, "Yes", "No");
		expect(changeStateSpy).toHaveBeenCalledWith(item, false);
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "APP-ROLES");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children?.length).toEqual(1);
	});

	test("Delete disabled app role", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "ROLE-ENABLED", value: mockAppRoleId, state: false, label: "Test Role" };
		const warningSpy = jest.spyOn(vscode.window, "showWarningMessage").mockResolvedValue("Yes" as any);

		// Act
		await appRoleService.delete(item, false);

		// Assert
		expect(warningSpy).toHaveBeenCalledWith(`Do you want to delete the App Role ${item.label}?`, "Yes", "No");
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "APP-ROLES");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(triggerCompleteSpy).toHaveBeenCalled();
		expect(treeItem?.children?.length).toEqual(1);
	});

	test("Delete disabled app role but decline warning", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "ROLE-ENABLED", value: mockAppRoleId, state: false, label: "Test Role" };
		const warningSpy = jest.spyOn(vscode.window, "showWarningMessage").mockResolvedValue("No" as any);

		// Act
		await appRoleService.delete(item, false);

		// Assert
		expect(warningSpy).toHaveBeenCalledWith(`Do you want to delete the App Role ${item.label}?`, "Yes", "No");
		await treeDataProvider.render();
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "APP-ROLES");
		expect(treeItem?.children?.length).toEqual(2);
	});

	test("Delete disabled app role with no existing roles returned", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "ROLE-ENABLED", value: mockAppRoleId, state: false, label: "Test Role" };
		const graphSpy = jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementationOnce(async () => ({ success: false, error: new Error("Test Error") }));
		const warningSpy = jest.spyOn(vscode.window, "showWarningMessage").mockResolvedValue("Yes" as any);

		// Act
		await appRoleService.delete(item, false);

		// Assert
		expect(warningSpy).toHaveBeenCalledWith(`Do you want to delete the App Role ${item.label}?`, "Yes", "No");
		expect(graphSpy).toHaveBeenCalled();
		await treeDataProvider.render();
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "APP-ROLES");
		expect(treeItem?.children?.length).toEqual(2);
	});

	test("Disable enabled app role but no existing roles returned", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "ROLE-ENABLED", value: mockAppRoleId, state: true, label: "Test Role" };
		const graphSpy = jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementationOnce(async () => ({ success: false, error: new Error("Test Error") }));

		// Act
		await appRoleService.changeState(item, false);

		// Assert
		expect(graphSpy).toHaveBeenCalled();
		await treeDataProvider.render();
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "APP-ROLES");
		expect(treeItem?.children?.length).toEqual(2);
	});

	test("Disable enabled app role", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "ROLE-ENABLED", value: mockAppRoleId, state: true, label: "Test Role" };

		// Act
		await appRoleService.changeState(item, false);

		// Assert
		await treeDataProvider.render();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "APP-ROLES");
		expect(treeItem?.children?.[0].state).toEqual(false);
	});

	test("Enable disabled app role", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "ROLE-ENABLED", value: mockAppRoleId, state: false, label: "Test Role" };

		// Act
		await appRoleService.changeState(item, true);

		// Assert
		await treeDataProvider.render();
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "APP-ROLES");
		expect(treeItem?.children?.[0].state).toEqual(true);
	});

	test("Edit value but no roles returned", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "ROLE-ENABLED", value: mockAppRoleId, state: false, label: "Test Role" };
		const graphSpy = jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementationOnce(async () => ({ success: false, error: new Error("Test Error") }));
		jest.spyOn(vscode.window, "showInputBox").mockResolvedValue("New Role" as any);

		// Act
		await appRoleService.editValue(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(graphSpy).toHaveBeenCalled();
	});

	test("Edit display name", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "ROLE-ENABLED", value: mockAppRoleId, state: false, label: "Test Role" };
		jest.spyOn(vscode.window, "showInputBox").mockResolvedValue("New Role" as any);

		// Act
		await appRoleService.editValue(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "APP-ROLES");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(treeItem?.children?.[0].label).toEqual("New Role");
	});

	test("Edit display name but cancel input", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "ROLE-ENABLED", value: mockAppRoleId, state: false, label: "Test Role" };
		jest.spyOn(vscode.window, "showInputBox").mockResolvedValue(undefined);

		// Act
		await appRoleService.editValue(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(vscode.window.showInputBox).toHaveBeenCalled();
	});

	test("Edit role value", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "ROLE-VALUE", value: mockAppRoleId, label: "Value: Test.Role" };
		jest.spyOn(vscode.window, "showInputBox").mockResolvedValue("New.Role" as any);

		// Act
		await appRoleService.editValue(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "APP-ROLES");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(treeItem?.children![0].children![0].label).toEqual("Value: New.Role");
	});

	test("Edit role value but cancel input", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "ROLE-VALUE", value: mockAppRoleId, label: "Value: Test.Role" };
		jest.spyOn(vscode.window, "showInputBox").mockResolvedValue(undefined);

		// Act
		await appRoleService.editValue(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(vscode.window.showInputBox).toHaveBeenCalled();
	});

	test("Edit role description", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "ROLE-DESCRIPTION", value: mockAppRoleId, label: "Description: Old Description" };
		jest.spyOn(vscode.window, "showInputBox").mockResolvedValue("New Description" as any);

		// Act
		await appRoleService.editValue(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "APP-ROLES");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(treeItem?.children![0].children![1].label).toEqual("Description: New Description");
	});

	test("Edit role description but cancel input", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "ROLE-DESCRIPTION", value: mockAppRoleId, state: false, label: "Description: Old Description" };
		jest.spyOn(vscode.window, "showInputBox").mockResolvedValue(undefined);

		// Act
		await appRoleService.editValue(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(vscode.window.showInputBox).toHaveBeenCalled();
	});

	test("Edit role allowed types to users", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "ROLE-ALLOWED", value: mockAppRoleId, label: "Allowed: Applications" };
		jest.spyOn(vscode.window, "showQuickPick").mockResolvedValue({ label: "Users/Groups", value: ["User"] } as any);

		// Act
		await appRoleService.editValue(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "APP-ROLES");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(treeItem?.children![0].children![2].label).toEqual("Allowed: Users/Groups");
	});

	test("Edit role allowed types to users and applications", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "ROLE-ALLOWED", value: mockAppRoleId, label: "Allowed: Applications" };
		jest.spyOn(vscode.window, "showQuickPick").mockResolvedValue({ label: "Users/Groups", value: ["Application", "User"] } as any);

		// Act
		await appRoleService.editValue(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "APP-ROLES");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(treeItem?.children![0].children![2].label).toEqual("Allowed: Applications, Users/Groups");
	});

	test("Edit role allowed types to applications", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "ROLE-ALLOWED", value: mockAppRoleId, label: "Allowed: Users/Groups" };
		jest.spyOn(vscode.window, "showQuickPick").mockResolvedValue({ label: "Users/Groups", value: ["Application"] } as any);

		// Act
		await appRoleService.editValue(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "APP-ROLES");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(treeItem?.children![0].children![2].label).toEqual("Allowed: Applications");
	});

	test("Edit role allowed types but cancel input", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "ROLE-ALLOWED", value: mockAppRoleId, state: false, label: "Allowed: Applications" };
		jest.spyOn(vscode.window, "showQuickPick").mockResolvedValue(undefined);

		// Act
		await appRoleService.editValue(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(vscode.window.showQuickPick).toHaveBeenCalled();
	});

	test("Add new role but fail to get existing roles", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "APP-ROLES", value: mockAppRoleId };
		const graphSpy = jest.spyOn(graphApiRepository, "getApplicationDetailsPartial")
			.mockImplementationOnce(async () => ({ success: false, error: new Error("Test Error") }));

		// Act
		await appRoleService.add(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(graphSpy).toHaveBeenCalled();
	});

	test("Add new role but cancel pressed on first input", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "APP-ROLES", value: mockAppRoleId };
		const inputSpy = jest.spyOn(vscode.window, "showInputBox").mockResolvedValue(undefined);

		// Act
		await appRoleService.add(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(inputSpy).toHaveBeenCalled();
	});

	test("Add new role but cancel pressed on second input", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "APP-ROLES", value: mockAppRoleId };
		const inputSpy = jest.spyOn(vscode.window, "showInputBox")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce("Test Role");

		// Act
		await appRoleService.add(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(inputSpy).toBeCalled();
	});

	test("Add new role but cancel pressed on third input", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "APP-ROLES", value: mockAppRoleId };
		const inputSpy = jest.spyOn(vscode.window, "showInputBox")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce("Test Role")
			.mockResolvedValueOnce("Test.Role");

		// Act
		await appRoleService.add(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(inputSpy).toBeCalled();
	});

	test("Add new role but cancel pressed on fourth input", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "APP-ROLES", value: mockAppRoleId };
		const inputSpy = jest.spyOn(vscode.window, "showInputBox")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce("Test Role")
			.mockResolvedValueOnce("Test.Role")
			.mockResolvedValueOnce("Test Description");
		const quickPickSpy = jest.spyOn(vscode.window, "showQuickPick").mockResolvedValue(undefined);

		// Act
		await appRoleService.add(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(inputSpy).toBeCalled();
		expect(quickPickSpy).toBeCalled();
	});

	test("Add new role but cancel pressed on fifth input", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "APP-ROLES", value: mockAppRoleId };
		const inputSpy = jest.spyOn(vscode.window, "showInputBox")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce("Test Role")
			.mockResolvedValueOnce("Test.Role")
			.mockResolvedValueOnce("Test Description");
		const quickPickSpy = jest
			.spyOn(vscode.window, "showQuickPick")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce({ label: "Users/Groups", value: ["Application", "User"] } as any);

		// Act
		await appRoleService.add(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(inputSpy).toBeCalled();
		expect(quickPickSpy).toBeCalled();
	});

	test("Add new role successfully", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "APP-ROLES", value: mockAppRoleId };
		const inputSpy = jest.spyOn(vscode.window, "showInputBox").mockResolvedValue(undefined);
		const quickPickSpy = jest
			.spyOn(vscode.window, "showQuickPick")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce({ label: "Users/Groups", value: ["Application", "User"] } as any)
			.mockResolvedValueOnce({ label: "Enabled", value: true } as any);
		jest.spyOn(appRoleService, "inputDisplayName").mockImplementation(async (_title: string, _existingValue: string | undefined, validation: (value: string) => string | undefined) => {
			const result = validation("Test Add Role");
			expect(result).toBeUndefined();
			return "Test Add Role";
		});
		jest.spyOn(appRoleService, "inputValue").mockImplementation(async (_title: string, _isEditing: boolean, _roles: AppRole[], _existingValue: string | undefined, validation: (value: string, isEditing: boolean, oldValue: string | undefined, roles: AppRole[]) => string | undefined) => {
			const result = validation("Test.Role", false, undefined, []);
			expect(result).toBeUndefined();
			return "Test.Role";
		});
		jest.spyOn(appRoleService, "inputDescription").mockImplementation(async (_title: string, _existingValue: string | undefined, validation: (value: string) => string | undefined) => {
			const result = validation("Test Description");
			expect(result).toBeUndefined();
			return "Test Description";
		});

		// Act
		await appRoleService.add(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "APP-ROLES");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(inputSpy).toBeCalled();
		expect(quickPickSpy).toBeCalled();
		expect(treeItem?.children!.length).toEqual(3);
		expect(treeItem?.children![2].label).toEqual("Test Add Role");
	});

	test("Add new role with display name too short error", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "APP-ROLES", value: mockAppRoleId };
		const inputSpy = jest.spyOn(vscode.window, "showInputBox").mockResolvedValue("");
		const validationSpy = jest.spyOn(validation, "validateAppRoleDisplayName");
		jest.spyOn(appRoleService, "inputDisplayName").mockImplementation(async (_title: string, _existingValue: string | undefined, validation: (value: string) => string | undefined) => {
			const result = validation("");
			expect(result).toBe("A display name cannot be empty.");
			return undefined;
		});

		// Act
		await appRoleService.add(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(inputSpy).toBeCalled();
		expect(validationSpy).toBeCalled();
	});

	test("Add new role with display name too long error", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "APP-ROLES", value: mockAppRoleId };
		const inputSpy = jest.spyOn(vscode.window, "showInputBox").mockResolvedValue("");
		const validationSpy = jest.spyOn(validation, "validateAppRoleDisplayName");
		jest.spyOn(appRoleService, "inputDisplayName").mockImplementation(async (_title: string, _existingValue: string | undefined, validation: (value: string) => string | undefined) => {
			const result = validation("X".padEnd(101, "X"));
			expect(result).toBe("A display name cannot be longer than 100 characters.");
			return undefined;
		});

		// Act
		await appRoleService.add(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(inputSpy).toBeCalled();
		expect(validationSpy).toBeCalled();
	});

	test("Add new role with description too short error", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "APP-ROLES", value: mockAppRoleId };
		const inputSpy = jest.spyOn(vscode.window, "showInputBox").mockResolvedValue(undefined).mockResolvedValue("Test Role").mockResolvedValue("Test.Role");
		const validationSpy = jest.spyOn(validation, "validateAppRoleDescription");
		jest.spyOn(appRoleService, "inputDescription").mockImplementation(async (_title: string, _existingValue: string | undefined, validation: (value: string) => string | undefined) => {
			const result = validation("");
			expect(result).toBe("A description cannot be empty.");
			return undefined;
		});

		// Act
		await appRoleService.add(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(inputSpy).toBeCalled();
		expect(validationSpy).toBeCalled();
	});

	test("Add new role with value too long error", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "APP-ROLES", value: mockAppRoleId };
		const inputSpy = jest.spyOn(vscode.window, "showInputBox").mockResolvedValue(undefined).mockResolvedValue("Test Role");
		const validationSpy = jest.spyOn(validation, "validateAppRoleValue");
		jest.spyOn(appRoleService, "inputValue").mockImplementation(async (_title: string, _isEditing: boolean, _roles: AppRole[], _existingValue: string | undefined, validation: (value: string, isEditing: boolean, oldValue: string | undefined, roles: AppRole[]) => string | undefined) => {
			const result = validation("X".padEnd(251, "X"), false, undefined, []);
			expect(result).toBe("A role value cannot be longer than 250 characters.");
			return undefined;
		});

		// Act
		await appRoleService.add(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(inputSpy).toBeCalled();
		expect(validationSpy).toBeCalled();
	});

	test("Add new role with value contains error", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "APP-ROLES", value: mockAppRoleId };
		const inputSpy = jest.spyOn(vscode.window, "showInputBox").mockResolvedValue(undefined).mockResolvedValue("Test Role");
		const validationSpy = jest.spyOn(validation, "validateAppRoleValue");
		jest.spyOn(appRoleService, "inputValue").mockImplementation(async (_title: string, _isEditing: boolean, _roles: AppRole[], _existingValue: string | undefined, validation: (value: string, isEditing: boolean, oldValue: string | undefined, roles: AppRole[]) => string | undefined) => {
			const result = validation("", false, undefined, []);
			expect(result).toBe("A role value cannot be empty.");
			return undefined;
		});

		// Act
		await appRoleService.add(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(inputSpy).toBeCalled();
		expect(validationSpy).toBeCalled();
	});

	test("Add new role with value contains space error", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "APP-ROLES", value: mockAppRoleId };
		const inputSpy = jest.spyOn(vscode.window, "showInputBox").mockResolvedValue(undefined).mockResolvedValueOnce("Test Role");
		const validationSpy = jest.spyOn(validation, "validateAppRoleValue");
		jest.spyOn(appRoleService, "inputValue").mockImplementation(async (_title: string, _isEditing: boolean, _roles: AppRole[], _existingValue: string | undefined, validation: (value: string, isEditing: boolean, oldValue: string | undefined, roles: AppRole[]) => string | undefined) => {
			const result = validation("TEST VALUE", false, undefined, []);
			expect(result).toBe("A role value cannot contain spaces.");
			return undefined;
		});

		// Act
		await appRoleService.add(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(inputSpy).toBeCalled();
		expect(validationSpy).toBeCalled();
	});

	test("Add new role with value starts with full stop error", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "APP-ROLES", value: mockAppRoleId };
		const inputSpy = jest.spyOn(vscode.window, "showInputBox").mockResolvedValue(undefined).mockResolvedValueOnce("Test Role");
		const validationSpy = jest.spyOn(validation, "validateAppRoleValue");
		jest.spyOn(appRoleService, "inputValue").mockImplementation(async (_title: string, _isEditing: boolean, _roles: AppRole[], _existingValue: string | undefined, validation: (value: string, isEditing: boolean, oldValue: string | undefined, roles: AppRole[]) => string | undefined) => {
			const result = validation(".TEST.VALUE", false, undefined, []);
			expect(result).toBe("A role value cannot start with a full stop.");
			return undefined;
		});

		// Act
		await appRoleService.add(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(inputSpy).toBeCalled();
		expect(validationSpy).toBeCalled();
	});

	test("Edit role with duplicate value error", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "APP-ROLES", value: mockAppRoleId };
		const inputSpy = jest.spyOn(vscode.window, "showInputBox").mockResolvedValue(undefined).mockResolvedValueOnce("Test Role");
		const validationSpy = jest.spyOn(validation, "validateAppRoleValue");
		jest.spyOn(appRoleService, "inputValue").mockImplementation(async (_title: string, _isEditing: boolean, _roles: AppRole[], _existingValue: string | undefined, validation: (value: string, isEditing: boolean, oldValue: string | undefined, roles: AppRole[]) => string | undefined) => {
			const result = validation("New.Value", true, "Old.Value", [{ value: "New.Value" }]);
			expect(result).toBe("The role value specified already exists.");
			return undefined;
		});

		// Act
		await appRoleService.edit(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(inputSpy).toBeCalled();
		expect(validationSpy).toBeCalled();
	});

	test("Edit role successfully", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "ROLE-ENABLED", value: mockAppRoleId };
		const inputSpy = jest.spyOn(vscode.window, "showInputBox").mockResolvedValue(undefined).mockResolvedValueOnce("New Write Files Role").mockResolvedValueOnce("Test.Role").mockResolvedValueOnce("Test Description");
		const quickPickSpy = jest
			.spyOn(vscode.window, "showQuickPick")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce({ label: "Users/Groups", value: ["Application", "User"] } as any)
			.mockResolvedValueOnce({ label: "Enabled", value: true } as any);

		// Act
		await appRoleService.edit(item);

		// Assert
		const treeItem = await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "APP-ROLES");
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(inputSpy).toBeCalled();
		expect(quickPickSpy).toBeCalled();
		expect(treeItem?.children!.length).toEqual(2);
		expect(treeItem?.children![0].label).toEqual("New Write Files Role");
	});

	test("Edit role with no existing app roles error", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "ROLE-ENABLED", value: mockAppRoleId };
		const graphSpy = jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementationOnce(async () => ({ success: false, error: new Error("Test Error") }));

		// Act
		await appRoleService.edit(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(iconSpy).toHaveBeenCalled();
		expect(graphSpy).toHaveBeenCalled();
	});

	test("Edit role but cancel pressed on first input", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "APP-ROLES", value: mockAppRoleId };
		const inputSpy = jest.spyOn(vscode.window, "showInputBox").mockResolvedValue(undefined);

		// Act
		await appRoleService.edit(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(inputSpy).toHaveBeenCalled();
	});

	test("Edit role but cancel pressed on second input", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "APP-ROLES", value: mockAppRoleId };
		const inputSpy = jest.spyOn(vscode.window, "showInputBox")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce("Test Role");

		// Act
		await appRoleService.edit(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(inputSpy).toBeCalled();
	});

	test("Edit role but cancel pressed on third input", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "APP-ROLES", value: mockAppRoleId };
		const inputSpy = jest.spyOn(vscode.window, "showInputBox")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce("Test Role")
			.mockResolvedValueOnce("Test.Role");

		// Act
		await appRoleService.edit(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(inputSpy).toBeCalled();
	});

	test("Edit role but cancel pressed on fourth input", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "APP-ROLES", value: mockAppRoleId };
		const inputSpy = jest.spyOn(vscode.window, "showInputBox")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce("Test Role")
			.mockResolvedValueOnce("Test.Role")
			.mockResolvedValueOnce("Test Description");
		const quickPickSpy = jest.spyOn(vscode.window, "showQuickPick").mockResolvedValue(undefined);

		// Act
		await appRoleService.edit(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(inputSpy).toBeCalled();
		expect(quickPickSpy).toBeCalled();
	});

	test("Edit role but cancel pressed on fifth input", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "APP-ROLES", value: mockAppRoleId };
		const inputSpy = jest.spyOn(vscode.window, "showInputBox")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce("Test Role")
			.mockResolvedValueOnce("Test.Role")
			.mockResolvedValueOnce("Test Description");
		const quickPickSpy = jest
			.spyOn(vscode.window, "showQuickPick")
			.mockResolvedValue(undefined)
			.mockResolvedValueOnce({ label: "Users/Groups", value: ["Application", "User"] } as any);

		// Act
		await appRoleService.edit(item);

		// Assert
		expect(statusBarSpy).toHaveBeenCalled();
		expect(inputSpy).toBeCalled();
		expect(quickPickSpy).toBeCalled();
	});

	test("Error getting app role children", async () => {
		// Arrange
		item = { objectId: mockAppObjectId, contextValue: "APP-ROLES" };
		const error = new Error("Error getting app role children");
		jest.spyOn(graphApiRepository, "getApplicationDetailsPartial").mockImplementation(async (id: string, select: string) => {
			if (select === "appRoles") {
				return { success: false, error };
			}
			return mockApplications.find((app) => app.id === id);
		});

		// Act
		await treeDataProvider.render();
		await getTopLevelTreeItem(mockAppObjectId, treeDataProvider, "APP-ROLES");

		// Assert
		expect(triggerTreeErrorSpy).toHaveBeenCalledWith(error);
	});
});
