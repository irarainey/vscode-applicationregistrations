import * as vscode from "vscode";
import * as copyUtils from "../utils/copy-value";
import { AppRegItem } from "../models/app-reg-item";

// Create Jest mocks
jest.mock("vscode");

// Create the test suite for app role service
describe("Copy Value Tests", () => {

	// Create variables used in the tests
	let item: AppRegItem;

	beforeEach(() => {
		//Restore the default mock implementations
		jest.restoreAllMocks();

		// Define the base item to be used in the tests
		item = { contextValue: "COPY", value: "test" };
	});

	test("Copy value successfully", () => {
        // Arrange
        const writeSpy = jest.spyOn(vscode.env.clipboard, "writeText");
		item = { contextValue: "COPY", value: "test" };

        // Act
        copyUtils.copyValue(item);

        // Assert
        expect(writeSpy).toBeCalledWith(item.value);
        expect(vscode.env.clipboard.readText()).toEqual(item.value);
    });

	test("Copy child value successfully", () => {
        // Arrange
        const writeSpy = jest.spyOn(vscode.env.clipboard, "writeText");
		item = { contextValue: "PARENT", value: "test", children: [{ value: "child test" }] };

        // Act
        copyUtils.copyValue(item);

        // Assert
        expect(writeSpy).toBeCalledWith(item.children![0].value);
        expect(vscode.env.clipboard.readText()).toEqual(item.children![0].value);
    });
});
