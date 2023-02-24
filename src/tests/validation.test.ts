import * as errorHandlerModule from "../error-handler";
import { validateDebouncedInput, validateOwner } from "../utils/validation";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { OwnerList } from "../models/owner-list";
import { mockUserId, seedMockData } from "./data/test-data";

jest.mock("vscode");
jest.mock("../repositories/graph-api-repository");

// Create the test suite for validation tests
describe("Validation Tests", () => {
	const graphApiRepository = new GraphApiRepository();
    const ownerList = new OwnerList();

	beforeAll(() => {
		// Suppress console output
		console.error = jest.fn();
	});

	beforeEach(() => {
		// Reset mock data
		seedMockData();

		//Restore the default mock implementations
		jest.restoreAllMocks();
	});

	test("Validate debounced input successfully", () => {
        // Arrange
        const input = "XXX";

        // Act
        const validationStatus = validateDebouncedInput(input);

        // Assert
        expect(validationStatus).toBeUndefined();
    });

	test("Validate debounced input unsuccessfully", () => {
        // Arrange
        const input = "XX";

        // Act
        const validationStatus = validateDebouncedInput(input);

        // Assert
        expect(validationStatus).toBe("You must enter at least partial name of the API Application to filter the list. A minimum of 3 characters is required.");
    });

	test("Validate owner with no owner provided", async () => {
        // Act
        const validationStatus = await validateOwner(undefined, [], graphApiRepository, ownerList);

        // Assert
        expect(validationStatus).toBeUndefined();
    });

	test("Validate owner with valid email provided", async () => {
        // Arrange
        const user = "first@user.com";

        // Act
        const validationStatus = await validateOwner(user, [], graphApiRepository, ownerList);

        // Assert
        expect(validationStatus).toBeUndefined();
    });

	test("Validate owner with valid username provided", async () => {
        // Arrange
        const user = "First User";

        // Act
        const validationStatus = await validateOwner(user, [], graphApiRepository, ownerList);

        // Assert
        expect(validationStatus).toBeUndefined();
    });

	test("Validate owner with email provided but with graph error", async () => {
        // Arrange
        const user = "first@user.com";
        const error = new Error("Validate owner with email provided but with graph error");
		jest.spyOn(graphApiRepository, "findUsersByEmail").mockImplementationOnce(async () => ({ success: false, error }));
        const errorSpy = jest.spyOn(errorHandlerModule, "errorHandler");

        // Act
        const validationStatus = await validateOwner(user, [], graphApiRepository, ownerList);

        // Assert
        expect(validationStatus).toBeUndefined();
        expect(errorSpy).toBeCalledWith({ error });
    });

	test("Validate owner with username provided but with graph error", async () => {
        // Arrange
        const user = "First User";
        const error = new Error("Validate owner with username provided but with graph error");
		jest.spyOn(graphApiRepository, "findUsersByName").mockImplementationOnce(async () => ({ success: false, error }));
        const errorSpy = jest.spyOn(errorHandlerModule, "errorHandler");

        // Act
        const validationStatus = await validateOwner(user, [], graphApiRepository, ownerList);

        // Assert
        expect(validationStatus).toBeUndefined();
        expect(errorSpy).toBeCalledWith({ error });
    });

	test("Validate owner with invalid email provided", async () => {
        // Arrange
        const user = "fourth@user.com";

        // Act
        const validationStatus = await validateOwner(user, [], graphApiRepository, ownerList);

        // Assert
        expect(validationStatus).toEqual("No user with an email address beginning with fourth@user.com can be found in your directory.");
    });

	test("Validate owner with invalid username provided", async () => {
        // Arrange
        const user = "Fourth User";

        // Act
        const validationStatus = await validateOwner(user, [], graphApiRepository, ownerList);

        // Assert
        expect(validationStatus).toEqual("No name beginning with Fourth User can be found in your directory.");
    });
    
	test("Validate owner with duplicate username provided", async () => {
        // Arrange
        const user = "Third User";

        // Act
        const validationStatus = await validateOwner(user, [], graphApiRepository, ownerList);

        // Assert
        expect(validationStatus).toEqual("More than one user with the name beginning with Third User exists in your directory.");
    });

	test("Validate owner with existing owner provided", async () => {
        // Arrange
        const user = "First User";

        // Act
        const validationStatus = await validateOwner(user, [{ id: mockUserId }], graphApiRepository, ownerList);

        // Assert
        expect(validationStatus).toEqual("First User is already an owner of this application.");
    });

});
