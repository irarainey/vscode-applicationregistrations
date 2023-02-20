import { validateDebouncedInput } from "../utils/validation";

// Create the test suite for app role service
describe("Validation Tests", () => {

	beforeEach(() => {
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
    
});
