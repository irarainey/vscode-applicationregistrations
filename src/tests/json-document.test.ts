import { JsonDocument } from "../models/json-document";

// Create the test suite for app role service
describe("Json Document Tests", () => {

	beforeEach(() => {
		//Restore the default mock implementations
		jest.restoreAllMocks();
	});

	test("Return value successfully", () => {
        // Arrange
        const jsonDocumentClass = new JsonDocument({ testProperty: "test value" });

        // Act
        const result = jsonDocumentClass.provideTextDocumentContent();

        // Assert
        expect(result).toEqual(JSON.stringify({ testProperty: "test value" }, null, 4));
    });
});
