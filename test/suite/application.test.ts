import * as vscode from "vscode";

jest.mock('vscode');

// jest.mock('vscode', () => {
// 	  return {
// 		window: {
// 			showInformationMessage: jest.fn().mockResolvedValue("YES")
// 		}
// 	};
// });

describe('Application Tests', () => {
	const retVal = vscode.window.showInformationMessage('Start all tests.', "YES", "NO");

	test('Sample test', () => {
		expect(retVal).resolves.toBe("YES");
	});
});
