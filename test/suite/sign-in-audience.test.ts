import * as vscode from "vscode";
import { GraphApiRepository } from "../../src/repositories/graph-api-repository";
import { AppRegTreeDataProvider } from "../../src/data/app-reg-tree-data-provider";
import { SignInAudienceService } from "../../src/services/sign-in-audience";

jest.mock('vscode');
jest.mock('../../src/repositories/graph-api-repository');
jest.mock('../../src/data/app-reg-tree-data-provider');

describe('Sign In Audience Tests', () => {

    const mockGraphApiRepository = new GraphApiRepository();
    const mockAppRegTreeDataProvider = new AppRegTreeDataProvider(mockGraphApiRepository);
    const signInAudienceService = new SignInAudienceService(mockGraphApiRepository, mockAppRegTreeDataProvider);
  
    // Test to see if class can be created
    test('Create Instance', () => {
        expect(signInAudienceService).toBeDefined();
    });

    // Test to see if sign in audience can be changed
    test('Update Sign In Audience', async () => {

        let item = { id: "1", label: "Test", objectId: "123", contextValue: "AUDIENCE-PARENT", children: [{ id: "2", objectId: "123", contextValue: "AUDIENCE-CHILD", label: "Multiple Tenants" }] };

        const result = await signInAudienceService.edit(item);

        expect(item.label).toBe("Single Tenant");
    });
});
