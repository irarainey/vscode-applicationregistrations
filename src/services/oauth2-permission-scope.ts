import { window } from "vscode";
import { AppRegTreeDataProvider } from "../data/app-reg-tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { PermissionScope, ApiApplication, Application, NullableOption } from "@microsoft/microsoft-graph-types";
import { v4 as uuidv4 } from "uuid";
import { ServiceBase } from "./service-base";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { debounce } from "ts-debounce";
import { GraphResult } from "../types/graph-result";

export class OAuth2PermissionScopeService extends ServiceBase {

    // The constructor for the OAuth2PermissionScopeService class.
    constructor(graphRepository: GraphApiRepository, treeDataProvider: AppRegTreeDataProvider) {
        super(graphRepository, treeDataProvider);
    }

    // Adds a new exposed api scope to an application registration.
    async add(item: AppRegItem): Promise<void> {

        // Check to see if the application has an appIdURI. If it doesn't then we don't want to add a scope.
        let appIdUri: string[];

        const result: GraphResult<Application> = await this.graphRepository.getApplicationDetailsPartial<Application>(item.objectId!, "identifierUris");
        if (result.success === true && result.value !== undefined) {
            appIdUri = result.value.identifierUris!;
        } else {
            this.triggerOnError({ success: false, error: result.error, treeDataProvider: this.treeDataProvider });
            return;
        }

        if (appIdUri === undefined || appIdUri.length === 0) {
            window.showWarningMessage("This application does not have an Application II URI. Please add one before adding a scope.", "OK");
            return;
        }

        // Capture the new scope details by passing in an empty scope.
        const scope = await this.inputScopeDetails({}, item.objectId!, false);

        // If the user cancels the input then return undefined.
        if (scope === undefined) {
            return;
        }

        // Set the added trigger to the status bar message.
        const previousIcon = item.iconPath;
        const status = this.triggerTreeChange("Adding Scope", item);

        // Get the existing scopes
        const api = await this.getScopes(item.objectId!);

        // Add the new scope to the existing scopes.
        api.oauth2PermissionScopes!.push(scope);

        // Update the application.
        this.graphRepository.updateApplication(item.objectId!, { api: api })
            .then(() => {
                this.triggerOnComplete({ success: true, statusBarHandle: status });
            })
            .catch((error) => {
                this.triggerOnError({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon, treeDataProvider: this.treeDataProvider });
            });
    }

    // Edits an exposed api scope from an application registration.
    async edit(item: AppRegItem): Promise<void> {

        // Get the parent application so we can read the app roles.
        const api = await this.getScopes(item.objectId!);

        // Capture the new app role details by passing in the existing role.
        const scope = await this.inputScopeDetails(api!.oauth2PermissionScopes!.filter(r => r.id === item.value!)[0], item.objectId!, true);

        // If the user cancels the input then return undefined.
        if (scope === undefined) {
            return;
        }

        // Set the added trigger to the status bar message.
        const previousIcon = item.iconPath;
        const status = this.triggerTreeChange("Updating Scope", item);

        // Update the application.
        this.graphRepository.updateApplication(item.objectId!, { api: api })
            .then(() => {
                this.triggerOnComplete({ success: true, statusBarHandle: status });
            })
            .catch((error) => {
                this.triggerOnError({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon, treeDataProvider: this.treeDataProvider });
            });
    }

    // Changes the enabled state of an exposed api scope for an application registration.
    async changeState(item: AppRegItem, state: boolean): Promise<void> {

        // Set the added trigger to the status bar message.
        const previousIcon = item.iconPath;
        const status = this.triggerTreeChange(state === true ? "Enabling Scope" : "Disabling Scope", item);

        // Get the parent application so we can read the scopes.
        const api = await this.getScopes(item.objectId!);

        // Toggle the state of the app role.
        api!.oauth2PermissionScopes!.filter(r => r.id === item.value!)[0].isEnabled = state;

        // Update the application.
        this.graphRepository.updateApplication(item.objectId!, { api: api })
            .then(() => {
                this.triggerOnComplete({ success: true, statusBarHandle: status });
            })
            .catch((error) => {
                this.triggerOnError({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon, treeDataProvider: this.treeDataProvider });
            });
    }

    // Deletes an exposed api scope from an application registration.
    async delete(item: AppRegItem): Promise<void> {

        if (item.state !== false) {
            window.showWarningMessage("Scopes cannot be deleted unless disabled first.", "OK");
            return;
        }

        // Prompt the user to confirm the removal.
        const answer = await window.showInformationMessage(`Do you want to delete the Scope ${item.label}?`, "Yes", "No");

        // If the user confirms the removal then delete the role.
        if (answer === "Yes") {
            // Set the added trigger to the status bar message.
            const previousIcon = item.iconPath;
            const status = this.triggerTreeChange("Deleting Scope", item);

            // Get the parent application so we can read the app roles.
            const api = await this.getScopes(item.objectId!);

            // Remove the app role from the array.
            api!.oauth2PermissionScopes!.splice(api!.oauth2PermissionScopes!.findIndex(r => r.id === item.value!), 1);

            // Update the application.
            this.graphRepository.updateApplication(item.objectId!, { api: api })
                .then(() => {
                    this.triggerOnComplete({ success: true, statusBarHandle: status });
                })
                .catch((error) => {
                    this.triggerOnError({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon, treeDataProvider: this.treeDataProvider });
                });
        }
    }

    // Gets the exposed api scopes for an application registration.
    private async getScopes(id: string): Promise<ApiApplication> {
        const result: GraphResult<Application> = await this.graphRepository.getApplicationDetailsPartial<Application>(id, "api");
        if (result.success === true && result.value !== undefined) {
            return result.value.api!;
        } else {
            this.triggerOnError({ success: false, error: result.error, treeDataProvider: this.treeDataProvider });
            return {};
        }
    }

    // Captures the details for a scope.
    private async inputScopeDetails(scope: PermissionScope, id: string, isEditing: boolean): Promise<PermissionScope | undefined> {

        // Debounce the validation function to prevent multiple calls to the Graph API.
        const validation = async (value: string, id: string, isEditing: boolean, oldValue: string | undefined) => this.validateValue(value, id, isEditing, scope.value ?? undefined);
        const debouncedValidation = debounce(validation, 500);

        // Prompt the user for the new value.
        const value = await window.showInputBox({
            prompt: "Scope name",
            placeHolder: "Enter a scope name (e.g. Files.Read)",
            title: isEditing === true ? "Edit Exposed API Permission (1/7)" : "Add API Exposed Permission (1/7)",
            ignoreFocusOut: true,
            value: scope.value ?? undefined,
            validateInput: async (value) => debouncedValidation(value, id, isEditing, scope.value ?? undefined)
        });

        // If escape is pressed or the new name is empty then return undefined.
        if (value === undefined) {
            return undefined;
        }

        // Prompt the user for the new allowed member types.
        const consentType = await window.showQuickPick(
            [
                {
                    label: "Administrators only",
                    description: "Only administrators can consent to the scope",
                    value: "Admin"
                },
                {
                    label: "Administrators and users",
                    description: "Administrators and users can consent to the scope",
                    value: "User"
                }
            ],
            {
                placeHolder: "Select who can consent to the scope",
                title: isEditing === true ? "Edit Exposed API Permission (2/7)" : "Add API Exposed Permission (2/7)",
                ignoreFocusOut: true
            });

        // If escape is pressed or the new allowed member types is empty then return undefined.
        if (consentType === undefined) {
            return undefined;
        }

        // Prompt the user for the new admin consent display name.
        const adminConsentDisplayName = await window.showInputBox({
            prompt: "Admin consent display name",
            placeHolder: "Enter an admin consent display name (e.g. Read files)",
            title: isEditing === true ? "Edit Exposed API Permission (3/7)" : "Add API Exposed Permission (3/7)",
            ignoreFocusOut: true,
            value: scope.adminConsentDisplayName ?? undefined,
            validateInput: async (value) => this.validateAdminDisplayName(value)
        });

        // If escape is pressed or the new display name is empty then return undefined.
        if (adminConsentDisplayName === undefined) {
            return undefined;
        }

        // Prompt the user for the new admin consent description.
        const adminConsentDescription = await window.showInputBox({
            prompt: "Admin consent description",
            placeHolder: "Enter an admin consent description (e.g. Allows the app to read files on your behalf.)",
            title: isEditing === true ? "Edit Exposed API Permission (4/7)" : "Add API Exposed Permission (4/7)",
            ignoreFocusOut: true,
            value: scope.adminConsentDescription ?? undefined,
            validateInput: async (value) => this.validateAdminDescription(value)
        });

        // If escape is pressed or the new description is empty then return undefined.
        if (adminConsentDescription === undefined) {
            return undefined;
        }

        // Prompt the user for the new user consent display name.
        const userConsentDisplayName = await window.showInputBox({
            prompt: "User consent display name",
            placeHolder: "Enter an optional user consent display name (e.g. Read your files)",
            title: isEditing === true ? "Edit Exposed API Permission (5/7)" : "Add API Exposed Permission (5/7)",
            ignoreFocusOut: true,
            value: scope.userConsentDisplayName ?? undefined,
            validateInput: async (value) => this.validateUserDisplayName(value)
        });

        // If escape is pressed or the new user consent display name is empty then return undefined.
        if (userConsentDisplayName === undefined) {
            return undefined;
        }

        // Prompt the user for the new user consent description.
        const userConsentDescription = await window.showInputBox({
            prompt: "User consent description",
            placeHolder: "Enter an optional user consent description (e.g. Allows the app to read your files.)",
            title: isEditing === true ? "Edit Exposed API Permission (6/7)" : "Add API Exposed Permission (6/7)",
            ignoreFocusOut: true,
            value: scope.userConsentDescription ?? undefined
        });

        // If escape is pressed or the new user consent description is empty then return undefined.
        if (userConsentDescription === undefined) {
            return undefined;
        }

        // Prompt the user for the new state.
        const state = await window.showQuickPick(
            [
                {
                    label: "Enabled",
                    value: true
                },
                {
                    label: "Disabled",
                    value: false
                }
            ],
            {
                placeHolder: "Select scope state",
                title: isEditing === true ? "Edit Exposed API Permission (7/7)" : "Add API Exposed Permission (7/7)",
                ignoreFocusOut: true
            });

        // If escape is pressed or the new state is empty then return undefined.
        if (state === undefined) {
            return undefined;
        }

        // Set the new values on the scope
        scope.adminConsentDescription = adminConsentDescription;
        scope.adminConsentDisplayName = adminConsentDisplayName;
        scope.id = scope.id ?? uuidv4();
        scope.isEnabled = state.value;
        scope.type = consentType.value;
        scope.userConsentDisplayName = userConsentDisplayName;
        scope.userConsentDescription = userConsentDescription;
        scope.value = value;

        return scope;
    }

    // Validates the admin display name of an scope.
    private validateAdminDisplayName(displayName: string): string | undefined {

        // Check the length of the display name.
        if (displayName.length > 100) {
            return "An admin display name cannot be longer than 100 characters.";
        }

        // Check the length of the display name.
        if (displayName.length < 1) {
            return "An admin display name cannot be empty.";
        }

        return undefined;
    }

    // Validates the user display name of an scope.
    private validateUserDisplayName(displayName: string): string | undefined {

        // Check the length of the display name.
        if (displayName.length > 100) {
            return "An admin display name cannot be longer than 100 characters.";
        }

        return undefined;
    }

    // Validates the value of an scope.
    private async validateValue(value: string, id: string, isEditing: boolean, oldValue: string | undefined): Promise<string | undefined> {

        // Check the length of the value.
        if (value.length > 40) {
            return "A value cannot be longer than 40 characters.";
        }

        // Check the length of the value.
        if (value.length < 1) {
            return "A value cannot be empty.";
        }

        // Check to see if the value already exists.
        if ((isEditing === true && oldValue !== value) || isEditing === false) {
            const scopes = await this.getScopes(id);
            if (scopes.oauth2PermissionScopes!.find(r => r.value === value) !== undefined) {
                return "The value specified already exists.";
            }
        }

        return undefined;
    }

    // Validates the value of the admin consent description.
    private validateAdminDescription(description: string): string | undefined {

        // Check the length of the admin consent description.
        if (description.length < 1) {
            return "A admin consent description cannot be empty.";
        }

        return undefined;
    }
}