import { window } from 'vscode';
import { AppRegDataProvider } from '../data/applicationRegistration';
import { AppRegItem } from '../models/appRegItem';
import { PermissionScope, ApiApplication } from "@microsoft/microsoft-graph-types";
import { v4 as uuidv4 } from 'uuid';
import { ServiceBase } from './serviceBase';
import { GraphClient } from '../clients/graph';

export class OAuth2PermissionScopeService extends ServiceBase {

    // The constructor for the OAuth2PermissionScopeService class.
    constructor(dataProvider: AppRegDataProvider, graphClient: GraphClient) {
        super(dataProvider, graphClient);
    }

    // Adds a new exposed api scope to an application registration.
    public async add(item: AppRegItem): Promise<void> {

        // Capture the new scope details by passing in an empty scope.
        const scope = await this.inputScopeDetails({}, item.objectId!, false);

        // If the user cancels the input then return undefined.
        if (scope === undefined) {
            return;
        }

        // Set the added trigger to the status bar message.
        const previousIcon = item.iconPath;
        const status = this.indicateChange("Adding new scope...", item);

        // Get the existing scopes
        const api = await this.getScopes(item.objectId!);

        // Add the new scope to the existing scopes.
        api.oauth2PermissionScopes!.push(scope);

        // Update the application.
        this._graphClient.updateApplication(item.objectId!, { api: api })
            .then(() => {
                this._onComplete.fire({ success: true, statusBarHandle: status });
            })
            .catch((error) => {
                this._onError.fire({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon });
            });
    }

    // Edits an exposed api scope from an application registration.
    public async edit(item: AppRegItem): Promise<void> {

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
        const status = this.indicateChange("Updating scope...", item);

        // Update the application.
        this._graphClient.updateApplication(item.objectId!, { api: api })
            .then(() => {
                this._onComplete.fire({ success: true, statusBarHandle: status });
            })
            .catch((error) => {
                this._onError.fire({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon });
            });
    }

    // Changes the enabled state of an exposed api scope for an application registration.
    public async changeState(item: AppRegItem): Promise<void> {

        // Set the added trigger to the status bar message.
        const previousIcon = item.iconPath;
        const status = this.indicateChange("Updating scope state...", item);

        // Get the parent application so we can read the scopes.
        const api = await this.getScopes(item.objectId!);

        // Toggle the state of the app role.
        api!.oauth2PermissionScopes!.filter(r => r.id === item.value!)[0].isEnabled = !item.state;

        // Update the application.
        this._graphClient.updateApplication(item.objectId!, { api: api })
            .then(() => {
                this._onComplete.fire({ success: true, statusBarHandle: status });
            })
            .catch((error) => {
                this._onError.fire({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon });
            });
    }

    // Deletes an exposed api scope from an application registration.
    public async delete(item: AppRegItem): Promise<void> {

        if (item.state !== false) {
            window.showWarningMessage("Scopes cannot be deleted unless disabled first.");
            return;
        }

        // Prompt the user to confirm the removal.
        const answer = await window.showInformationMessage(`Do you want to delete the scope ${item.label}?`, "Yes", "No");

        // If the user confirms the removal then delete the role.
        if (answer === "Yes") {
            // Set the added trigger to the status bar message.
            const previousIcon = item.iconPath;
            const status = this.indicateChange("Deleting scope...", item);

            // Get the parent application so we can read the app roles.
            const api = await this.getScopes(item.objectId!);

            // Remove the app role from the array.
            api!.oauth2PermissionScopes!.splice(api!.oauth2PermissionScopes!.findIndex(r => r.id === item.value!), 1);

            // Update the application.
            this._graphClient.updateApplication(item.objectId!, { api: api })
                .then(() => {
                    this._onComplete.fire({ success: true, statusBarHandle: status });
                })
                .catch((error) => {
                    this._onError.fire({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon });
                });
        }
    }

    // Gets the exposed api scopes for an application registration.
    private async getScopes(id: string): Promise<ApiApplication> {
        return (await this._dataProvider.getApplicationPartial(id, "api")).api!;
    }

    // Captures the details for a scope.
    private async inputScopeDetails(scope: PermissionScope, id: string, isEditing: boolean): Promise<PermissionScope | undefined> {

        // Prompt the user for the new value.
        const value = await window.showInputBox({
            prompt: "Edit scope name",
            placeHolder: "Enter a new scope name (e.g. Files.Read)",
            ignoreFocusOut: true,
            value: scope.value ?? undefined,
            validateInput: async (value) => this.validateValue(value, id, isEditing, scope.value ?? undefined)
        });

        // If escape is pressed or the new name is empty then return undefined.
        if (value === undefined) {
            return undefined;
        }

        // Prompt the user for the new allowed member types.
        const consentType = await window.showQuickPick(
            [
                "Administrators only",
                "Administrators and users"
            ],
            {
                placeHolder: "Select who can consent to the scope",
                ignoreFocusOut: true
            });

        // If escape is pressed or the new allowed member types is empty then return undefined.
        if (consentType === undefined) {
            return undefined;
        }

        // Prompt the user for the new admin consent display name.
        const adminConsentDisplayName = await window.showInputBox({
            prompt: "Edit admin consent display name",
            placeHolder: "Enter a new admin consent display name (e.g. Read files)",
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
            prompt: "Edit admin consent description",
            placeHolder: "Enter a new admin consent description (e.g. Allows the app to read files on your behalf.)",
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
            prompt: "Edit user consent display name",
            placeHolder: "Enter a new user consent display name (e.g. Read your files)",
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
            prompt: "Edit user consent description",
            placeHolder: "Enter a new admin consent description (e.g. Allows the app to read your files.)",
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
                "Enabled",
                "Disabled"
            ],
            {
                placeHolder: "Select scope state",
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
        scope.isEnabled = state === "Enabled" ? true : false;
        scope.type = consentType === "Administrators only" ? "Admin" : "User";
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