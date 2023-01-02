import { window, ThemeIcon, Disposable } from 'vscode';
import { GraphClient } from '../clients/graph';
import { AppRegDataProvider } from '../data/applicationRegistration';
import { AppRegItem } from '../models/appRegItem';
import { PermissionScope, ApiApplication } from "@microsoft/microsoft-graph-types";
import { v4 as uuidv4 } from 'uuid';

export class OAuth2PermissionScopeService {

    // A private instance of the GraphClient class.
    private _graphClient: GraphClient;

    // A private instance of the AppRegDataProvider class.
    private _dataProvider: AppRegDataProvider;

    // The constructor for the OAuth2PermissionScopeService class.
    constructor(dataProvider: AppRegDataProvider) {
        this._dataProvider = dataProvider;
        this._graphClient = dataProvider.graphClient;
    }

    // Adds a new exposed api scope to an application registration.
    public async add(item: AppRegItem): Promise<Disposable | undefined> {

        // Set the added trigger default to undefined.
        let added = undefined;

        // Prompt the user for the new value.
        const newValue = await window.showInputBox({
            prompt: "Edit scope name",
            placeHolder: "Enter a new scope name (e.g. Files.Read)",
            ignoreFocusOut: true
        });

        // If escape is pressed or the new name is empty then return undefined.
        if (newValue === undefined || newValue === "") {
            return undefined;
        }

        // Prompt the user for the new admin consent display name.
        const newAdminConsentDisplayName = await window.showInputBox({
            prompt: "Edit admin consent display name",
            placeHolder: "Enter a new admin consent display name (e.g. Read files)",
            ignoreFocusOut: true
        });

        // If escape is pressed or the new value is empty then return undefined.
        if (newAdminConsentDisplayName === undefined || newAdminConsentDisplayName === "") {
            return undefined;
        }

        // Prompt the user for the new admin consent description.
        const newAdminConsentDescription = await window.showInputBox({
            prompt: "Edit admin consent description",
            placeHolder: "Enter a new admin consent description (e.g. Allows the app to read files on your behalf.)",
            ignoreFocusOut: true
        });

        // If escape is pressed or the new description is empty then return undefined.
        if (newAdminConsentDescription === undefined || newAdminConsentDescription === "") {
            return undefined;
        }

        // Prompt the user for the new allowed member types.
        const newConsentType = await window.showQuickPick(
            [
                "Administrators only",
                "Administrators and users"
            ],
            {
                placeHolder: "Select who can consent to the scope",
                ignoreFocusOut: true
            });

        // If escape is pressed or the new allowed member types is empty then return undefined.
        if (newConsentType === undefined || newConsentType === "") {
            return undefined;
        }

        // Prompt the user for the new state.
        const newState = await window.showQuickPick(
            [
                "Enabled",
                "Disabled"
            ],
            {
                placeHolder: "Select scope state",
                ignoreFocusOut: true
            });

        // If escape is pressed or the new state is empty then return undefined.
        if (newState === undefined || newState === "") {
            return undefined;
        }

        added = window.setStatusBarMessage("$(loading~spin) Adding new Scope...");
        item.iconPath = new ThemeIcon("loading~spin");
        this._dataProvider.triggerOnDidChangeTreeData();

        const api = await this.getScopes(item.objectId!);

        const newRole: PermissionScope = {
            id: uuidv4(),
            type: newConsentType === "Administrators only" ? "Admin" : "User",
            adminConsentDescription: newAdminConsentDescription,
            adminConsentDisplayName: newAdminConsentDisplayName,
            isEnabled: newState === "Enabled" ? true : false,
            value: newValue
        };

        api.oauth2PermissionScopes!.push(newRole);

        // Update the application.
        await this._graphClient.updateApplication(item.objectId!, { api: api })
            .catch((error) => {
                console.error(error);
            });

        return added;
    }

    // Edits an exposed api scope from an application registration.
    public async edit(item: AppRegItem): Promise<Disposable | undefined> {

        // Set the edited trigger default to undefined.
        let edited = undefined;

        // Get the parent application so we can read the app roles.
        const api = await this.getScopes(item.objectId!);

        // Get the app role to edit.
        const scope = api!.oauth2PermissionScopes!.filter(r => r.id === item.value!)[0];

        // Prompt the user for the new value.
        const newValue = await window.showInputBox({
            prompt: "Edit scope name",
            value: scope.value!,
            ignoreFocusOut: true
        });

        // If escape is pressed or the new name is empty then return undefined.
        if (newValue === undefined || newValue === "") {
            return undefined;
        }

        // Prompt the user for the new admin consent display name.
        const newAdminConsentDisplayName = await window.showInputBox({
            prompt: "Edit admin consent display name",
            value: scope.adminConsentDisplayName!,
            ignoreFocusOut: true
        });

        // If escape is pressed or the new value is empty then return undefined.
        if (newAdminConsentDisplayName === undefined || newAdminConsentDisplayName === "") {
            return undefined;
        }

        // Prompt the user for the new admin consent description.
        const newAdminConsentDescription = await window.showInputBox({
            prompt: "Edit admin consent description",
            value: scope.adminConsentDescription!,
            ignoreFocusOut: true
        });

        // If escape is pressed or the new description is empty then return undefined.
        if (newAdminConsentDescription === undefined || newAdminConsentDescription === "") {
            return undefined;
        }

        // Prompt the user for the new allowed member types.
        const newConsentType = await window.showQuickPick(
            [
                "Administrators only",
                "Administrators and users"
            ],
            {
                placeHolder: "Select who can consent",
                ignoreFocusOut: true
            });

        // If escape is pressed or the new allowed member types is empty then return undefined.
        if (newConsentType === undefined || newConsentType === "") {
            return undefined;
        }

        // Prompt the user for the new state.
        const newState = await window.showQuickPick(
            [
                "Enabled",
                "Disabled"
            ],
            {
                placeHolder: "Select scope state",
                ignoreFocusOut: true
            });

        // If escape is pressed or the new state is empty then return undefined.
        if (newState === undefined || newState === "") {
            return undefined;
        }

        edited = window.setStatusBarMessage("$(loading~spin) Updating Scope...");
        item.iconPath = new ThemeIcon("loading~spin");
        this._dataProvider.triggerOnDidChangeTreeData();

        // Update the app role in place
        this.updateScope(api!.oauth2PermissionScopes!.filter(r => r.id === item.value!)[0], newAdminConsentDisplayName, newValue, newAdminConsentDescription, newConsentType, newState!);

        // Update the application.
        await this._graphClient.updateApplication(item.objectId!, { api: api })
            .catch((error) => {
                console.error(error);
            });

        return edited;
    }

    // Changes the enabled state of an exposed api scope for an application registration.
    public async changeState(item: AppRegItem): Promise<Disposable | undefined> {

        // Set the state changed trigger default to undefined.
        let stateChanged = undefined;

        stateChanged = window.setStatusBarMessage("$(loading~spin) Updating Scope State...");
        item.iconPath = new ThemeIcon("loading~spin");
        this._dataProvider.triggerOnDidChangeTreeData();

        // Get the parent application so we can read the scopes.
        const api = await this.getScopes(item.objectId!);

        // Toggle the state of the app role.
        api!.oauth2PermissionScopes!.filter(r => r.id === item.value!)[0].isEnabled = !item.state;

        // Update the application.
        await this._graphClient.updateApplication(item.objectId!, { api : api })
            .catch((error) => {
                console.error(error);
            });

        return stateChanged;
    }

    // Deletes an exposed api scope from an application registration.
    public async delete(item: AppRegItem): Promise<Disposable | undefined> {

        if (item.state !== false) {
            window.showWarningMessage("Scopes cannot be deleted unless disabled first.");
            return;
        }

        // Set the deleted trigger default to undefined.
        let deleted = undefined;

        // Prompt the user to confirm the removal.
        const answer = await window.showInformationMessage(`Do you want to delete the scope ${item.label}?`, "Yes", "No");

        // If the user confirms the removal then delete the role.
        if (answer === "Yes") {
            deleted = window.setStatusBarMessage("$(loading~spin) Deleting Scope...");
            item.iconPath = new ThemeIcon("loading~spin");
            this._dataProvider.triggerOnDidChangeTreeData();

            // Get the parent application so we can read the app roles.
            const api = await this.getScopes(item.objectId!);

            // Remove the app role from the array.
            api!.oauth2PermissionScopes!.splice(api!.oauth2PermissionScopes!.findIndex(r => r.id === item.value!), 1);

            // Update the application.
            await this._graphClient.updateApplication(item.objectId!, { api: api })
                .catch((error) => {
                    console.error(error);
                });
        }

        return deleted;
    }

    // Gets the exposed api scopes for an application registration.
    private async getScopes(id: string): Promise<ApiApplication> {
        return (await this._dataProvider.getApplicationPartial(id, "api")).api!;
    }

    // Updates an exposed api scope.
    private updateScope(scope: PermissionScope, adminConsentDisplayName: string, value: string, adminConsentDescription: string, type: string, state: string) {
        scope.value = value;
        scope.adminConsentDisplayName = adminConsentDisplayName;
        scope.adminConsentDescription = adminConsentDescription;
        scope.type = type === "Administrators only" ? "Admin" : "User";
        scope.isEnabled = state === "Enabled" ? true : false;
    }    
}