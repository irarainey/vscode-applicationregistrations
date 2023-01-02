import { window, ThemeIcon, Disposable } from 'vscode';
import { GraphClient } from '../clients/graph';
import { AppRegDataProvider } from '../data/applicationRegistration';
import { AppRegItem } from '../models/appRegItem';
import { AppRole } from "@microsoft/microsoft-graph-types";
import { v4 as uuidv4 } from 'uuid';

export class AppRoleService {

    // A private instance of the GraphClient class.
    private _graphClient: GraphClient;

    // A private instance of the AppRegDataProvider class.
    private _dataProvider: AppRegDataProvider;

    // The constructor for the AppRolesService class.
    constructor(dataProvider: AppRegDataProvider) {
        this._dataProvider = dataProvider;
        this._graphClient = dataProvider.graphClient;
    }

    // Adds a new app role to an application registration.
    public async add(item: AppRegItem): Promise<Disposable | undefined> {

        // Set the added trigger default to undefined.
        let added = undefined;

        // Prompt the user for the new display name.
        const newName = await window.showInputBox({
            prompt: "Edit display name",
            placeHolder: "Enter a display name for the new app role",
            ignoreFocusOut: true
        });

        // If escape is pressed or the new name is empty then return undefined.
        if (newName === undefined || newName === "") {
            return undefined;
        }

        // Prompt the user for the new value.
        const newValue = await window.showInputBox({
            prompt: "Edit value",
            placeHolder: "Enter a value for the new app role",
            ignoreFocusOut: true
        });

        // If escape is pressed or the new value is empty then return undefined.
        if (newValue === undefined || newValue === "") {
            return undefined;
        }

        // Prompt the user for the new display name.
        const newDescription = await window.showInputBox({
            prompt: "Edit description",
            placeHolder: "Enter a description for the new app role",
            ignoreFocusOut: true
        });

        // If escape is pressed or the new description is empty then return undefined.
        if (newDescription === undefined || newDescription === "") {
            return undefined;
        }

        // Prompt the user for the new allowed member types.
        const newAllowedMemberTypes = await window.showQuickPick(
            [
                "Users/Groups",
                "Applications",
                "Both (Users/Groups + Applications)"
            ],
            {
                placeHolder: "Select allowed member types",
                ignoreFocusOut: true
            });

        // If escape is pressed or the new allowed member types is empty then return undefined.
        if (newAllowedMemberTypes === undefined || newAllowedMemberTypes === "") {
            return undefined;
        }

        // Prompt the user for the new state.
        const newState = await window.showQuickPick(
            [
                "Enabled",
                "Disabled"
            ],
            {
                placeHolder: "Select role state",
                ignoreFocusOut: true
            });

        // If escape is pressed or the new state is empty then return undefined.
        if (newState === undefined || newState === "") {
            return undefined;
        }

        added = window.setStatusBarMessage("$(loading~spin) Adding new app role...");
        item.iconPath = new ThemeIcon("loading~spin");
        this._dataProvider.triggerOnDidChangeTreeData();

        const roles = await this.getAppRoles(item.objectId!);

        const newRole: AppRole = {
            id: uuidv4(),
            allowedMemberTypes: newAllowedMemberTypes === "Users/Groups" ? ["User"] : newAllowedMemberTypes === "Applications" ? ["Application"] : ["User", "Application"],
            description: newDescription,
            displayName: newName,
            isEnabled: newState === "Enabled" ? true : false,
            value: newValue
        };

        roles.push(newRole);

        // Update the application.
        await this._graphClient.updateApplication(item.objectId!, { appRoles: roles })
            .catch((error) => {
                console.error(error);
            });

        return added;
    }

    // Edits an app role from an application registration.
    public async edit(item: AppRegItem): Promise<Disposable | undefined> {

        // Set the edited trigger default to undefined.
        let edited = undefined;

        // Get the parent application so we can read the app roles.
        const roles = await this.getAppRoles(item.objectId!);

        // Get the app role to edit.
        const role = roles.filter(r => r.id === item.value!)[0];

        // Prompt the user for the new display name.
        const newName = await window.showInputBox({
            prompt: "Edit display name",
            value: role.displayName!,
            ignoreFocusOut: true
        });

        // If escape is pressed or the new name is empty then return undefined.
        if (newName === undefined || newName === "") {
            return undefined;
        }

        // Prompt the user for the new value.
        const newValue = await window.showInputBox({
            prompt: "Edit value",
            value: role.value!,
            ignoreFocusOut: true
        });

        // If escape is pressed or the new value is empty then return undefined.
        if (newValue === undefined || newValue === "") {
            return undefined;
        }

        // Prompt the user for the new display name.
        const newDescription = await window.showInputBox({
            prompt: "Edit description",
            value: role.description!,
            ignoreFocusOut: true
        });

        // If escape is pressed or the new description is empty then return undefined.
        if (newDescription === undefined || newDescription === "") {
            return undefined;
        }

        // Prompt the user for the new allowed member types.
        const newAllowedMemberTypes = await window.showQuickPick(
            [
                "Users/Groups",
                "Applications",
                "Both (Users/Groups + Applications)"
            ],
            {
                placeHolder: "Select allowed member types",
                ignoreFocusOut: true
            });

        // If escape is pressed or the new allowed member types is empty then return undefined.
        if (newAllowedMemberTypes === undefined || newAllowedMemberTypes === "") {
            return undefined;
        }

        // Prompt the user for the new state.
        const newState = await window.showQuickPick(
            [
                "Enabled",
                "Disabled"
            ],
            {
                placeHolder: "Select role state",
                ignoreFocusOut: true
            });

        // If escape is pressed or the new state is empty then return undefined.
        if (newState === undefined || newState === "") {
            return undefined;
        }

        edited = window.setStatusBarMessage("$(loading~spin) Updating app role...");
        item.iconPath = new ThemeIcon("loading~spin");
        this._dataProvider.triggerOnDidChangeTreeData();

        // Update the app role in place
        this.updateRole(roles.filter(r => r.id === item.value!)[0], newName, newValue, newDescription, newAllowedMemberTypes, newState!);

        // Update the application.
        await this._graphClient.updateApplication(item.objectId!, { appRoles: roles })
            .catch((error) => {
                console.error(error);
            });

        return edited;
    }

    // Changes the enabled state of an app role from an application registration.
    public async changeState(item: AppRegItem): Promise<Disposable | undefined> {

        // Set the state changed trigger default to undefined.
        let stateChanged = undefined;

        stateChanged = window.setStatusBarMessage("$(loading~spin) Updating app role...");
        item.iconPath = new ThemeIcon("loading~spin");
        this._dataProvider.triggerOnDidChangeTreeData();

        // Get the parent application so we can read the app roles.
        const roles = await this.getAppRoles(item.objectId!);

        // Toggle the state of the app role.
        roles.filter(r => r.id === item.value!)[0].isEnabled = !item.state;

        // Update the application.
        await this._graphClient.updateApplication(item.objectId!, { appRoles: roles })
            .catch((error) => {
                console.error(error);
            });

        return stateChanged;
    }

    // Deletes an app role from an application registration.
    public async delete(item: AppRegItem): Promise<Disposable | undefined> {

        if (item.state !== false) {
            window.showWarningMessage("Role cannot be deleted unless disabled first.");
            return;
        }

        // Set the deleted trigger default to undefined.
        let deleted = undefined;

        // Prompt the user to confirm the removal.
        const answer = await window.showInformationMessage(`Do you want to delete the app role ${item.label}?`, "Yes", "No");

        // If the user confirms the removal then delete the role.
        if (answer === "Yes") {
            deleted = window.setStatusBarMessage("$(loading~spin) Deleting app role...");
            item.iconPath = new ThemeIcon("loading~spin");
            this._dataProvider.triggerOnDidChangeTreeData();

            // Get the parent application so we can read the app roles.
            const roles = await this.getAppRoles(item.objectId!);

            // Remove the app role from the array.
            roles.splice(roles.findIndex(r => r.id === item.value!), 1);

            // Update the application.
            await this._graphClient.updateApplication(item.objectId!, { appRoles: roles })
                .catch((error) => {
                    console.error(error);
                });
        }

        return deleted;
    }

    // Gets the app roles for an application registration.
    private async getAppRoles(id: string): Promise<AppRole[]> {
        return (await this._dataProvider.getApplicationPartial(id, "appRoles")).appRoles!;
    }

    // Updates an app role.
    private updateRole(role: AppRole, displayName: string, value: string, description: string, allowed: string, state: string) {
        role.displayName = displayName;
        role.value = value;
        role.description = description;
        role.allowedMemberTypes = allowed === "Users/Groups" ? ["User"] : allowed === "Applications" ? ["Application"] : ["User", "Application"];
        role.isEnabled = state === "Enabled" ? true : false;
    }
}