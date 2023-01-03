import { window, ThemeIcon, Event, EventEmitter, Disposable } from 'vscode';
import { GraphClient } from '../clients/graph';
import { AppRegDataProvider } from '../data/applicationRegistration';
import { AppRegItem } from '../models/appRegItem';
import { ActivityStatus } from '../interfaces/activityStatus';
import { AppRole } from "@microsoft/microsoft-graph-types";
import { v4 as uuidv4 } from 'uuid';

export class AppRoleService {

    // A private instance of the GraphClient class.
    private _graphClient: GraphClient;

    // A private instance of the AppRegDataProvider class.
    private _dataProvider: AppRegDataProvider;

    // A private instance of the EventEmitter class to handle error events.
    private _onError: EventEmitter<ActivityStatus> = new EventEmitter<ActivityStatus>();

    // A private instance of the EventEmitter class to handle complete events.
    private _onComplete: EventEmitter<ActivityStatus> = new EventEmitter<ActivityStatus>();

    // A public readonly property to expose the error event.
    public readonly onError: Event<ActivityStatus> = this._onError.event;

    // A public readonly property to expose the complete event.
    public readonly onComplete: Event<ActivityStatus> = this._onComplete.event;

    // The constructor for the AppRolesService class.
    constructor(dataProvider: AppRegDataProvider) {
        this._dataProvider = dataProvider;
        this._graphClient = dataProvider.graphClient;
    }

    // Adds a new app role to an application registration.
    public async add(item: AppRegItem): Promise<void> {

        // Capture the new app role details by passing in an empty app role.
        const role = await this.inputRoleDetails({}, item.objectId!, false);

        // If the user cancels the input then return undefined.
        if (role === undefined) {
            return;
        }

        // Set the added trigger to the status bar message.
        const previousIcon = item.iconPath;
        const status = this.indicateChange("Adding new app role...", item);

        // Get the existing app roles.
        const roles = await this.getAppRoles(item.objectId!);

        // Add the new app role to the array.
        roles.push(role);

        // Update the application.
        await this._graphClient.updateApplication(item.objectId!, { appRoles: roles })
            .then(() => {
                this._onComplete.fire({ success: true, statusBarHandle: status });
            })
            .catch((error) => {
                this._onError.fire({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon });
            });
    }

    // Edits an app role from an application registration.
    public async edit(item: AppRegItem): Promise<void> {

        // Get the parent application so we can read the existing app roles.
        const roles = await this.getAppRoles(item.objectId!);

        // Capture the new app role details by passing in the existing role.
        const role = await this.inputRoleDetails(roles.filter(r => r.id === item.value!)[0], item.objectId!, true);

        // If the user cancels the input then return undefined.
        if (role === undefined) {
            return;
        }

        // Set the added trigger to the status bar message.
        const previousIcon = item.iconPath;
        const status = this.indicateChange("Updating app role...", item);

        // Update the application.
        this._graphClient.updateApplication(item.objectId!, { appRoles: roles })
            .then(() => {
                this._onComplete.fire({ success: true, statusBarHandle: status });
            })
            .catch((error) => {
                this._onError.fire({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon });
            });
    }

    // Changes the enabled state of an app role from an application registration.
    public async changeState(item: AppRegItem): Promise<void> {

        // Set the added trigger to the status bar message.
        const previousIcon = item.iconPath;
        const status = this.indicateChange("Updating app role state...", item);

        // Get the parent application so we can read the app roles.
        const roles = await this.getAppRoles(item.objectId!);

        // Toggle the state of the app role.
        roles.filter(r => r.id === item.value!)[0].isEnabled = !item.state;

        // Update the application.
        this._graphClient.updateApplication(item.objectId!, { appRoles: roles })
            .then(() => {
                this._onComplete.fire({ success: true, statusBarHandle: status });
            })
            .catch((error) => {
                this._onError.fire({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon });
            });
    }

    // Deletes an app role from an application registration.
    public async delete(item: AppRegItem): Promise<void> {

        if (item.state !== false) {
            window.showWarningMessage("Role cannot be deleted unless disabled first.");
            return;
        }

        // Prompt the user to confirm the removal.
        const answer = await window.showInformationMessage(`Do you want to delete the app role ${item.label}?`, "Yes", "No");

        // If the user confirms the removal then delete the role.
        if (answer === "Yes") {
            // Set the added trigger to the status bar message.
            const previousIcon = item.iconPath;
            const status = this.indicateChange("Deleting app role...", item);

            // Get the parent application so we can read the app roles.
            const roles = await this.getAppRoles(item.objectId!);

            // Remove the app role from the array.
            roles.splice(roles.findIndex(r => r.id === item.value!), 1);

            // Update the application.
            await this._graphClient.updateApplication(item.objectId!, { appRoles: roles })
                .then(() => {
                    this._onComplete.fire({ success: true, statusBarHandle: status });
                })
                .catch((error) => {
                    this._onError.fire({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon });
                });
        }
    }

    // Gets the app roles for an application registration.
    private async getAppRoles(id: string): Promise<AppRole[]> {
        return (await this._dataProvider.getApplicationPartial(id, "appRoles")).appRoles!;
    }

    // Indicates that a change is taking place.
    private indicateChange(statusBarMessage: string, item: AppRegItem): Disposable {
        item.iconPath = new ThemeIcon("loading~spin");
        this._dataProvider.triggerOnDidChangeTreeData();
        return window.setStatusBarMessage(`$(loading~spin) ${statusBarMessage}`);
    }

    // Captures the details for an app role.
    private async inputRoleDetails(role: AppRole, id: string, isEditing: boolean): Promise<AppRole | undefined> {

        // Prompt the user for the new display name.
        const displayName = await window.showInputBox({
            prompt: "Edit display name",
            placeHolder: "Enter a display name for the new app role",
            ignoreFocusOut: true,
            value: role.displayName ?? "",
            validateInput: (value) => this.validateDisplayName(value)
        });

        // If escape is pressed then return undefined.
        if (displayName === undefined) {
            return undefined;
        }

        // Prompt the user for the new value.
        const value = await window.showInputBox({
            prompt: "Edit value",
            placeHolder: "Enter a value for the new app role",
            ignoreFocusOut: true,
            value: role.value ?? "",
            validateInput: async (value) => this.validateValue(value, id, isEditing, role.value ?? undefined)
        });

        // If escape is pressed then return undefined.
        if (value === undefined) {
            return undefined;
        }

        // Prompt the user for the new display name.
        const description = await window.showInputBox({
            prompt: "Edit description",
            placeHolder: "Enter a description for the new app role",
            ignoreFocusOut: true,
            value: role.description ?? "",
            validateInput: (value) => this.validateDescription(value)
        });

        // If escape is pressed then return undefined.
        if (description === undefined) {
            return undefined;
        }

        // Prompt the user for the new allowed member types.
        const allowed = await window.showQuickPick(
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
        if (allowed === undefined) {
            return undefined;
        }

        // Prompt the user for the new state.
        const state = await window.showQuickPick(
            [
                "Enabled",
                "Disabled"
            ],
            {
                placeHolder: "Select role state",
                ignoreFocusOut: true
            });

        // If escape is pressed or the new state is empty then return undefined.
        if (state === undefined) {
            return undefined;
        }

        // Set the new values on the app role.
        role.id = role.id ?? uuidv4();
        role.displayName = displayName;
        role.value = value;
        role.description = description;
        role.allowedMemberTypes = allowed === "Users/Groups" ? ["User"] : allowed === "Applications" ? ["Application"] : ["User", "Application"];
        role.isEnabled = state === "Enabled" ? true : false;

        return role;
    }

    // Validates the display name of an app role.
    private validateDisplayName(displayName: string): string | undefined {

        // Check the length of the display name.
        if (displayName.length > 100) {
            return "A display name cannot be longer than 100 characters.";
        }

        // Check the length of the display name.
        if (displayName.length < 1) {
            return "A display name cannot be empty.";
        }

        return undefined;
    }

    // Validates the value of an app role.
    private async validateValue(value: string, id: string, isEditing: boolean, oldValue: string | undefined): Promise<string | undefined> {

        // Check the length of the display name.
        if (value.length > 250) {
            return "A value cannot be longer than 250 characters.";
        }

        // Check the length of the display name.
        if (value.length < 1) {
            return "A value cannot be empty.";
        }

        // Check to see if the value already exists.
        if (isEditing !== true && oldValue !== value) {
            const roles = await this.getAppRoles(id);
            if (roles!.find(r => r.value === value) !== undefined) {
                return "The value specified already exists.";
            }
        }

        return undefined;
    }

    // Validates the value of an app role.
    private validateDescription(description: string): string | undefined {

        // Check the length of the display name.
        if (description.length < 1) {
            return "A description cannot be empty.";
        }

        return undefined;
    }
}