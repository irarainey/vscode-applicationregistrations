import { window } from 'vscode';
import { AppRegTreeDataProvider } from '../data/appRegTreeDataProvider';
import { AppRegItem } from '../models/appRegItem';
import { AppRole } from "@microsoft/microsoft-graph-types";
import { v4 as uuidv4 } from 'uuid';
import { ServiceBase } from './serviceBase';
import { GraphClient } from '../clients/graph';
import { debounce } from "ts-debounce";

export class AppRoleService extends ServiceBase {

    // The constructor for the AppRolesService class.
    constructor(treeDataProvider: AppRegTreeDataProvider, graphClient: GraphClient) {
        super(treeDataProvider, graphClient);
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
        const status = this.triggerTreeChange("Adding new app role...", item);

        // Get the existing app roles.
        const roles = await this.getAppRoles(item.objectId!);

        // Add the new app role to the array.
        roles.push(role);

        // Update the application.
        this.graphClient.updateApplication(item.objectId!, { appRoles: roles })
            .then(() => {
                this.triggerOnComplete({ success: true, statusBarHandle: status });
            })
            .catch((error) => {
                this.triggerOnError({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon });
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
        const status = this.triggerTreeChange("Updating app role...", item);

        // Update the application.
        this.graphClient.updateApplication(item.objectId!, { appRoles: roles })
            .then(() => {
                this.triggerOnComplete({ success: true, statusBarHandle: status });
            })
            .catch((error) => {
                this.triggerOnError({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon });
            });
    }

    // Changes the enabled state of an app role from an application registration.
    public async changeState(item: AppRegItem): Promise<void> {

        // Set the added trigger to the status bar message.
        const previousIcon = item.iconPath;
        const status = this.triggerTreeChange("Updating app role state...", item);

        // Get the parent application so we can read the app roles.
        const roles = await this.getAppRoles(item.objectId!);

        // Toggle the state of the app role.
        roles.filter(r => r.id === item.value!)[0].isEnabled = !item.state;

        // Update the application.
        this.graphClient.updateApplication(item.objectId!, { appRoles: roles })
            .then(() => {
                this.triggerOnComplete({ success: true, statusBarHandle: status });
            })
            .catch((error) => {
                this.triggerOnError({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon });
            });
    }

    // Deletes an app role from an application registration.
    public async delete(item: AppRegItem): Promise<void> {

        if (item.state !== false) {
            window.showWarningMessage("Role cannot be deleted unless disabled first.", "OK");
            return;
        }

        // Prompt the user to confirm the removal.
        const answer = await window.showInformationMessage(`Do you want to delete the app role ${item.label}?`, "Yes", "No");

        // If the user confirms the removal then delete the role.
        if (answer === "Yes") {
            // Set the added trigger to the status bar message.
            const previousIcon = item.iconPath;
            const status = this.triggerTreeChange("Deleting app role...", item);

            // Get the parent application so we can read the app roles.
            const roles = await this.getAppRoles(item.objectId!);

            // Remove the app role from the array.
            roles.splice(roles.findIndex(r => r.id === item.value!), 1);

            // Update the application.
            this.graphClient.updateApplication(item.objectId!, { appRoles: roles })
                .then(() => {
                    this.triggerOnComplete({ success: true, statusBarHandle: status });
                })
                .catch((error) => {
                    this.triggerOnError({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon });
                });
        }
    }

    // Gets the app roles for an application registration.
    private async getAppRoles(id: string): Promise<AppRole[]> {
        return (await this.dataProvider.getApplicationPartial(id, "appRoles")).appRoles!;
    }

    // Captures the details for an app role.
    private async inputRoleDetails(role: AppRole, id: string, isEditing: boolean): Promise<AppRole | undefined> {

        // Prompt the user for the new display name.
        const displayName = await window.showInputBox({
            prompt: "Edit display name",
            placeHolder: "Enter a display name for the new app role",
            ignoreFocusOut: true,
            value: role.displayName ?? undefined,
            validateInput: (value) => this.validateDisplayName(value)
        });

        // If escape is pressed then return undefined.
        if (displayName === undefined) {
            return undefined;
        }

        // Debounce the validation function to prevent multiple calls to the Graph API.
        const validation = async (value: string, id: string, isEditing: boolean, oldValue: string | undefined) => this.validateValue(value, id, isEditing, role.value ?? undefined);
        const debouncedValidation = debounce(validation, 500);

        // Prompt the user for the new value.
        const value = await window.showInputBox({
            prompt: "Edit value",
            placeHolder: "Enter a value for the new app role",
            ignoreFocusOut: true,
            value: role.value ?? undefined,
            validateInput: async (value) => debouncedValidation(value, id, isEditing, role.value ?? undefined)
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
            value: role.description ?? undefined,
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

        // Check the length of the value.
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

        // Check the length of the description.
        if (description.length < 1) {
            return "A description cannot be empty.";
        }

        return undefined;
    }
}