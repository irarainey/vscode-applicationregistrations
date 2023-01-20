import { window } from "vscode";
import { AppRegTreeDataProvider } from "../data/app-reg-tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { AppRole, Application } from "@microsoft/microsoft-graph-types";
import { v4 as uuidv4 } from "uuid";
import { ServiceBase } from "./service-base";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { debounce } from "ts-debounce";
import { GraphResult } from "../types/graph-result";

export class AppRoleService extends ServiceBase {

    // The constructor for the AppRolesService class.
    constructor(graphRepository: GraphApiRepository, treeDataProvider: AppRegTreeDataProvider) {
        super(graphRepository, treeDataProvider);
    }

    // Adds a new app role to an application registration.
    async add(item: AppRegItem): Promise<void> {

        // Capture the new app role details by passing in an empty app role.
        const role = await this.inputRoleDetails({}, item.objectId!, false);

        // If the user cancels the input then return undefined.
        if (role === undefined) {
            return;
        }

        // Set the added trigger to the status bar message.
        this.triggerTreeChange("Adding App Role", item);

        // Get the existing app roles.
        const roles = await this.getAppRoles(item.objectId!);

        // If the array is undefined then it'll be an Azure CLI authentication issue.
        if (roles === undefined) {
            return;
        }

        // Add the new app role to the array.
        roles.push(role);

        // Update the application.
        await this.updateApplication(item.objectId!, { appRoles: roles });
    }

    // Edits an app role from an application registration.
    async edit(item: AppRegItem): Promise<void> {

        // Get the parent application so we can read the existing app roles.
        const roles = await this.getAppRoles(item.objectId!);

        // If the array is undefined then it'll be an Azure CLI authentication issue.
        if (roles === undefined) {
            return;
        }

        // Capture the new app role details by passing in the existing role.
        const role = await this.inputRoleDetails(roles.filter(r => r.id === item.value!)[0], item.objectId!, true);

        // If the user cancels the input then return undefined.
        if (role === undefined) {
            return;
        }

        // Set the added trigger to the status bar message.
        this.triggerTreeChange("Updating App Role", item);

        // Update the application.
        await this.updateApplication(item.objectId!, { appRoles: roles });
    }

    // Changes the enabled state of an app role from an application registration.
    async changeState(item: AppRegItem, state: boolean): Promise<void> {

        // Set the added trigger to the status bar message.
        this.triggerTreeChange(state === true ? "Enabling App Role" : "Disabling App Role", item);

        // Get the parent application so we can read the app roles.
        const roles = await this.getAppRoles(item.objectId!);

        // If the array is undefined then it'll be an Azure CLI authentication issue.
        if (roles === undefined) {
            return;
        }

        // Toggle the state of the app role.
        roles.filter(r => r.id === item.value!)[0].isEnabled = state;

        // Update the application.
        await this.updateApplication(item.objectId!, { appRoles: roles });
    }

    // Deletes an app role from an application registration.
    async delete(item: AppRegItem): Promise<void> {

        if (item.state !== false) {
            window.showWarningMessage("Role cannot be deleted unless disabled first.", "OK");
            return;
        }

        // Prompt the user to confirm the removal.
        const answer = await window.showInformationMessage(`Do you want to delete the App Role ${item.label}?`, "Yes", "No");

        // If the user confirms the removal then delete the role.
        if (answer === "Yes") {
            // Set the added trigger to the status bar message.
            this.triggerTreeChange("Deleting App Role", item);

            // Get the parent application so we can read the app roles.
            const roles = await this.getAppRoles(item.objectId!);

            // If the array is undefined then it'll be an Azure CLI authentication issue.
            if (roles === undefined) {
                return;
            }

            // Remove the app role from the array.
            roles.splice(roles.findIndex(r => r.id === item.value!), 1);

            // Update the application.
            await this.updateApplication(item.objectId!, { appRoles: roles });
        }
    }

    // Updates an application registration with the new app roles.
    private async updateApplication(id: string, application: Application) {
        const update: GraphResult<void> = await this.graphRepository.updateApplication(id, application);
        update.success === true ? this.triggerOnComplete() : this.triggerOnError(update.error);
    }

    // Gets the app roles for an application registration.
    private async getAppRoles(id: string): Promise<AppRole[] | undefined> {
        const result: GraphResult<Application> = await this.graphRepository.getApplicationDetailsPartial(id, "appRoles");
        if (result.success === true && result.value !== undefined) {
            return result.value?.appRoles;
        }
        else {
            this.triggerOnError(result.error);
            return undefined;
        }
    }

    // Captures the details for an app role.
    private async inputRoleDetails(role: AppRole, id: string, isEditing: boolean): Promise<AppRole | undefined> {

        // Prompt the user for the new display name.
        const displayName = await window.showInputBox({
            prompt: "App Role Display Name",
            placeHolder: "Enter a display name for the App Role",
            ignoreFocusOut: true,
            title: isEditing === true ? "Edit App Role (1/5)" : "Add App Role (1/5)",
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
            prompt: "App Role Value",
            placeHolder: "Enter a value for the App Role",
            title: isEditing === true ? "Edit App Role (2/5)" : "Add App Role (2/5)",
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
            prompt: "App Role Description",
            placeHolder: "Enter a description for the App Role",
            title: isEditing === true ? "Edit App Role (3/5)" : "Add App Role (3/5)",
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
                {
                    label: "Users/Groups",
                    description: "Users and groups can be assigned this role.",
                    value: ["User"]
                },
                {
                    label: "Applications",
                    description: "Only applications can be assigned this role.",
                    value: ["Application"]
                },
                {
                    label: "Both (Users/Groups + Applications)",
                    description: "Users, groups and applications can be assigned this role.",
                    value: ["User", "Application"]
                }
            ],
            {
                placeHolder: "Select allowed member types",
                title: isEditing === true ? "Edit App Role (4/5)" : "Add App Role (4/5)",
                ignoreFocusOut: true
            });

        // If escape is pressed or the new allowed member types is empty then return undefined.
        if (allowed === undefined) {
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
                placeHolder: "Select role state",
                title: isEditing === true ? "Edit App Role (5/5)" : "Add App Role (5/5)",
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
        role.allowedMemberTypes = allowed.value;
        role.isEnabled = state.value;

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

            // If the array is undefined then it'll be an Azure CLI authentication issue.
            if (roles === undefined) {
                return undefined;
            }

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