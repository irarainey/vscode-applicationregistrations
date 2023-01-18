import { window, env, Uri } from "vscode";
import { PORTAL_USER_URI } from "../constants";
import { AppRegTreeDataProvider } from "../data/app-reg-tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { ServiceBase } from "./service-base";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { User } from "@microsoft/microsoft-graph-types";
import { debounce } from "ts-debounce";
import { GraphResult } from "../types/graph-result";

export class OwnerService extends ServiceBase {

    // The list of users in the directory.
    private userList: any = undefined;

    // The constructor for the OwnerService class.
    constructor(graphRepository: GraphApiRepository, treeDataProvider: AppRegTreeDataProvider) {
        super(graphRepository, treeDataProvider);
    }

    // Adds a new owner to an application registration.
    async add(item: AppRegItem): Promise<void> {

        // Get the existing owners.
        const result: GraphResult<User[]> = await this.graphRepository.getApplicationOwners<User[]>(item.objectId!);
        if (result.success === true && result.value !== undefined) {
            // Debounce the validation function to prevent multiple calls to the Graph API.
            const validation = async (value: string) => this.validateOwner(value, result.value!);
            const debouncedValidation = debounce(validation, 500);

            // Prompt the user for the new owner.
            const owner = await window.showInputBox({
                placeHolder: "Enter user name or email address",
                prompt: "Add new owner to application",
                title: "Add Owner",
                ignoreFocusOut: true,
                validateInput: async (value) => await debouncedValidation(value)
            });

            // If the new owner name is not empty then add as an owner.
            if (owner !== undefined) {
                // Set the added trigger to the status bar message.
                const previousIcon = item.iconPath;
                const status = this.triggerTreeChange("Adding Owner", item);

                const result: GraphResult<undefined> = await this.graphRepository.addApplicationOwner<undefined>(item.objectId!, this.userList.value[0].id);
                if (result.success === true) {
                    this.triggerOnComplete({ success: true, statusBarHandle: status });
                } else {
                    this.triggerOnError({ success: false, statusBarHandle: status, error: result.error, treeViewItem: item, previousIcon: previousIcon, treeDataProvider: this.treeDataProvider });
                }
            }
        } else {
            this.triggerOnError({ success: false, error: result.error, treeDataProvider: this.treeDataProvider });
        }
    }

    // Removes an owner from an application registration.
    async remove(item: AppRegItem): Promise<void> {

        // Prompt the user to confirm the removal.
        const response = await window.showInformationMessage(`Do you want to remove ${item.label} as an owner of this application?`, "Yes", "No");

        // If the user confirms the removal then remove the user.
        if (response === "Yes") {
            // Set the added trigger to the status bar message.
            const previousIcon = item.iconPath;
            const status = this.triggerTreeChange("Removing Owner", item);

            const result: GraphResult<undefined> = await this.graphRepository.removeApplicationOwner<undefined>(item.objectId!, item.userId!);
            if (result.success === true) {
                this.triggerOnComplete({ success: true, statusBarHandle: status });
            } else {
                this.triggerOnError({ success: false, statusBarHandle: status, error: result.error, treeViewItem: item, previousIcon: previousIcon, treeDataProvider: this.treeDataProvider });
            }
        }
    }

    // Opens the user in the Azure Portal.
    openInPortal(item: AppRegItem): void {
        env.openExternal(Uri.parse(`${PORTAL_USER_URI}${item.userId}`));
    }

    // Validates the owner name or email address.
    private async validateOwner(owner: string, existing: User[]): Promise<string | undefined> {

        this.userList = undefined;
        let identifier: string = "";
        if (owner.indexOf('@') > -1) {
            // Try to find the user by email.
            this.userList = await this.graphRepository.findUsersByEmail(owner);
            identifier = "user with an email address";
        } else {
            // Try to find the user by name.
            this.userList = await this.graphRepository.findUsersByName(owner);
            identifier = "name";
        }

        if (this.userList.value.length === 0) {
            // User not found
            return `No ${identifier} beginning with ${owner} can be found in your directory.`;
        } else if (this.userList.value.length > 1) {
            // More than one user found
            return `More than one user with the ${identifier} beginning with ${owner} exists in your directory.`;
        }

        // Check if the user is already an owner.
        for (let i = 0; i < existing.length; i++) {
            if (existing[i].id === this.userList.value[0].id) {
                return `${owner} is already an owner of this application.`;
            }
        }

        return undefined;
    }
}
