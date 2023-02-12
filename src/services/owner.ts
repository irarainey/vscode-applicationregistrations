import { window, env, Uri } from "vscode";
import { AZURE_PORTAL_APP_ROOT, AZURE_PORTAL_USER_PATH } from "../constants";
import { AppRegTreeDataProvider } from "../data/app-reg-tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { ServiceBase } from "./service-base";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { User } from "@microsoft/microsoft-graph-types";
import { debounce } from "ts-debounce";
import { GraphResult } from "../types/graph-result";
import { AccountProvider } from "../data/account-provider";

export class OwnerService extends ServiceBase {

    private accountProvider : AccountProvider;

    // The list of users in the directory.
    private userList: any = undefined;

    // The constructor for the OwnerService class.
    constructor(graphRepository: GraphApiRepository, treeDataProvider: AppRegTreeDataProvider, accountProvider : AccountProvider) {
        super(graphRepository, treeDataProvider);
        this.accountProvider = accountProvider;
    }

    // Adds a new owner to an application registration.
    async add(item: AppRegItem): Promise<void> {

        // Get the existing owners.
        const result: GraphResult<User[]> = await this.graphRepository.getApplicationOwners(item.objectId!);
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
                const status = this.indicateChange("Adding Owner...", item);
                const result: GraphResult<void> = await this.graphRepository.addApplicationOwner(item.objectId!, this.userList[0].id);
                result.success === true ? await this.triggerRefresh(status) : await this.handleError(result.error);
            }
        } else {
            await this.handleError(result.error);
        }
    }

    // Removes an owner from an application registration.
    async remove(item: AppRegItem): Promise<void> {

        // Prompt the user to confirm the removal.
        const response = await window.showWarningMessage(`Do you want to remove ${item.label} as an owner of this application?`, "Yes", "No");

        // If the user confirms the removal then remove the user.
        if (response === "Yes") {
            // Set the added trigger to the status bar message.
            const status = this.indicateChange("Removing Owner...", item);
            const result: GraphResult<void> = await this.graphRepository.removeApplicationOwner(item.objectId!, item.userId!);
            result.success === true ? await this.triggerRefresh(status) : await this.handleError(result.error);
        }
    }

    // Opens the user in the Azure Portal.
    async openInAzurePortal(item: AppRegItem): Promise<void> {
        const accountInformation = await this.accountProvider.getAccountInformation();
        let uriText = "";
        if (accountInformation.tenantId){
            uriText = `${AZURE_PORTAL_APP_ROOT}/${accountInformation.tenantId}${AZURE_PORTAL_USER_PATH}${item.userId}`;
        }
        else{
            uriText = `${AZURE_PORTAL_APP_ROOT}${AZURE_PORTAL_USER_PATH}${item.userId}`;
        }
        env.openExternal(Uri.parse(uriText));
    }

    // Validates the owner name or email address.
    private async validateOwner(owner: string, existing: User[]): Promise<string | undefined> {

        // Check if the owner name is empty.
        if (owner === undefined || owner === null || owner.length === 0) {
            return undefined;
        }

        this.userList = undefined;
        let identifier: string = "";
        if (owner.indexOf('@') > -1) {
            // Try to find the user by email.
            const result: GraphResult<User[]> = await this.graphRepository.findUsersByEmail(owner);
            if (result.success === true && result.value !== undefined) {
                this.userList = result.value;
                identifier = "user with an email address";
            } else {
                await this.handleError(result.error);
                return;
            }
        } else {
            // Try to find the user by name.
            const result: GraphResult<User[]> = await this.graphRepository.findUsersByName(owner);
            if (result.success === true && result.value !== undefined) {
                this.userList = result.value;
                identifier = "name";
            } else {
                await this.handleError(result.error);
                return;
            }
        }

        if (this.userList.length === 0) {
            // User not found
            return `No ${identifier} beginning with ${owner} can be found in your directory.`;
        } else if (this.userList.length > 1) {
            // More than one user found
            return `More than one user with the ${identifier} beginning with ${owner} exists in your directory.`;
        }

        // Check if the user is already an owner.
        for (let i = 0; i < existing.length; i++) {
            if (existing[i].id === this.userList[0].id) {
                return `${owner} is already an owner of this application.`;
            }
        }

        return undefined;
    }
}
