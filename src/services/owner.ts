import { window, env, Uri, workspace } from "vscode";
import { AZURE_PORTAL_ROOT, ENTRA_PORTAL_ROOT, AZURE_AND_ENTRA_PORTAL_USER_PATH } from "../constants";
import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { ServiceBase } from "./service-base";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { User } from "@microsoft/microsoft-graph-types";
import { debounce } from "ts-debounce";
import { GraphResult } from "../types/graph-result";
import { validateOwner } from "../utils/validation";
import { OwnerList } from "../models/owner-list";
import { AccountProvider } from "../types/account-provider";

export class OwnerService extends ServiceBase {

	// The account provider.
    private accountProvider : AccountProvider;

	// The list of users in the directory.
	private owners = new OwnerList();

	// The constructor for the OwnerService class.
	constructor(graphRepository: GraphApiRepository, treeDataProvider: AppRegTreeDataProvider, accountProvider : AccountProvider) {
		super(graphRepository, treeDataProvider);
        this.accountProvider = accountProvider;
	}

	// Sets the user list.
	setUserList(userList: any): void {
		this.owners.users = userList;
	}

	// Adds a new owner to an application registration.
	async add(item: AppRegItem): Promise<void> {
		// Get the existing owners.
		const result: GraphResult<User[]> = await this.graphRepository.getApplicationOwners(item.objectId!);
		if (result.success === true && result.value !== undefined) {
			// Debounce the validation function to prevent multiple calls to the Graph API.
			const validation = async (value: string) => validateOwner(value, result.value!, this.graphRepository, this.owners);
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
				const result: GraphResult<void> = await this.graphRepository.addApplicationOwner(item.objectId!, this.owners.users![0].id!);
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
    async openInAzurePortal(item: AppRegItem): Promise<boolean> {
        return await this.openInPortal(AZURE_PORTAL_ROOT, item);
    }

    // Opens the user in the Entra Portal.
    async openInEntraPortal(item: AppRegItem): Promise<boolean> {
		return await this.openInPortal(ENTRA_PORTAL_ROOT, item);
    }

    // Opens the user in the Azure or Entra Portal.
    async openInPortal(portalRoot: string, item: AppRegItem): Promise<boolean>{
        // Determine if "omit tenant ID from portal requests" has been set.
        const omitTenantIdFromPortalRequests = workspace.getConfiguration("applicationRegistrations").get("omitTenantIdFromPortalRequests") as boolean;

        let uriText = "";
        if (omitTenantIdFromPortalRequests === false){
            const accountInformation = await this.accountProvider.getAccountInformation();
            if (accountInformation.tenantId){
                uriText = `${portalRoot}/${accountInformation.tenantId}${AZURE_AND_ENTRA_PORTAL_USER_PATH}${item.userId}`;
            }
        }

        // Check if uriText has been set to anything meaningful yet. If not, then set it w/o a tenant ID.
        if (typeof(uriText) === "string" && uriText.trim().length === 0){
            uriText = `${portalRoot}${AZURE_AND_ENTRA_PORTAL_USER_PATH}${item.userId}`;
        }
		
        return await env.openExternal(Uri.parse(uriText));
    }
}
