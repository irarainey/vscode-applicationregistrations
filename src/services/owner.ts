import { window, env, Uri } from "vscode";
import { PORTAL_USER_URI } from "../constants";
import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { ServiceBase } from "./service-base";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { User } from "@microsoft/microsoft-graph-types";
import { debounce } from "ts-debounce";
import { GraphResult } from "../types/graph-result";
import { validateOwner } from "../utils/validation";
import { OwnerList } from "../models/owner-list";

export class OwnerService extends ServiceBase {
	// The list of users in the directory.
	private owners = new OwnerList();

	// The constructor for the OwnerService class.
	constructor(graphRepository: GraphApiRepository, treeDataProvider: AppRegTreeDataProvider) {
		super(graphRepository, treeDataProvider);
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
	openInPortal(item: AppRegItem): void {
		env.openExternal(Uri.parse(`${PORTAL_USER_URI}${item.userId}`));
	}
}
