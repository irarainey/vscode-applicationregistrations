import { env, window } from "vscode";
import { AppRegTreeDataProvider } from "../data/app-reg-tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { ServiceBase } from "./service-base";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { format } from "date-fns";
import { GraphResult } from "../types/graph-result";
import { PasswordCredential } from "@microsoft/microsoft-graph-types";
import { validatePasswordCredentialExpiryDate } from "../utils/validation";

export class PasswordCredentialService extends ServiceBase {
	// The constructor for the PasswordCredentialsService class.
	constructor(graphRepository: GraphApiRepository, treeDataProvider: AppRegTreeDataProvider) {
		super(graphRepository, treeDataProvider);
	}

	// Adds a new password credential.
	async add(item: AppRegItem): Promise<void> {
		// Prompt the user for the description.
		const description = await window.showInputBox({
			placeHolder: "Password description",
			prompt: "Set new password credential description",
			title: "Add Password Credential (1/2)",
			ignoreFocusOut: true
		});

		// If the description is not undefined then add.
		if (description !== undefined) {
			const expiryDate = new Date();
			expiryDate.setDate(expiryDate.getDate() + 90);

			// Prompt the user for the description.
			const expiry = await window.showInputBox({
				placeHolder: "Password expiry",
				prompt: "Set password expiry date",
				value: format(new Date(expiryDate), "yyyy-MM-dd"),
				title: "Add Password Credential (2/2)",
				ignoreFocusOut: true,
				validateInput: (value) => validatePasswordCredentialExpiryDate(value)
			});

			if (expiry !== undefined) {
				// Set the added trigger to the status bar message.
				const status = this.indicateChange("Adding Password Credential...", item);
				const update: GraphResult<PasswordCredential> = await this.graphRepository.addPasswordCredential(item.objectId!, description, expiry);
				if (update.success === true && update.value !== undefined) {
					env.clipboard.writeText(update.value.secretText!);
					await this.triggerRefresh(status);
					window.showInformationMessage("New password copied to clipboard.", "OK");
				} else {
					await this.handleError(update.error);
				}
			}
		}
	}

	// Deletes a password credential from an application registration.
	async delete(item: AppRegItem): Promise<void> {
		// Prompt the user to confirm the removal.
		const answer = await window.showWarningMessage("Do you want to delete this password credential?", "Yes", "No");

		// If the user confirms the removal then remove the password credential.
		if (answer === "Yes") {
			// Set the added trigger to the status bar message.
			const status = this.indicateChange("Deleting Password Credential...", item);
			const update: GraphResult<void> = await this.graphRepository.deletePasswordCredential(item.objectId!, item.value!);
			update.success === true ? await this.triggerRefresh(status) : await this.handleError(update.error);
		}
	}
}
