import { env, window } from "vscode";
import { AppRegTreeDataProvider } from "../data/app-reg-tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { addYears, isAfter, isBefore, isDate } from "date-fns";
import { ServiceBase } from "./service-base";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { format } from "date-fns";
import { GraphResult } from "../types/graph-result";
import { PasswordCredential } from "@microsoft/microsoft-graph-types";


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
                value: format(new Date(expiryDate), 'yyyy-MM-dd'),
                title: "Add Password Credential (2/2)",
                ignoreFocusOut: true,
                validateInput: (value) => this.validateExpiryDate(value)
            });

            if (expiry !== undefined) {
                // Set the added trigger to the status bar message.
                const status = this.indicateChange("Adding Password Credential...", item);
                const update: GraphResult<PasswordCredential> = await this.graphRepository.addPasswordCredential(item.objectId!, description, expiry);
                if (update.success === true && update.value !== undefined) {
                    env.clipboard.writeText(update.value.secretText!);
                    this.triggerOnComplete(status);
                    window.showInformationMessage("New password copied to clipboard.", "OK");
                } else {
                    this.triggerOnError(update.error);
                }
            }
        }
    }

    // Deletes a password credential from an application registration.
    async delete(item: AppRegItem): Promise<void> {

        // Prompt the user to confirm the removal.
        const answer = await window.showInformationMessage("Do you want to delete this password credential?", "Yes", "No");

        // If the user confirms the removal then remove the password credential.
        if (answer === "Yes") {
            // Set the added trigger to the status bar message.
            const status = this.indicateChange("Deleting Password Credential...", item);
            const update: GraphResult<void> = await this.graphRepository.deletePasswordCredential(item.objectId!, item.value!);
            update.success === true ? this.triggerOnComplete(status) : this.triggerOnError(update.error);
        }
    }

    // Validates the expiry date.
    private validateExpiryDate(expiry: string): string | undefined {

        const expiryDate = Date.parse(expiry);
        const now = Date.now();
        const maximumExpireDate = addYears(new Date(), 2);

        //Check if the expiry date is a valid date.
        if (isDate(expiryDate)) {
            return "Expiry must be a valid date.";
        }

        // Check if the expiry date is in the future.
        if (isBefore(expiryDate, now)) {
            return "Expiry must be in the future.";
        }

        // Check if the expiry date is less than 2 years in the future.
        if (isAfter(expiryDate, maximumExpireDate)) {
            return "Expiry must be less than 2 years in the future.";
        }

        return undefined;
    }
}