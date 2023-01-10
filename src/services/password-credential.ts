import { env, window } from "vscode";
import { AppRegTreeDataProvider } from "../data/app-reg-tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { addYears, isAfter, isBefore, isDate } from "date-fns";
import { ServiceBase } from "./service-base";
import { GraphClient } from "../clients/graph-client";
import { format } from "date-fns";

export class PasswordCredentialService extends ServiceBase {

    // The constructor for the PasswordCredentialsService class.
    constructor(graphClient: GraphClient, treeDataProvider: AppRegTreeDataProvider) {
        super(graphClient, treeDataProvider);
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
                const previousIcon = item.iconPath;
                const status = this.triggerTreeChange("Adding Password Credential", item);
                this.graphClient.addPasswordCredential(item.objectId!, description, expiry)
                    .then((response) => {
                        env.clipboard.writeText(response.secretText!);
                        this.triggerOnComplete({ success: true, statusBarHandle: status });
                        window.showInformationMessage("New password copied to clipboard.", "OK");
                    })
                    .catch((error) => {
                        this.triggerOnError({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon });
                    });
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
            const previousIcon = item.iconPath;
            const status = this.triggerTreeChange("Deleting Password Credential", item);
            this.graphClient.deletePasswordCredential(item.objectId!, item.value!)
                .then(() => {
                    this.triggerOnComplete({ success: true, statusBarHandle: status });
                })
                .catch((error) => {
                    this.triggerOnError({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon });
                });
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