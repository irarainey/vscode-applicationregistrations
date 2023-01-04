import { env, window } from 'vscode';
import { AppRegDataProvider } from '../data/applicationRegistration';
import { AppRegItem } from '../models/appRegItem';
import { addYears, isAfter, isBefore, isDate } from 'date-fns';
import { ServiceBase } from './serviceBase';
import { GraphClient } from '../clients/graph';

export class PasswordCredentialService extends ServiceBase {

    // The constructor for the PasswordCredentialsService class.
    constructor(dataProvider: AppRegDataProvider, graphClient: GraphClient) {
        super(dataProvider, graphClient);
    }

    // Adds a new password credential.
    public async add(item: AppRegItem): Promise<void> {

        // Prompt the user for the description.
        const description = await window.showInputBox({
            placeHolder: "Password description...",
            prompt: "Set new password credential description",
            ignoreFocusOut: true
        });

        // If the description is not undefined then add.
        if (description !== undefined) {
            const expiryDate = new Date();
            expiryDate.setDate(expiryDate.getDate() + 90);

            // Prompt the user for the description.
            const expiry = await window.showInputBox({
                placeHolder: "Password expiry...",
                prompt: "Set password expiry date",
                value: expiryDate.toISOString(),
                ignoreFocusOut: true,
                validateInput: (value) => this.validateExpiryDate(value)
            });

            if (expiry !== undefined) {
                // Set the added trigger to the status bar message.
                const previousIcon = item.iconPath;
                const status = this.indicateChange("Adding password credential...", item);
                this._graphClient.addPasswordCredential(item.objectId!, description, expiry)
                    .then((response) => {
                        env.clipboard.writeText(response.secretText!);
                        window.showInformationMessage("New password copied to clipboard.");
                        this._onComplete.fire({ success: true, statusBarHandle: status });
                    })
                    .catch((error) => {
                        this._onError.fire({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon });
                    });
            }
        }
    };

    // Deletes a password credential from an application registration.
    public async delete(item: AppRegItem): Promise<void> {

        // Prompt the user to confirm the removal.
        const answer = await window.showInformationMessage(`Do you want to delete this credential?`, "Yes", "No");

        // If the user confirms the removal then remove the user.
        if (answer === "Yes") {
            // Set the added trigger to the status bar message.
            const previousIcon = item.iconPath;
            const status = this.indicateChange("Deleting password credential...", item);
            this._graphClient.deletePasswordCredential(item.objectId!, item.value!)
                .then(() => {
                    this._onComplete.fire({ success: true, statusBarHandle: status });
                })
                .catch((error) => {
                    this._onError.fire({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon });
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