import { env, window, ThemeIcon, Disposable } from 'vscode';
import { GraphClient } from '../clients/graph';
import { AppRegDataProvider } from '../data/applicationRegistration';
import { AppRegItem } from '../models/appRegItem';
import { addYears, isAfter, isBefore, isDate } from 'date-fns';

export class PasswordCredentialsService {

    // A private instance of the GraphClient class.
    private _graphClient: GraphClient;

    // A private instance of the AppRegDataProvider class.
    private _dataProvider: AppRegDataProvider;

    // The constructor for the PasswordCredentialsService class.
    constructor(dataProvider: AppRegDataProvider) {
        this._dataProvider = dataProvider;
        this._graphClient = dataProvider.graphClient;
    }

    // Adds a new password credential.
    public async add(item: AppRegItem): Promise<Disposable | undefined> {

        // Set the created trigger default to undefined.
        let added = undefined;

        // Prompt the user for the description.
        const description = await window.showInputBox({
            placeHolder: "Password description...",
            prompt: "Add new password credential description..."
        });

        // If the description is not undefined then add.
        if (description !== undefined) {
            added = window.setStatusBarMessage("$(loading~spin) Adding password credential...");

            const expiryDate = new Date();
            expiryDate.setDate(expiryDate.getDate() + 90);

            // Prompt the user for the description.
            const expiry = await window.showInputBox({
                placeHolder: "Password expiry...",
                prompt: "Add new password expiry...",
                value: expiryDate.toISOString(),
                validateInput: (value) => {
                    return this.validateExpiryDate(value);
                }
            });

            if (expiry !== undefined) {
                item.iconPath = new ThemeIcon("loading~spin");
                this._dataProvider.triggerOnDidChangeTreeData();
                await this._graphClient.addPasswordCredential(item.objectId!, description, expiry)
                    .then((response) => {
                        env.clipboard.writeText(response.secretText!);
                        window.showInformationMessage("New password copied to clipboard.");
                    })
                    .catch((error) => {
                        console.error(error);
                    });
            }
        }

        // Return the state of the action to refresh the list if required.
        return added;
    };

    // Deletes a password credential from an application registration.
    public async delete(item: AppRegItem): Promise<Disposable | undefined> {

        // Set the removed trigger default to undefined.
        let deleted = undefined;

        // Prompt the user to confirm the removal.
        const answer = await window.showInformationMessage(`Do you want to delete this credential?`, "Yes", "No");

        // If the user confirms the removal then remove the user.
        if (answer === "Yes") {
            deleted = window.setStatusBarMessage("$(loading~spin) Removing credential...");
            item.iconPath = new ThemeIcon("loading~spin");
            this._dataProvider.triggerOnDidChangeTreeData();
            await this._graphClient.deletePasswordCredential(item.objectId!, item.value!)
                .catch((error) => {
                    console.error(error);
                });
        }

        return deleted;
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