import { window } from "vscode";
import { AppRegTreeDataProvider } from "../data/app-reg-tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { ServiceBase } from "./service-base";
import { GraphClient } from "../clients/graph-client";
import { KeyCredential } from "@microsoft/microsoft-graph-types";

export class KeyCredentialService extends ServiceBase {

    // The constructor for the KeyCredentialsService class.
    constructor(treeDataProvider: AppRegTreeDataProvider, graphClient: GraphClient) {
        super(treeDataProvider, graphClient);
    }

    // Adds a new key credential by uploading a certificate.
    async upload(item: AppRegItem): Promise<void> {

    }

    // Deletes a key credential from an application registration.
    async delete(item: AppRegItem): Promise<void> {

        // Prompt the user to confirm the removal.
        const answer = await window.showInformationMessage("Do you want to delete this certificate credential?", "Yes", "No");

        // If the user confirms the removal then remove the key credential.
        if (answer === "Yes") {

            // Get all the key credentials for the application.
            const keyCredentials = await this.getKeyCredentials(item.objectId!);

            // Remove the scope requested.
            keyCredentials!.splice(keyCredentials!.findIndex(x => x.keyId === item.keyId!), 1);

            // Set the added trigger to the status bar message.
            const previousIcon = item.iconPath;
            const status = this.triggerTreeChange("Deleting Key Credential", item);
            this.graphClient.deleteKeyCredential(item.objectId!, keyCredentials)
                .then(() => {
                    this.triggerOnComplete({ success: true, statusBarHandle: status });
                })
                .catch((error) => {
                    this.triggerOnError({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon });
                });
        }
    }

    // Gets the key credentials for an application registration.
    private async getKeyCredentials(id: string): Promise<KeyCredential[]> {
        return (await this.treeDataProvider.getApplicationPartial(id, "keyCredentials")).keyCredentials!;
    }
}