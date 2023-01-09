import * as forge from 'node-forge';
import { v4 as uuidv4 } from "uuid";
import { window, workspace } from "vscode";
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

        // Prompt the certificate description.
        const description = await window.showInputBox({
            prompt: "Description",
            placeHolder: "Enter a description for this certificate",
            title: "Upload Certificate",
            ignoreFocusOut: true,
        });

        // If escape is pressed then return undefined.
        if (description === undefined) {
            return undefined;
        }

        // Get the certificate file from the user.
        const file = await window.showOpenDialog({
            canSelectFiles: true,
            canSelectFolders: false,
            canSelectMany: false,
            openLabel: "Select Certificate",
            filters: {
                // eslint-disable-next-line @typescript-eslint/naming-convention
                "Certificate": ["cer", "pem", "crt"],
            },
        });

        // If escape is pressed then return undefined.
        if (file === undefined) {
            return undefined;
        }

        // Read the certificate file.
        const certificateFile = await workspace.openTextDocument(file[0]);

        // Get the public key as a base64 encoded string.
        const certificate = forge.pki.certificateFromPem(certificateFile.getText());
        const pem = forge.pki.certificateToPem(certificate);
        const base64EncodedCertificate = forge.util.encode64(pem);

        // Create the key credential object.
        const keyCredential: KeyCredential = {
            keyId: uuidv4(),
            displayName: description,
            type: "AsymmetricX509Cert",
            usage: "Verify",
            key: base64EncodedCertificate
        };

        // Get all the key credentials for the application.
        const keyCredentials = await this.getKeyCredentials(item.objectId!);

        // Push the new key credential to the list.
        keyCredentials.push(keyCredential);

        // Set the added trigger to the status bar message.
        const previousIcon = item.iconPath;
        const status = this.triggerTreeChange("Adding Certificate Credential", item);
        this.graphClient.deleteKeyCredential(item.objectId!, keyCredentials)
            .then(() => {
                this.triggerOnComplete({ success: true, statusBarHandle: status });
            })
            .catch((error) => {
                this.triggerOnError({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon });
            });
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
            const status = this.triggerTreeChange("Deleting Certificate Credential", item);
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