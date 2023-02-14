import * as fs from "fs";
import * as forge from "node-forge";
import { v4 as uuidv4 } from "uuid";
import { window, workspace } from "vscode";
import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { ServiceBase } from "./service-base";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { KeyCredential, Application } from "@microsoft/microsoft-graph-types";
import { GraphResult } from "../types/graph-result";

export class KeyCredentialService extends ServiceBase {
	// The constructor for the KeyCredentialsService class.
	constructor(graphRepository: GraphApiRepository, treeDataProvider: AppRegTreeDataProvider) {
		super(graphRepository, treeDataProvider);
	}

	// Adds a new key credential by uploading a certificate.
	async upload(item: AppRegItem): Promise<void> {
		// Prompt the certificate description.
		const description = await window.showInputBox({
			prompt: "Description",
			placeHolder: "Enter a description for this certificate. If empty the Common Name will be used.",
			title: "Upload Certificate",
			ignoreFocusOut: true
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
				Certificate: ["cer", "pem", "crt"]
			}
		});

		// If escape is pressed then return undefined.
		if (file === undefined) {
			return undefined;
		}

		// Determine the file type and handle it accordingly.
		const fileType = file[0].path.split(".").pop();
		let certificate: forge.pki.Certificate;

		if (fileType === "pem") {
			// Read the certificate file as text.
			const certificateFile = await workspace.openTextDocument(file[0].fsPath);
			certificate = forge.pki.certificateFromPem(certificateFile.getText());
		} else {
			// Read the certificate file as binary.
			const buffer = await fs.promises.readFile(file[0].fsPath);
			const asn1 = forge.asn1.fromDer(forge.util.decode64(buffer.toString("base64")));
			certificate = forge.pki.certificateFromAsn1(asn1);
		}
		const pem = forge.pki.certificateToPem(certificate);
		const base64EncodedCertificate = forge.util.encode64(pem);

		// Create the key credential object.
		const keyCredential: KeyCredential = {
			keyId: uuidv4(),
			displayName: description === "" ? `CN=${certificate.subject.getField("CN").value}` : description,
			type: "AsymmetricX509Cert",
			usage: "Verify",
			key: base64EncodedCertificate
		};

		// Get all the key credentials for the application.
		const keyCredentials = await this.getKeyCredentials(item.objectId!);

		// If the array is undefined then it'll be an Azure CLI authentication issue.
		if (keyCredentials === undefined) {
			return;
		}

		// Push the new key credential to the list.
		keyCredentials.push(keyCredential);

		// Set the added trigger to the status bar message.
		const status = this.indicateChange("Adding Certificate Credential...", item);
		await this.updateKeyCredentials(item.objectId!, keyCredentials, status);
	}

	// Deletes a key credential from an application registration.
	async delete(item: AppRegItem): Promise<void> {
		// Prompt the user to confirm the removal.
		const answer = await window.showWarningMessage("Do you want to delete this certificate credential?", "Yes", "No");

		// If the user confirms the removal then remove the key credential.
		if (answer === "Yes") {
			// Get all the key credentials for the application.
			const keyCredentials = await this.getKeyCredentials(item.objectId!);

			// If the array is undefined then it'll be an Azure CLI authentication issue.
			if (keyCredentials === undefined) {
				return;
			}

			// Remove the scope requested.
			keyCredentials!.splice(
				keyCredentials!.findIndex((x) => x.keyId === item.keyId!),
				1
			);

			// Set the added trigger to the status bar message.
			const status = this.indicateChange("Deleting Certificate Credential...", item);
			await this.updateKeyCredentials(item.objectId!, keyCredentials, status);
		}
	}

	// Gets the key credentials for an application registration.
	private async getKeyCredentials(id: string): Promise<KeyCredential[] | undefined> {
		const result: GraphResult<Application> = await this.graphRepository.getApplicationDetailsPartial(id, "keyCredentials");
		if (result.success === true && result.value !== undefined) {
			return result.value.keyCredentials!;
		} else {
			await this.handleError(result.error);
			return undefined;
		}
	}

	// Updates the key credentials.
	private async updateKeyCredentials(id: string, credentials: KeyCredential[], status: string | undefined = undefined): Promise<void> {
		const update: GraphResult<void> = await this.graphRepository.updateKeyCredentials(id, credentials);
		update.success === true ? await this.triggerRefresh(status) : await this.handleError(update.error);
	}
}
