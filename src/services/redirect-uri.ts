import { window } from "vscode";
import { AppRegTreeDataProvider } from "../data/app-reg-tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { ServiceBase } from "./service-base";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { GraphResult } from "../types/graph-result";
import { Application } from "@microsoft/microsoft-graph-types";
import { clearStatusBarMessage, setStatusBarMessage } from "../utils/status-bar";
import { validateRedirectUri } from "../utils/validation";

export class RedirectUriService extends ServiceBase {
	// The constructor for the RedirectUriService class.
	constructor(graphRepository: GraphApiRepository, treeDataProvider: AppRegTreeDataProvider) {
		super(graphRepository, treeDataProvider);
	}

	// Adds a new redirect URI to an application registration.
	async add(item: AppRegItem): Promise<void> {
		// Show that we're doing something
		const check = setStatusBarMessage("Checking Sign In Audience...");

		// Get all existing redirect URIs to check duplicates.
		let allExistingRedirectUris = await this.getAllExistingUris(item);

		// If the array is undefined then it'll be an Azure CLI authentication issue.
		if (allExistingRedirectUris === undefined) {
			return;
		}

		const signInAudience: GraphResult<string> = await this.graphRepository.getSignInAudience(item.objectId!);
		if (signInAudience.success !== true || signInAudience.value === undefined) {
			await this.handleError(signInAudience.error);
			return;
		}

		clearStatusBarMessage(check!);

		switch (signInAudience.value) {
			case "AzureADMyOrg":
			case "AzureADMultipleOrgs":
				if (allExistingRedirectUris.length > 255) {
					window.showWarningMessage("You cannot add any more Redirect URIs. The maximum for this application type is 256.", "OK");
					return;
				}
				break;
			case "AzureADandPersonalMicrosoftAccount":
			case "PersonalMicrosoftAccount":
				if (allExistingRedirectUris.length > 99) {
					window.showWarningMessage("You cannot add any more Redirect URIs. The maximum for this application type is 100.", "OK");
					return;
				}
				break;
			default:
				break;
		}

		// Prompt the user for the new redirect URI.
		const redirectUri = await window.showInputBox({
			placeHolder: "Enter Redirect URI",
			prompt: "Add a new Redirect URI to the application",
			title: "Add Redirect URI",
			ignoreFocusOut: true,
			validateInput: (value) => validateRedirectUri(value, item.contextValue!, allExistingRedirectUris!, false, undefined)
		});

		// If the redirect URI is not empty then add it to the application.
		if (redirectUri !== undefined && redirectUri.length > 0) {
			// Get existing redirect URIs for this section to add new one.
			let existingRedirectUris = await this.getExistingUris(item);

			// If the array is undefined then it'll be an Azure CLI authentication issue.
			if (existingRedirectUris !== undefined) {
				existingRedirectUris.push(redirectUri);
				await this.updateRedirectUris(item, existingRedirectUris);
			}
		}
	}

	// Deletes a redirect URI.
	async delete(item: AppRegItem): Promise<void> {
		// Prompt the user to confirm the deletion.
		const answer = await window.showWarningMessage(`Do you want to delete the Redirect URI ${item.label!}?`, "Yes", "No");

		// If the answer is yes then delete the redirect URI.
		if (answer === "Yes") {
			// Get the parent application so we can read the redirect uris.
			let newArray: string[] = [];
			// Remove the redirect URI from the array.
			switch (item.contextValue) {
				case "WEB-REDIRECT-URI":
					const resultWeb: GraphResult<Application> = await this.graphRepository.getApplicationDetailsPartial(item.objectId!, "web");
					if (resultWeb.success === true && resultWeb.value !== undefined) {
						resultWeb.value.web!.redirectUris!.splice(resultWeb.value.web!.redirectUris!.indexOf(item.label!.toString()), 1);
						newArray = resultWeb.value.web!.redirectUris!;
						break;
					} else {
						await this.handleError(resultWeb.error);
						return;
					}
				case "SPA-REDIRECT-URI":
					const resultSpa: GraphResult<Application> = await this.graphRepository.getApplicationDetailsPartial(item.objectId!, "spa");
					if (resultSpa.success === true && resultSpa.value !== undefined) {
						resultSpa.value.spa!.redirectUris!.splice(resultSpa.value.spa!.redirectUris!.indexOf(item.label!.toString()), 1);
						newArray = resultSpa.value.spa!.redirectUris!;
						break;
					} else {
						await this.handleError(resultSpa.error);
						return;
					}
				case "NATIVE-REDIRECT-URI":
					const resultPublic: GraphResult<Application> = await this.graphRepository.getApplicationDetailsPartial(item.objectId!, "publicClient");
					if (resultPublic.success === true && resultPublic.value !== undefined) {
						resultPublic.value.publicClient!.redirectUris!.splice(resultPublic.value.publicClient!.redirectUris!.indexOf(item.label!.toString()), 1);
						newArray = resultPublic.value.publicClient!.redirectUris!;
						break;
					} else {
						await this.handleError(resultPublic.error);
						return;
					}
				default:
					// Do nothing.
					break;
			}

			// Update the application.
			await this.updateRedirectUris(item, newArray);
		}
	}

	// Deletes all redirect URIs for a type.
	async deleteAll(item: AppRegItem): Promise<void> {
		// Prompt the user to confirm the deletion.
		const answer = await window.showWarningMessage("Do you want to delete all Redirect URIs of this type?", "Yes", "No");

		// If the answer is yes then delete the redirect URIs.
		if (answer === "Yes") {
			// Update the application with an empty array for this type.
			await this.updateRedirectUris(item, []);
		}
	}

	// Edits a redirect URI.
	async edit(item: AppRegItem): Promise<void> {
		// Get the existing redirect URIs.
		let allExistingRedirectUris = await this.getAllExistingUris(item);

		// If the array is undefined then it'll be an Azure CLI authentication issue.
		if (allExistingRedirectUris === undefined) {
			return;
		}

		// Prompt the user for the new redirect URI.
		const redirectUri = await window.showInputBox({
			placeHolder: "Redirect URI",
			prompt: "Provide a new Redirect URI for the application",
			value: item.label!.toString(),
			title: "Edit Redirect URI",
			ignoreFocusOut: true,
			validateInput: (value) => validateRedirectUri(value, item.contextValue!, allExistingRedirectUris!, true, item.label!.toString())
		});

		// If the new application name is not empty then update the application.
		if (redirectUri !== undefined && redirectUri !== item.label!.toString()) {
			// Get existing redirect URIs for this section to add new one.
			let existingRedirectUris = await this.getExistingUris(item);

			// If the array is undefined then it'll be an Azure CLI authentication issue.
			if (existingRedirectUris !== undefined) {
				// Remove the old redirect URI and add the new one.
				existingRedirectUris.splice(existingRedirectUris.indexOf(item.label!.toString()), 1);
				existingRedirectUris.push(redirectUri);

				// Update the application.
				await this.updateRedirectUris(item, existingRedirectUris);
			}
		}
	}

	// Gets the existing redirect URIs for an application.
	private async getExistingUris(item: AppRegItem): Promise<string[] | undefined> {
		let existingRedirectUris: string[] = [];

		switch (item.contextValue) {
			case "WEB-REDIRECT":
			case "WEB-REDIRECT-URI":
				const resultWeb: GraphResult<Application> = await this.graphRepository.getApplicationDetailsPartial(item.objectId!, "web");
				if (resultWeb.success === true && resultWeb.value !== undefined) {
					existingRedirectUris = resultWeb.value.web!.redirectUris!;
					break;
				} else {
					await this.handleError(resultWeb.error);
					return undefined;
				}
			case "SPA-REDIRECT":
			case "SPA-REDIRECT-URI":
				const resultSpa: GraphResult<Application> = await this.graphRepository.getApplicationDetailsPartial(item.objectId!, "spa");
				if (resultSpa.success === true && resultSpa.value !== undefined) {
					existingRedirectUris = resultSpa.value.spa!.redirectUris!;
					break;
				} else {
					await this.handleError(resultSpa.error);
					return undefined;
				}
			case "NATIVE-REDIRECT":
			case "NATIVE-REDIRECT-URI":
				const resultPublic: GraphResult<Application> = await this.graphRepository.getApplicationDetailsPartial(item.objectId!, "publicClient");
				if (resultPublic.success === true && resultPublic.value !== undefined) {
					existingRedirectUris = resultPublic.value.publicClient!.redirectUris!;
					break;
				} else {
					await this.handleError(resultPublic.error);
					return undefined;
				}
			default:
				// Do nothing.
				break;
		}

		return existingRedirectUris;
	}

	// Gets the existing redirect URIs for an application.
	private async getAllExistingUris(item: AppRegItem): Promise<string[] | undefined> {
		const resultWeb: GraphResult<Application> = await this.graphRepository.getApplicationDetailsPartial(item.objectId!, "web,spa,publicClient");
		if (resultWeb.success === true && resultWeb.value !== undefined) {
			let existingRedirectUris: string[] = [];
			resultWeb.value.web!.redirectUris!.map((uri) => {
				existingRedirectUris.push(uri);
			});
			resultWeb.value.spa!.redirectUris!.map((uri) => {
				existingRedirectUris.push(uri);
			});
			resultWeb.value.publicClient!.redirectUris!.map((uri) => {
				existingRedirectUris.push(uri);
			});
			return existingRedirectUris;
		} else {
			await this.handleError(resultWeb.error);
			return undefined;
		}
	}

	// Updates the redirect URIs for an application.
	private async updateRedirectUris(item: AppRegItem, redirectUris: string[]): Promise<void> {
		// Show progress indicator.
		const status = this.indicateChange("Updating Redirect URIs...", item);

		// Determine which section to add the redirect URI to.
		if (item.contextValue! === "WEB-REDIRECT-URI" || item.contextValue! === "WEB-REDIRECT") {
			const update: GraphResult<void> = await this.graphRepository.updateApplication(item.objectId!, { web: { redirectUris: redirectUris } });
			update.success === true ? await this.triggerRefresh(status) : await this.handleError(update.error);
		} else if (item.contextValue! === "SPA-REDIRECT-URI" || item.contextValue! === "SPA-REDIRECT") {
			const update: GraphResult<void> = await this.graphRepository.updateApplication(item.objectId!, { spa: { redirectUris: redirectUris } });
			update.success === true ? await this.triggerRefresh(status) : await this.handleError(update.error);
		} else if (item.contextValue! === "NATIVE-REDIRECT-URI" || item.contextValue! === "NATIVE-REDIRECT") {
			const update: GraphResult<void> = await this.graphRepository.updateApplication(item.objectId!, { publicClient: { redirectUris: redirectUris } });
			update.success === true ? await this.triggerRefresh(status) : await this.handleError(update.error);
		}
	}
}
