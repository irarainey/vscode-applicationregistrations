import { window } from "vscode";
import { AppRegTreeDataProvider } from "../data/app-reg-tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { ServiceBase } from "./service-base";
import { GraphClient } from "../clients/graph-client";

export class RedirectUriService extends ServiceBase {

    // The constructor for the RedirectUriService class.
    constructor(graphClient: GraphClient, treeDataProvider: AppRegTreeDataProvider) {
        super(graphClient, treeDataProvider);
    }

    // Adds a new redirect URI to an application registration.
    async add(item: AppRegItem): Promise<void> {

        // Get the existing redirect URIs.
        let existingRedirectUris = await this.getExistingUris(item);

        // Prompt the user for the new redirect URI.
        const redirectUri = await window.showInputBox({
            placeHolder: "Enter Redirect URI",
            prompt: "Add a new Redirect URI to the application",
            title: "Add Redirect URI",
            ignoreFocusOut: true,
            validateInput: (value) => this.validateRedirectUri(value, item.contextValue!, existingRedirectUris, false, undefined)
        });

        // If the redirect URI is not empty then add it to the application.
        if (redirectUri !== undefined && redirectUri.length > 0) {
            existingRedirectUris.push(redirectUri);
            await this.update(item, existingRedirectUris);
        }
    }

    // Deletes a redirect URI.
    async delete(item: AppRegItem): Promise<void> {

        // Prompt the user to confirm the deletion.
        const answer = await window.showInformationMessage(`Do you want to delete the Redirect URI ${item.label!}?`, "Yes", "No");

        // If the answer is yes then delete the redirect URI.
        if (answer === "Yes") {
            // Get the parent application so we can read the redirect uris.
            let newArray: string[] = [];
            // Remove the redirect URI from the array.
            switch (item.contextValue) {
                case "WEB-REDIRECT-URI":
                    const webParent = await this.graphClient.getApplicationDetailsPartial(item.objectId!, "web");
                    webParent.web!.redirectUris!.splice(webParent.web!.redirectUris!.indexOf(item.label!.toString()), 1);
                    newArray = webParent.web!.redirectUris!;
                    break;
                case "SPA-REDIRECT-URI":
                    const spaParent = await this.graphClient.getApplicationDetailsPartial(item.objectId!, "spa");
                    spaParent.spa!.redirectUris!.splice(spaParent.spa!.redirectUris!.indexOf(item.label!.toString()), 1);
                    newArray = spaParent.spa!.redirectUris!;
                    break;
                case "NATIVE-REDIRECT-URI":
                    const publicClientParent = await this.graphClient.getApplicationDetailsPartial(item.objectId!, "publicClient");
                    publicClientParent.publicClient!.redirectUris!.splice(publicClientParent.publicClient!.redirectUris!.indexOf(item.label!.toString()), 1);
                    newArray = publicClientParent.publicClient!.redirectUris!;
                    break;
                default:
                    // Do nothing.
                    break;
            }

            // Update the application.
            await this.update(item, newArray);
        }
    }

    // Edits a redirect URI.   
    async edit(item: AppRegItem): Promise<void> {

        // Get the existing redirect URIs.
        let existingRedirectUris = await this.getExistingUris(item);

        // Prompt the user for the new redirect URI.
        const redirectUri = await window.showInputBox({
            placeHolder: "Redirect URI",
            prompt: "Provide a new Redirect URI for the application",
            value: item.label!.toString(),
            title: "Edit Redirect URI",
            ignoreFocusOut: true,
            validateInput: (value) => this.validateRedirectUri(value, item.contextValue!, existingRedirectUris, true, item.label!.toString())
        });

        // If the new application name is not empty then update the application.
        if (redirectUri !== undefined && redirectUri !== item.label!.toString()) {
            // Remove the old redirect URI and add the new one.
            existingRedirectUris.splice(existingRedirectUris.indexOf(item.label!.toString()), 1);
            existingRedirectUris.push(redirectUri);

            // Update the application.
            await this.update(item, existingRedirectUris);
        }
    }

    // Gets the existing redirect URIs for an application.
    private async getExistingUris(item: AppRegItem): Promise<string[]> {

        let existingRedirectUris: string[] = [];

        switch (item.contextValue) {
            case "WEB-REDIRECT":
            case "WEB-REDIRECT-URI":
                const webParent = await this.graphClient.getApplicationDetailsPartial(item.objectId!, "web");
                existingRedirectUris = webParent.web!.redirectUris!;
                break;
            case "SPA-REDIRECT":
            case "SPA-REDIRECT-URI":
                const spaParent = await this.graphClient.getApplicationDetailsPartial(item.objectId!, "spa");
                existingRedirectUris = spaParent.spa!.redirectUris!;
                break;
            case "NATIVE-REDIRECT":
            case "NATIVE-REDIRECT-URI":
                const publicClientParent = await this.graphClient.getApplicationDetailsPartial(item.objectId!, "publicClient");
                existingRedirectUris = publicClientParent.publicClient!.redirectUris!;
                break;
            default:
                // Do nothing.
                break;

        }

        return existingRedirectUris;
    }

    // Updates the redirect URIs for an application.
    private async update(item: AppRegItem, redirectUris: string[]): Promise<void> {

        // Show progress indicator.
        const previousIcon = item.iconPath;
        const status = this.triggerTreeChange("Updating Redirect URIs", item);

        // Determine which section to add the redirect URI to.
        if (item.contextValue! === "WEB-REDIRECT-URI" || item.contextValue! === "WEB-REDIRECT") {
            this.graphClient.updateApplication(item.objectId!, { web: { redirectUris: redirectUris } })
                .then(() => {
                    this.triggerOnComplete({ success: true, statusBarHandle: status });
                })
                .catch((error) => {
                    this.triggerOnError({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon, treeDataProvider: this.treeDataProvider });
                });
        }
        else if (item.contextValue! === "SPA-REDIRECT-URI" || item.contextValue! === "SPA-REDIRECT") {
            this.graphClient.updateApplication(item.objectId!, { spa: { redirectUris: redirectUris } })
                .then(() => {
                    this.triggerOnComplete({ success: true, statusBarHandle: status });
                })
                .catch((error) => {
                    this.triggerOnError({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon, treeDataProvider: this.treeDataProvider });
                });
        }
        else if (item.contextValue! === "NATIVE-REDIRECT-URI" || item.contextValue! === "NATIVE-REDIRECT") {
            this.graphClient.updateApplication(item.objectId!, { publicClient: { redirectUris: redirectUris } })
                .then(() => {
                    this.triggerOnComplete({ success: true, statusBarHandle: status });
                })
                .catch((error) => {
                    this.triggerOnError({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon, treeDataProvider: this.treeDataProvider });
                });
        }
    }

    // Validates the redirect URI as per https://learn.microsoft.com/en-us/azure/active-directory/develop/reply-url
    private validateRedirectUri(uri: string, context: string, existingRedirectUris: string[], isEditing: boolean, oldValue: string | undefined): string | undefined {

        // Check to see if the uri already exists.
        if ((isEditing === true && oldValue !== uri) || isEditing === false) {
            if (existingRedirectUris.includes(uri)) {
                return "The Redirect URI specified already exists.";
            }
        }

        if (context === "WEB-REDIRECT-URI" || context === "WEB-REDIRECT") {
            // Check the redirect URI starts with https://
            if (uri.startsWith("https://") === false && uri.startsWith("http://localhost") === false) {
                return "The Redirect URI is not valid. A Redirect URI must start with https:// unless it is using http://localhost.";
            }
        }
        else if (context === "SPA-REDIRECT-URI" || context === "SPA-REDIRECT" || context === "NATIVE-REDIRECT-URI" || context === "NATIVE-REDIRECT") {
            // Check the redirect URI starts with https:// or http:// or customScheme://
            if (uri.includes("://") === false) {
                return "The Redirect URI is not valid. A Redirect URI must start with https, http, or customScheme://.";
            }
        }

        // Check the length of the redirect URI.
        if (uri.length > 256) {
            return "The Redirect URI is not valid. A Redirect URI cannot be longer than 256 characters.";
        }

        return undefined;
    }
}