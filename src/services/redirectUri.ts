import { window, Disposable } from 'vscode';
import { AppRegDataProvider } from '../data/applicationRegistration';
import { AppRegItem } from '../models/appRegItem';
import { ServiceBase } from './serviceBase';
import { GraphClient } from '../clients/graph';

export class RedirectUriService extends ServiceBase {

    // The constructor for the RedirectUriService class.
    constructor(dataProvider: AppRegDataProvider, graphClient: GraphClient) {
        super(dataProvider, graphClient);
    }

    // Adds a new redirect URI to an application registration.
    public async add(item: AppRegItem): Promise<void> {

        // Get the existing redirect URIs.
        let existingRedirectUris = await this.getExistingUris(item);

        // Prompt the user for the new redirect URI.
        const redirectUri = await window.showInputBox({
            placeHolder: "Enter redirect URI...",
            prompt: "Add a new redirect URI to the application",
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
    public async delete(item: AppRegItem): Promise<void> {

        // Prompt the user to confirm the deletion.
        const answer = await window.showInformationMessage(`Do you want to delete the redirect uri ${item.label!}?`, "Yes", "No");

        // If the answer is yes then delete the redirect URI.
        if (answer === "Yes") {
            // Get the parent application so we can read the redirect uris.
            let newArray: string[] = [];
            // Remove the redirect URI from the array.
            switch (item.contextValue) {
                case "WEB-REDIRECT-URI":
                    const webParent = await this._dataProvider.getApplicationPartial(item.objectId!, "web");
                    webParent.web!.redirectUris!.splice(webParent.web!.redirectUris!.indexOf(item.label!.toString()), 1);
                    newArray = webParent.web!.redirectUris!;
                    break;
                case "SPA-REDIRECT-URI":
                    const spaParent = await this._dataProvider.getApplicationPartial(item.objectId!, "spa");
                    spaParent.spa!.redirectUris!.splice(spaParent.spa!.redirectUris!.indexOf(item.label!.toString()), 1);
                    newArray = spaParent.spa!.redirectUris!;
                    break;
                case "NATIVE-REDIRECT-URI":
                    const publicClientParent = await this._dataProvider.getApplicationPartial(item.objectId!, "publicClient");
                    publicClientParent.publicClient!.redirectUris!.splice(publicClientParent.publicClient!.redirectUris!.indexOf(item.label!.toString()), 1);
                    newArray = publicClientParent.publicClient!.redirectUris!;
                    break;
            }

            // Update the application.
            await this.update(item, newArray);
        }
    }

    // Edits a redirect URI.   
    public async edit(item: AppRegItem): Promise<void> {

        // Get the existing redirect URIs.
        let existingRedirectUris = await this.getExistingUris(item);

        // Prompt the user for the new redirect URI.
        const redirectUri = await window.showInputBox({
            placeHolder: "New application name...",
            prompt: "Rename application with new display name",
            value: item.label!.toString(),
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
                const webParent = await this._dataProvider.getApplicationPartial(item.objectId!, "web");
                existingRedirectUris = webParent.web!.redirectUris!;
                break;
            case "SPA-REDIRECT":
            case "SPA-REDIRECT-URI":
                const spaParent = await this._dataProvider.getApplicationPartial(item.objectId!, "spa");
                existingRedirectUris = spaParent.spa!.redirectUris!;
                break;
            case "NATIVE-REDIRECT":
            case "NATIVE-REDIRECT-URI":
                const publicClientParent = await this._dataProvider.getApplicationPartial(item.objectId!, "publicClient");
                existingRedirectUris = publicClientParent.publicClient!.redirectUris!;
                break;
        }

        return existingRedirectUris;
    }

    // Updates the redirect URIs for an application.
    private async update(item: AppRegItem, redirectUris: string[]): Promise<void> {

        // Show progress indicator.
        const previousIcon = item.iconPath;
        const status = this.indicateChange("Updating redirect uris...", item);

        // Determine which section to add the redirect URI to.
        if (item.contextValue! === "WEB-REDIRECT-URI" || item.contextValue! === "WEB-REDIRECT") {
            this._graphClient.updateApplication(item.objectId!, { web: { redirectUris: redirectUris } })
                .then(() => {
                    this._onComplete.fire({ success: true, statusBarHandle: status });
                })
                .catch((error) => {
                    this._onError.fire({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon });
                });
        }
        else if (item.contextValue! === "SPA-REDIRECT-URI" || item.contextValue! === "SPA-REDIRECT") {
            this._graphClient.updateApplication(item.objectId!, { spa: { redirectUris: redirectUris } })
                .then(() => {
                    this._onComplete.fire({ success: true, statusBarHandle: status });
                })
                .catch((error) => {
                    this._onError.fire({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon });
                });
        }
        else if (item.contextValue! === "NATIVE-REDIRECT-URI" || item.contextValue! === "NATIVE-REDIRECT") {
            this._graphClient.updateApplication(item.objectId!, { publicClient: { redirectUris: redirectUris } })
                .then(() => {
                    this._onComplete.fire({ success: true, statusBarHandle: status });
                })
                .catch((error) => {
                    this._onError.fire({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon });
                });
        }
    }

    // Validates the redirect URI as per https://learn.microsoft.com/en-us/azure/active-directory/develop/reply-url
    private validateRedirectUri(uri: string, context: string, existingRedirectUris: string[], isEditing: boolean, oldValue: string | undefined): string | undefined {

        // Check to see if the uri already exists.
        if ((isEditing === true && oldValue !== uri) || isEditing === false) {
            if (existingRedirectUris.includes(uri)) {
                return "The redirect URI specified already exists.";
            }
        }

        if (context === "WEB-REDIRECT-URI" || context === "WEB-REDIRECT") {
            // Check the redirect URI starts with https://
            if (uri.startsWith("https://") === false && uri.startsWith("http://localhost") === false) {
                return "The redirect URI is not valid. A redirect URI must start with https:// unless it is using http://localhost.";
            }
        }
        else if (context === "SPA-REDIRECT-URI" || context === "SPA-REDIRECT" || context === "NATIVE-REDIRECT-URI" || context === "NATIVE-REDIRECT") {
            // Check the redirect URI starts with https:// or http:// or customScheme://
            if (uri.includes("://") === false) {
                return "The redirect URI is not valid. A redirect URI must start with https, http, or customScheme://.";
            }
        }

        // Check the length of the redirect URI.
        if (uri.length > 256) {
            return "The redirect URI is not valid. A redirect URI cannot be longer than 256 characters.";
        }

        return undefined;
    }
}