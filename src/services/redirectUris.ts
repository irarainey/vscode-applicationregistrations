import { window, ThemeIcon, Disposable } from 'vscode';
import { GraphClient } from '../clients/graph';
import { AppRegDataProvider } from '../data/applicationRegistration';
import { AppRegItem } from '../models/appRegItem';

export class RedirectUriService {

    // A private instance of the GraphClient class.
    private _graphClient: GraphClient;

    // A private instance of the AppRegDataProvider class.
    private _dataProvider: AppRegDataProvider;

    // The constructor for the RedirectUriService class.
    constructor(dataProvider: AppRegDataProvider) {
        this._dataProvider = dataProvider;
        this._graphClient = dataProvider.graphClient;
    }

    // Adds a new redirect URI to an application registration.
    public async add(item: AppRegItem): Promise<Disposable | undefined> {

        let existingRedirectUris: string[] = [];
        if (item.children !== undefined) {
            existingRedirectUris = item.children!.map((child) => {
                return child.label!.toString();
            });
        }

        // Prompt the user for the new redirect URI.
        const redirectUri = await window.showInputBox({
            placeHolder: "Enter redirect URI...",
            prompt: "Add a new redirect URI to the application",
            validateInput: (value) => {
                return this.validateRedirectUri(value, item.contextValue!, existingRedirectUris, false);
            }
        });

        // If the redirect URI is not empty then add it to the application.
        if (redirectUri !== undefined && redirectUri.length > 0) {
            existingRedirectUris.push(redirectUri);
            return await this.update(item, existingRedirectUris);
        } else {
            return undefined;
        }
    }

    // Deletes a redirect URI.
    public async delete(item: AppRegItem): Promise<Disposable | undefined> {

        // Prompt the user to confirm the deletion.
        const answer = await window.showInformationMessage(`Do you want to delete the Redirect URI ${item.label!}?`, "Yes", "No");

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
            return await this.update(item, newArray);
        }
    }

    // Edits a redirect URI.   
    public async edit(item: AppRegItem): Promise<Disposable | undefined> {

        let existingRedirectUris: string[] = [];

        // Get the existing redirect URIs.
        switch (item.contextValue) {
            case "WEB-REDIRECT-URI":
                const webParent = await this._dataProvider.getApplicationPartial(item.objectId!, "web");
                existingRedirectUris = webParent.web!.redirectUris!;
                break;
            case "SPA-REDIRECT-URI":
                const spaParent = await this._dataProvider.getApplicationPartial(item.objectId!, "spa");
                existingRedirectUris = spaParent.spa!.redirectUris!;
                break;
            case "NATIVE-REDIRECT-URI":
                const publicClientParent = await this._dataProvider.getApplicationPartial(item.objectId!, "publicClient");
                existingRedirectUris = publicClientParent.publicClient!.redirectUris!;
                break;
        }

        // Prompt the user for the new redirect URI.
        const redirectUri = await window.showInputBox({
            placeHolder: "New application name...",
            prompt: "Rename application with new display name",
            value: item.label!.toString(),
            validateInput: (value) => {
                return this.validateRedirectUri(value, item.contextValue!, existingRedirectUris, true);
            }
        });

        // If the new application name is not empty then update the application.
        if (redirectUri !== undefined && redirectUri !== item.label!.toString()) {
            // Remove the old redirect URI and add the new one.
            existingRedirectUris.splice(existingRedirectUris.indexOf(item.label!.toString()), 1);
            existingRedirectUris.push(redirectUri);

            // Update the application.
            return await this.update(item, existingRedirectUris);
        }
    }

    // Updates the redirect URIs for an application.
    private async update(item: AppRegItem, redirectUris: string[]): Promise<Disposable | undefined> {

        // Show progress indicator.
        let updated = window.setStatusBarMessage("$(loading~spin) Updating redirect URIs...");
        item.iconPath = new ThemeIcon("loading~spin");
        this._dataProvider.triggerOnDidChangeTreeData();

        // Determine which section to add the redirect URI to.
        if (item.contextValue! === "WEB-REDIRECT-URI" || item.contextValue! === "WEB-REDIRECT") {
            await this._graphClient.updateApplication(item.objectId!, { web: { redirectUris: redirectUris } })
                .catch((error) => {
                    console.error(error);
                });
        }
        else if (item.contextValue! === "SPA-REDIRECT-URI" || item.contextValue! === "SPA-REDIRECT") {
            await this._graphClient.updateApplication(item.objectId!, { spa: { redirectUris: redirectUris } })
                .catch((error) => {
                    console.error(error);
                });
        }
        else if (item.contextValue! === "NATIVE-REDIRECT-URI" || item.contextValue! === "NATIVE-REDIRECT") {
            await this._graphClient.updateApplication(item.objectId!, { publicClient: { redirectUris: redirectUris } })
                .catch((error) => {
                    console.error(error);
                });
        }

        return updated;
    }

    // Validates the redirect URI as per https://learn.microsoft.com/en-us/azure/active-directory/develop/reply-url
    private validateRedirectUri(uri: string, context: string, existingRedirectUris: string[], isEditing: boolean): string | undefined {

        // Check to see if the redirect URI already exists.
        if (existingRedirectUris.includes(uri) && isEditing === false) {
            return "The redirect URI specified already exists.";
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