import { window, ThemeIcon, Disposable } from 'vscode';
import { GraphClient } from '../clients/graph';
import { AppRegDataProvider } from '../dataProviders/applicationRegistration';
import { AppRegItem } from '../models/appRegItem';

export class RedirectUriService {

    // A private instance of the GraphClient class.
    private graphClient: GraphClient;

    // A private instance of the AppRegDataProvider class.
    private dataProvider: AppRegDataProvider;

    // The constructor for the ApplicationRegistrations class.
    constructor(graphClient: GraphClient, dataProvider: AppRegDataProvider) {
        this.graphClient = graphClient;
        this.dataProvider = dataProvider;
    }

    // Adds a new redirect URI to an application registration.
    public async add(item: AppRegItem): Promise<Disposable | undefined> {

        // Prompt the user for the new redirect URI.
        const redirectUri = await window.showInputBox({
            placeHolder: "Enter redirect URI...",
            prompt: "Add a new redirect URI to the application"
        });

        // If the redirect URI is not empty then add it to the application.
        if (redirectUri !== undefined && redirectUri.length > 0) {

            let existingRedirectUris: string[] = [];
            if(item.children !== undefined) {
                existingRedirectUris = item.children!.map((child) => {
                    return child.label!.toString();
                });
                }

            // Validate the redirect URI.
            if (this.validateRedirectUri(redirectUri, item.contextValue!, existingRedirectUris) === false) {
                return undefined;
            }

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
            // Get the parent application so we can read the manifest.
            const parent = await this.dataProvider.getParentApplication(item.objectId!);
            let newArray: string[] = [];
            // Remove the redirect URI from the array.
            switch (item.contextValue) {
                case "WEB-REDIRECT-URI":
                    parent.web!.redirectUris!.splice(parent.web!.redirectUris!.indexOf(item.label!.toString()), 1);
                    newArray = parent.web!.redirectUris!;
                    break;
                case "SPA-REDIRECT-URI":
                    parent.spa!.redirectUris!.splice(parent.spa!.redirectUris!.indexOf(item.label!.toString()), 1);
                    newArray = parent.spa!.redirectUris!;
                    break;
                case "NATIVE-REDIRECT-URI":
                    parent.publicClient!.redirectUris!.splice(parent.publicClient!.redirectUris!.indexOf(item.label!.toString()), 1);
                    newArray = parent.publicClient!.redirectUris!;
                    break;
            }

            // Update the application.
            return await this.update(item, newArray);
        }
    };

    // Edits a redirect URI.   
    public async edit(item: AppRegItem): Promise<Disposable | undefined> {

        // Prompt the user for the new redirect URI.
        const redirectUri = await window.showInputBox({
            placeHolder: "New application name...",
            prompt: "Rename application with new display name",
            value: item.label!.toString()
        });

        // If the new application name is not empty then update the application.
        if (redirectUri !== undefined && redirectUri !== item.label!.toString()) {

            const parent = await this.dataProvider.getParentApplication(item.objectId!);
            let existingRedirectUris: string[] = [];

            // Get the existing redirect URIs.
            switch (item.contextValue) {
                case "WEB-REDIRECT-URI":
                    existingRedirectUris = parent.web!.redirectUris!;
                    break;
                case "SPA-REDIRECT-URI":
                    existingRedirectUris = parent.spa!.redirectUris!;
                    break;
                case "NATIVE-REDIRECT-URI":
                    existingRedirectUris = parent.publicClient!.redirectUris!;
                    break;
            }

            // Validate the edited redirect URI.
            if (this.validateRedirectUri(redirectUri, item.contextValue!, existingRedirectUris) === false) {
                return undefined;
            }

            // Remove the old redirect URI and add the new one.
            existingRedirectUris.splice(existingRedirectUris.indexOf(item.label!.toString()), 1);
            existingRedirectUris.push(redirectUri);

            // Update the application.
            return await this.update(item, existingRedirectUris);
        }
    };

    // Updates the redirect URIs for an application.
    private async update(item: AppRegItem, redirectUris: string[]): Promise<Disposable | undefined> {

        // Show progress indicator.
        let updated = window.setStatusBarMessage("$(loading~spin) Updating redirect URIs...");
        item.iconPath = new ThemeIcon("loading~spin");
        this.dataProvider.triggerOnDidChangeTreeData();

        // Determine which section to add the redirect URI to.
        if (item.contextValue! === "WEB-REDIRECT-URI" || item.contextValue! === "WEB-REDIRECT") {
            await this.graphClient.updateApplication(item.objectId!, { web: { redirectUris: redirectUris } })
                .catch((error) => {
                    console.error(error);
                });
        }
        else if (item.contextValue! === "SPA-REDIRECT-URI" || item.contextValue! === "SPA-REDIRECT") {
            await this.graphClient.updateApplication(item.objectId!, { spa: { redirectUris: redirectUris } })
                .catch((error) => {
                    console.error(error);
                });
        }
        else if (item.contextValue! === "NATIVE-REDIRECT-URI" || item.contextValue! === "NATIVE-REDIRECT") {
            await this.graphClient.updateApplication(item.objectId!, { publicClient: { redirectUris: redirectUris } })
                .catch((error) => {
                    console.error(error);
                });
        }

        return updated;
    }

    // Validates the redirect URI as per https://learn.microsoft.com/en-us/azure/active-directory/develop/reply-url
    private validateRedirectUri(uri: string, context: string, existingRedirectUris: string[]): boolean {

        // Check to see if the redirect URI already exists.
        if (existingRedirectUris.includes(uri)) {
            window.showErrorMessage("The redirect URI specified already exists.");
            return false;
        }

        if (context === "WEB-REDIRECT-URI" || context === "WEB-REDIRECT") {
            // Check the redirect URI starts with https://
            if (uri.startsWith("https://") === false && uri.startsWith("http://localhost") === false) {
                window.showErrorMessage("The redirect URI is not valid. A redirect URI must start with https:// unless it is using http://localhost.");
                return false;
            }
        }
        else if (context === "SPA-REDIRECT-URI" || context === "SPA-REDIRECT" || context === "NATIVE-REDIRECT-URI" || context === "NATIVE-REDIRECT") {
            // Check the redirect URI starts with https:// or http:// or customScheme://
            if (uri.includes("://") === false) {
                window.showErrorMessage("The redirect URI is not valid. A redirect URI must start with https, http, or customScheme://.");
                return false;
            }
        }

        // Check the length of the redirect URI.
        if (uri.length > 256) {
            window.showErrorMessage("The redirect URI is not valid. A redirect URI cannot be longer than 256 characters.");
            return false;
        }

        return true;
    }
}