import * as path from 'path';
import { window, ThemeIcon, env, Uri, TextDocumentContentProvider, EventEmitter, workspace, Disposable, ExtensionContext } from 'vscode';
import { portalAppUri, signInAudienceOptions } from '../constants';
import { GraphClient } from '../clients/graph';
import { AppRegDataProvider } from '../data/applicationRegistration';
import { AppRegItem } from '../models/appRegItem';
import { convertSignInAudience } from '../utils/signInAudienceUtils';

export class ApplicationService {

    // A private instance of the GraphClient class.
    private _graphClient: GraphClient;

    // A private instance of the AppRegDataProvider class.
    private _dataProvider: AppRegDataProvider;

    // A private array to store the subscriptions.
    private _subscriptions: Disposable[] = [];

    // The constructor for the ApplicationService class.
    constructor(dataProvider: AppRegDataProvider, context: ExtensionContext) {
        this._dataProvider = dataProvider;
        this._graphClient = dataProvider.graphClient;
        this._subscriptions = context.subscriptions;
    }

    // Creates a new application registration.
    public async add(): Promise<Disposable | undefined> {

        // Set the created trigger default to undefined.
        let added = undefined;

        // Prompt the user for the application name.
        const newName = await window.showInputBox({
            placeHolder: "Application name...",
            prompt: "Create new application registration",
            validateInput: (value) => {
                return this.validateDisplayName(value);
            }
        });

        // If the application name is not undefined then prompt the user for the sign in audience.
        if (newName !== undefined) {
            // Prompt the user for the sign in audience.
            const audience = await window.showQuickPick(signInAudienceOptions, {
                placeHolder: "Select the sign in audience...",
            });

            // If the sign in audience is not undefined then create the application.
            if (audience !== undefined) {
                added = window.setStatusBarMessage("$(loading~spin) Creating application registration...");
                const newApp = await this._graphClient.createApplication({ displayName: newName, signInAudience: convertSignInAudience(audience) })
                    .catch((error) => {
                        console.error(error);
                    });
            }
        }

        // Return the state of the action to refresh the list if required.
        return added;
    }

    // Edit an application id URI.
    public async editAppIdUri(app: AppRegItem): Promise<Disposable | undefined> {

        // Set the update trigger default to undefined.
        let updated = undefined;

        // Prompt the user for the new uri.
        const newUri = await window.showInputBox({
            placeHolder: "Application Id URI...",
            prompt: "Set Application Id URI",
            value: app.value! === "Not set" ? `api://${app.appId!}` : app.value!,
            validateInput: (value) => {
                return this.validateAppIdUri(value);
            }
        });

        // If the new application id uri is not undefined then update the application.
        if (newUri !== undefined) {
            updated = window.setStatusBarMessage("$(loading~spin) Setting application id uri...");
            app.iconPath = new ThemeIcon("loading~spin");
            this._dataProvider.triggerOnDidChangeTreeData();
            await this._graphClient.updateApplication(app.objectId!, { identifierUris: [newUri] })
                .catch((error) => {
                    console.error(error);
                });
        }

        // Return the state of the action to refresh the list if required.
        return updated;
    }

    // Renames an application registration.
    public async rename(app: AppRegItem): Promise<Disposable | undefined> {

        // Set the update trigger default to undefined.
        let updated = undefined;

        // Prompt the user for the new application name.
        const newName = await window.showInputBox({
            placeHolder: "New application name...",
            prompt: "Rename application with new display name",
            value: app.label?.toString(),
            validateInput: (value) => {
                return this.validateDisplayName(value);
            }
        });

        // If the new application name is not undefined then update the application.
        if (newName !== undefined) {
            updated = window.setStatusBarMessage("$(loading~spin) Renaming application registration...");
            app.iconPath = new ThemeIcon("loading~spin");
            this._dataProvider.triggerOnDidChangeTreeData();
            await this._graphClient.updateApplication(app.objectId!, { displayName: newName })
                .catch((error) => {
                    console.error(error);
                });
        }

        // Return the state of the action to refresh the list if required.
        return updated;
    }

    // Deletes an application registration.
    public async delete(app: AppRegItem): Promise<Disposable | undefined> {

        // Set the deleted trigger default to undefined.
        let deleted = undefined;

        // Prompt the user to confirm the deletion.
        const answer = await window.showInformationMessage(`Do you want to delete the application ${app.label}?`, "Yes", "No");

        // If the user confirms the deletion then delete the application.
        if (answer === "Yes") {
            deleted = window.setStatusBarMessage("$(loading~spin) Deleting application registration...");
            app.iconPath = new ThemeIcon("loading~spin");
            this._dataProvider.triggerOnDidChangeTreeData();
            await this._graphClient.deleteApplication(app.objectId!)
                .catch((error) => {
                    console.error(error);
                });
        }

        // Return the state of the action to refresh the list if required.
        return deleted;
    }

    // Deletes an application registration.
    public async removeAppIdUri(app: AppRegItem): Promise<Disposable | undefined> {

        // Set the deleted trigger default to undefined.
        let removed = undefined;

        // Prompt the user to confirm the deletion.
        const answer = await window.showInformationMessage("Do you want to remove the application id uri?", "Yes", "No");

        // If the user confirms the deletion then delete the application.
        if (answer === "Yes") {
            removed = window.setStatusBarMessage("$(loading~spin) Removing application id uri...");
            app.iconPath = new ThemeIcon("loading~spin");
            this._dataProvider.triggerOnDidChangeTreeData();
            await this._graphClient.updateApplication(app.objectId!, { identifierUris: [] })
                .catch((error) => {
                    console.error(error);
                });
        }

        // Return the state of the action to refresh the list if required.
        return removed;
    }

    // Copies the application Id to the clipboard.
    public copyClientId(app: AppRegItem): void {
        env.clipboard.writeText(app.appId!);
    }

    // Opens the application registration in the Azure Portal.
    public openInPortal(app: AppRegItem): void {
        env.openExternal(Uri.parse(`${portalAppUri}${app.appId}`));
    }

    // Opens the application manifest in a new editor window.
    public async viewManifest(app: AppRegItem): Promise<void> {

        const status = window.setStatusBarMessage("$(loading~spin) Loading application manifest...");
        app.iconPath = new ThemeIcon("loading~spin");
        this._dataProvider.triggerOnDidChangeTreeData();

        const manifest = await this._graphClient.getApplicationDetailsFull(app.objectId!);

        const newDocument = new class implements TextDocumentContentProvider {
            onDidChangeEmitter = new EventEmitter<Uri>();
            onDidChange = this.onDidChangeEmitter.event;
            provideTextDocumentContent(): string {
                return JSON.stringify(manifest, null, 4);
            }
        };

        this._subscriptions.push(workspace.registerTextDocumentContentProvider('manifest', newDocument));
        const uri = Uri.parse('manifest:' + app.label + ".json");
        workspace.openTextDocument(uri)
            .then(doc => {
                window.showTextDocument(doc, { preview: false });
                app.iconPath = path.join(__filename, "..", "..", "..", "resources", "icons", "app.svg");
                this._dataProvider.triggerOnDidChangeTreeData();
                status.dispose();
            });
    }

    // Validates the display name of the application.
    private validateDisplayName(displayName: string): string | undefined {

        // Check the length of the application name.
        if (displayName.length < 1) {
            return "An application name must be at least one character.";
        }

        if (displayName.length > 120) {
            return "An application name cannot be longer than 120 characters.";
        }

        return undefined;
    }

    // Validates the app id URI.
    private validateAppIdUri(uri: string): string | undefined {

        if (uri.includes("://") === false || uri.startsWith("://") === true) {
            return "The application id URI is not valid. It must start with http://, https://, api://; MS-APPX://, or customScheme://";
        }

        if (uri.endsWith("/") === true) {
            return "The Application Id URI cannot end with a trailing slash.";
        }

        if (uri.length > 120) {
            return "The Application Id URI is not valid. A URI cannot be longer than 120 characters.";
        }

        return undefined;
    }
}