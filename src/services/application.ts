import { window, ThemeIcon, env, Uri, TextDocumentContentProvider, EventEmitter, workspace, Disposable, ExtensionContext } from 'vscode';
import { portalAppUri, signInAudienceOptions } from '../constants';
import { GraphClient } from '../clients/graph';
import { AppRegDataProvider } from '../dataProviders/applicationRegistration';
import { AppRegItem } from '../models/appRegItem';
import { convertSignInAudience } from '../utils/signInAudienceUtils';

export class ApplicationService {

    // A private instance of the GraphClient class.
    private graphClient: GraphClient;

    // A private instance of the AppRegDataProvider class.
    private dataProvider: AppRegDataProvider;

    // A private array to store the subscriptions.
    private subscriptions: Disposable[] = [];

    // The constructor for the ApplicationRegistrations class.
    constructor(graphClient: GraphClient, dataProvider: AppRegDataProvider, context: ExtensionContext) {
        this.graphClient = graphClient;
        this.dataProvider = dataProvider;
        this.subscriptions = context.subscriptions;
    }

    // Creates a new application registration.
    public async add(): Promise<boolean> {

        // Set the created trigger default to false.
        let added = false;

        // Prompt the user for the application name.
        const newName = await window.showInputBox({
            placeHolder: "Application name...",
            prompt: "Create new application registration",
        });

        // If the application name is not undefined then prompt the user for the sign in audience.
        if (newName !== undefined) {
            // Prompt the user for the sign in audience.
            const audience = await window.showQuickPick(signInAudienceOptions, {
                placeHolder: "Select the sign in audience...",
            });

            // If the sign in audience is not undefined then create the application.
            if (audience !== undefined) {
                await this.graphClient.createApplication({ displayName: newName, signInAudience: convertSignInAudience(audience) })
                    .then(() => {
                        added = true;
                    }).catch((error) => {
                        console.error(error);
                    });
            }
        }

        // Return the state of the action to refresh the list if required.
        return added;
    };

    // Renames an application registration.
    public async rename(app: AppRegItem): Promise<boolean> {

        // Set the update trigger default to false.
        let updated = false;

        // Prompt the user for the new application name.
        const newName = await window.showInputBox({
            placeHolder: "New application name...",
            prompt: "Rename application with new display name",
            value: app.manifest!.displayName!
        });

        // If the new application name is not undefined then update the application.
        if (newName !== undefined) {
            app.iconPath = new ThemeIcon("loading~spin");
            this.dataProvider.triggerOnDidChangeTreeData();
            await this.graphClient.updateApplication(app.objectId!, { displayName: newName })
                .then(() => {
                    updated = true;
                }).catch((error) => {
                    console.error(error);
                });
        }

        // Return the state of the action to refresh the list if required.
        return updated;
    };

    // Deletes an application registration.
    public async delete(app: AppRegItem): Promise<boolean> {

        // Set the deleted trigger default to false.
        let deleted = false;

        // Prompt the user to confirm the deletion.
        const answer = await window.showInformationMessage(`Do you want to delete the application ${app.label}?`, "Yes", "No");

        // If the user confirms the deletion then delete the application.
        if (answer === "Yes") {
            app.iconPath = new ThemeIcon("loading~spin");
            this.dataProvider.triggerOnDidChangeTreeData();
            await this.graphClient.deleteApplication(app.objectId!)
                .then(() => {
                    deleted = true;
                }).catch((error) => {
                    console.error(error);
                });
        }

        // Return the state of the action to refresh the list if required.
        return deleted;
    };

    // Copies the application Id to the clipboard.
    public copyId(app: AppRegItem): void {
        env.clipboard.writeText(app.appId!);
    };

    // Opens the application registration in the Azure Portal.
    public openInPortal(app: AppRegItem): void {
        env.openExternal(Uri.parse(`${portalAppUri}${app.appId}`));
    }

    // Opens the application manifest in a new editor window.
    public viewManifest(app: AppRegItem): void {
        const newDocument = new class implements TextDocumentContentProvider {
            onDidChangeEmitter = new EventEmitter<Uri>();
            onDidChange = this.onDidChangeEmitter.event;
            provideTextDocumentContent(uri: Uri): string {
                return JSON.stringify(app.manifest, null, 4);
            }
        };
        this.subscriptions.push(workspace.registerTextDocumentContentProvider('manifest', newDocument));
        const uri = Uri.parse('manifest:' + app.label + ".json");
        workspace.openTextDocument(uri)
            .then(doc => window.showTextDocument(
                doc, { preview: false }
            ));
    };
}