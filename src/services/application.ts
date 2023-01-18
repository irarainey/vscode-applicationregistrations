import { PORTAL_APP_URI, SIGNIN_AUDIENCE_OPTIONS } from "../constants";
import { window, env, Uri, TextDocumentContentProvider, EventEmitter, workspace } from "vscode";
import { AppRegTreeDataProvider } from "../data/app-reg-tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { ServiceBase } from "./service-base";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { Application } from "@microsoft/microsoft-graph-types";
import { GraphResult } from "../types/graph-result";

export class ApplicationService extends ServiceBase {

    // The constructor for the ApplicationService class.
    constructor(graphRepository: GraphApiRepository, treeDataProvider: AppRegTreeDataProvider) {
        super(graphRepository, treeDataProvider);
    }

    // Creates a new application registration.
    async add(): Promise<void> {

        if (this.graphRepository.isGraphClientInitialised === false) {
            await this.treeDataProvider.initialiseGraphClient();
            return;
        }

        // Prompt the user for the application name.
        const displayName = await window.showInputBox({
            placeHolder: "Application name",
            prompt: "Create new Application Registration",
            title: "Add New Application (1/2)",
            ignoreFocusOut: true,
            validateInput: (value) => this.validateDisplayName(value)
        });

        // If the application name is not undefined then prompt the user for the sign in audience.
        if (displayName !== undefined) {
            // Prompt the user for the sign in audience.
            const signInAudience = await window.showQuickPick(
                SIGNIN_AUDIENCE_OPTIONS,
                {
                    placeHolder: "Select the sign in audience",
                    title: "Add New Application (2/2)",
                    ignoreFocusOut: true
                });

            // If the sign in audience is not undefined then create the application.
            if (signInAudience !== undefined) {
                // Set the added trigger to the status bar message.
                const status = this.triggerTreeChange("Creating Application Registration");
                this.graphRepository.createApplication({ displayName: displayName, signInAudience: signInAudience.value })
                    .then(() => {
                        this.triggerOnComplete({ success: true, statusBarHandle: status });
                    })
                    .catch((error) => {
                        this.triggerOnError({ success: false, statusBarHandle: status, error: error, treeDataProvider: this.treeDataProvider });
                    });
            }
        }
    }

    // Edit an application id URI.
    async editAppIdUri(item: AppRegItem): Promise<void> {

        // Prompt the user for the new uri.
        const uri = await window.showInputBox({
            placeHolder: "Application ID URI",
            prompt: "Set Application ID URI",
            value: item.value! === "Not set" ? `api://${item.appId!}` : item.value!,
            title: "Edit Application ID URI",
            ignoreFocusOut: true,
            validateInput: (value) => this.validateAppIdUri(value)
        });

        // If the new application id uri is not undefined then update the application.
        if (uri !== undefined) {
            const previousIcon = item.iconPath;
            const status = this.triggerTreeChange("Setting Application ID URI", item);
            this.graphRepository.updateApplication(item.objectId!, { identifierUris: [uri] })
                .then(() => {
                    this.triggerOnComplete({ success: true, statusBarHandle: status });
                })
                .catch((error) => {
                    this.triggerOnError({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon, treeDataProvider: this.treeDataProvider });
                });
        }
    }

    // Renames an application registration.
    async rename(item: AppRegItem): Promise<void> {

        // Prompt the user for the new application name.
        const displayName = await window.showInputBox({
            placeHolder: "Application name",
            prompt: "Rename application with new display name",
            value: item.label?.toString(),
            title: "Rename Application",
            ignoreFocusOut: true,
            validateInput: (value) => this.validateDisplayName(value)
        });

        // If the new application name is not undefined then update the application.
        if (displayName !== undefined) {
            const previousIcon = item.iconPath;
            const status = this.triggerTreeChange("Renaming Application Registration", item);
            this.graphRepository.updateApplication(item.objectId!, { displayName: displayName })
                .then(() => {
                    this.triggerOnComplete({ success: true, statusBarHandle: status });
                })
                .catch((error) => {
                    this.triggerOnError({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon, treeDataProvider: this.treeDataProvider });
                });
        }
    }

    // Deletes an application registration.
    async delete(item: AppRegItem): Promise<void> {

        // Prompt the user to confirm the deletion.
        const answer = await window.showInformationMessage(`Do you want to delete the Application ${item.label}?`, "Yes", "No");

        // If the user confirms the deletion then delete the application.
        if (answer === "Yes") {
            const previousIcon = item.iconPath;
            const status = this.triggerTreeChange("Deleting Application Registration", item);
            this.graphRepository.deleteApplication(item.objectId!)
                .then(() => {
                    this.triggerOnComplete({ success: true, statusBarHandle: status });
                })
                .catch((error) => {
                    this.triggerOnError({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon, treeDataProvider: this.treeDataProvider });
                });
        }
    }

    // Deletes an application registration.
    async removeAppIdUri(item: AppRegItem): Promise<void> {

        // Prompt the user to confirm the deletion.
        const answer = await window.showInformationMessage("Do you want to remove the Application ID URI?", "Yes", "No");

        // If the user confirms the deletion then delete the application.
        if (answer === "Yes") {
            const previousIcon = item.iconPath;
            const status = this.triggerTreeChange("Removing Application ID URI", item);
            this.graphRepository.updateApplication(item.objectId!, { identifierUris: [] })
                .then(() => {
                    this.triggerOnComplete({ success: true, statusBarHandle: status });
                })
                .catch((error) => {
                    this.triggerOnError({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon, treeDataProvider: this.treeDataProvider });
                });
        }
    }

    // Opens the application manifest in a new editor window.
    async viewManifest(item: AppRegItem): Promise<void> {
        const previousIcon = item.iconPath;
        const status = this.triggerTreeChange("Loading Application Manifest", item);
        const result: GraphResult<Application> = await this.graphRepository.getApplicationDetailsFull<Application>(item.objectId!);
        if (result.success === true && result.value !== undefined) {
            const newDocument = new class implements TextDocumentContentProvider {
                onDidChangeEmitter = new EventEmitter<Uri>();
                onDidChange = this.onDidChangeEmitter.event;
                provideTextDocumentContent(): string {
                    return JSON.stringify(result.value, null, 4);
                }
            };
            this.disposable.push(workspace.registerTextDocumentContentProvider("manifest", newDocument));
            const uri = Uri.parse(`manifest:Manifest - ${item.label}.json`);
            workspace.openTextDocument(uri)
                .then(async (doc) => {
                    await window.showTextDocument(doc, { preview: false });
                    status!.dispose();
                    item.iconPath = previousIcon;
                    this.treeDataProvider.triggerOnDidChangeTreeData(item);
                });
        } else {
            this.triggerOnError({ success: false, statusBarHandle: status, error: result.error, treeViewItem: item, previousIcon: previousIcon, treeDataProvider: this.treeDataProvider });
            return;
        }
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
            return "The Application ID URI is not valid. It must start with http://, https://, api://; MS-APPX://, or customScheme://";
        }

        if (uri.endsWith("/") === true) {
            return "The Application ID URI cannot end with a trailing slash.";
        }

        if (uri.length > 120) {
            return "The Application ID URI is not valid. A URI cannot be longer than 120 characters.";
        }

        return undefined;
    }

    // Copies the application Id to the clipboard.
    copyClientId(item: AppRegItem): void {
        env.clipboard.writeText(item.appId!);
    }

    // Opens the application registration in the Azure Portal.
    openInPortal(item: AppRegItem): void {
        env.openExternal(Uri.parse(`${PORTAL_APP_URI}${item.appId}`));
    }
}