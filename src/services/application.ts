/* eslint-disable @typescript-eslint/naming-convention */
import { PORTAL_APP_URI, SIGNIN_AUDIENCE_OPTIONS, BASE_ENDPOINT, CLI_TENANT_CMD } from "../constants";
import { window, env, Uri, TextDocumentContentProvider, EventEmitter, workspace } from "vscode";
import { AppRegTreeDataProvider } from "../data/app-reg-tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { ServiceBase } from "./service-base";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { Application } from "@microsoft/microsoft-graph-types";
import { GraphResult } from "../types/graph-result";
import { clearStatusBarMessage } from "../utils/status-bar";
import { execShellCmd } from "../utils/exec-shell-cmd";
import { v4 as uuidv4 } from "uuid";

export class ApplicationService extends ServiceBase {

    // The constructor for the ApplicationService class.
    constructor(graphRepository: GraphApiRepository, treeDataProvider: AppRegTreeDataProvider) {
        super(graphRepository, treeDataProvider);
    }

    // Creates a new application registration.
    async add(): Promise<void> {

        // Prompt the user for the sign in audience.
        const signInAudience = await window.showQuickPick(
            SIGNIN_AUDIENCE_OPTIONS,
            {
                placeHolder: "Select the sign in audience",
                title: "Add New Application (1/2)",
                ignoreFocusOut: true
            });

        if (signInAudience !== undefined) {
            // Prompt the user for the application name.
            const displayName = await window.showInputBox({
                placeHolder: "Application name",
                prompt: "Create new Application Registration",
                title: "Add New Application (2/2)",
                ignoreFocusOut: true,
                validateInput: (value) => this.validateDisplayName(value, signInAudience.value)
            });

            // If the application name is not undefined then prompt the user for the sign in audience.
            if (displayName !== undefined) {
                // Set the added trigger to the status bar message.
                const status = this.indicateChange("Creating Application Registration...");
                const update: GraphResult<Application> = await this.graphRepository.createApplication({ displayName: displayName, signInAudience: signInAudience.value });
                (update.success === true && update.value !== undefined) ? this.triggerOnComplete(status) : this.triggerOnError(update.error);
            }
        }
    }

    // Edit an application id URI.
    async editAppIdUri(item: AppRegItem): Promise<void> {

        const result: GraphResult<string> = await this.graphRepository.getSignInAudience(item.objectId!);
        if (result.success !== true || result.value === undefined) {
            this.triggerOnError(result.error);
            return;
        }

        // Prompt the user for the new uri.
        const uri = await window.showInputBox({
            placeHolder: "Application ID URI",
            prompt: "Set Application ID URI",
            value: item.value! === "Not set" ? `api://${item.appId!}` : item.value!,
            title: "Edit Application ID URI",
            ignoreFocusOut: true,
            validateInput: (value) => this.validateAppIdUri(value, result.value!)
        });

        // If the new application id uri is not undefined then update the application.
        if (uri !== undefined) {
            const status = this.indicateChange("Setting Application ID URI...", item);
            await this.updateApplication(item.objectId!, { identifierUris: [uri] }, status);
        }
    }

    // Renames an application registration.
    async rename(item: AppRegItem): Promise<void> {

        const result: GraphResult<string> = await this.graphRepository.getSignInAudience(item.objectId!);
        if (result.success !== true || result.value === undefined) {
            this.triggerOnError(result.error);
            return;
        }

        // Prompt the user for the new application name.
        const displayName = await window.showInputBox({
            placeHolder: "Application name",
            prompt: "Rename application with new display name",
            value: item.label?.toString(),
            title: "Rename Application",
            ignoreFocusOut: true,
            validateInput: (value) => this.validateDisplayName(value, result.value!)
        });

        // If the new application name is not undefined then update the application.
        if (displayName !== undefined) {
            const status = this.indicateChange("Renaming Application Registration...", item);
            await this.updateApplication(item.objectId!, { displayName: displayName }, status);
        }
    }

    // Deletes an application registration.
    async delete(item: AppRegItem): Promise<void> {

        // Prompt the user to confirm the deletion.
        const answer = await window.showWarningMessage(`Do you want to delete the Application ${item.label}?`, "Yes", "No");

        // If the user confirms the deletion then delete the application.
        if (answer === "Yes") {
            const status = this.indicateChange("Deleting Application Registration...", item);
            const update: GraphResult<void> = await this.graphRepository.deleteApplication(item.objectId!);
            update.success === true ? this.triggerOnComplete(status) : this.triggerOnError(update.error);
        }
    }

    // Deletes an application registration.
    async removeAppIdUri(item: AppRegItem): Promise<void> {

        // Prompt the user to confirm the deletion.
        const answer = await window.showWarningMessage("Do you want to remove the Application ID URI?", "Yes", "No");

        // If the user confirms the deletion then delete the application.
        if (answer === "Yes") {
            const status = this.indicateChange("Removing Application ID URI...", item);
            await this.updateApplication(item.objectId!, { identifierUris: [] }, status);
        }
    }

    // Shows the application endpoints.
    async showEndpoints(item: AppRegItem): Promise<void> {
        const status = this.indicateChange("Loading Endpoints...");
        const result: GraphResult<string> = await this.graphRepository.getSignInAudience(item.objectId!);
        if (result.success === true && result.value !== undefined) {

            let endpoints: { [key: string]: string } = {};
            // Create the endpoints.
            switch (result.value) {
                case "AzureADMyOrg":
                    await execShellCmd(CLI_TENANT_CMD)
                        .then(async (response) => {
                            response = response.replace(/(\r\n)/gm, "");
                            endpoints = {
                                "OAuth 2.0 Authorization Endpoint": `${BASE_ENDPOINT}${response}/oauth2/v2.0/authorize`,
                                "OAuth 2.0 Token Endpoint": `${BASE_ENDPOINT}${response}/oauth2/v2.0/token`,
                                "OAuth 2.0 Device Authorization Endpoint": `${BASE_ENDPOINT}${response}/oauth2/v2.0/devicecode`,
                                "OAuth 2.0 Token Revocation Endpoint": `${BASE_ENDPOINT}${response}/oauth2/v2.0/logout`,
                                "OpenID Connect Discovery Document": `${BASE_ENDPOINT}${response}/v2.0/.well-known/openid-configuration`,
                                "OpenID Connect Metadata Document": `${BASE_ENDPOINT}${response}/v2.0/.well-known/openid-configuration?p=${item.appId}`,
                                "OpenID Connect Keys Document": `${BASE_ENDPOINT}${response}/discovery/v2.0/keys`
                            };
                        })
                        .catch(async (error) => {
                            this.triggerOnError(error);
                        });
                    break;
                case "AzureADMultipleOrgs":
                    endpoints = {
                        "OAuth 2.0 Authorization Endpoint (Organizations)": `${BASE_ENDPOINT}organizations/oauth2/v2.0/authorize`,
                        "OAuth 2.0 Token Endpoint (Organizations)": `${BASE_ENDPOINT}organizations/oauth2/v2.0/token`,
                        "OAuth 2.0 Device Authorization Endpoint (Organizations)": `${BASE_ENDPOINT}organizations/oauth2/v2.0/devicecode`,
                        "OAuth 2.0 Token Revocation Endpoint (Organizations)": `${BASE_ENDPOINT}organizations/oauth2/v2.0/logout`,
                        "OpenID Connect Discovery Document (Organizations)": `${BASE_ENDPOINT}organizations/v2.0/.well-known/openid-configuration`,
                        "OpenID Connect Metadata Document (Organizations)": `${BASE_ENDPOINT}organizations/v2.0/.well-known/openid-configuration?p=${item.appId}`,
                        "OpenID Connect Keys Document (Organizations)": `${BASE_ENDPOINT}organizations/discovery/v2.0/keys`
                    };
                    break;
                case "AzureADandPersonalMicrosoftAccount":
                    endpoints = {
                        "OAuth 2.0 Authorization Endpoint (Common)": `${BASE_ENDPOINT}common/oauth2/v2.0/authorize`,
                        "OAuth 2.0 Token Endpoint (Common)": `${BASE_ENDPOINT}common/oauth2/v2.0/token`,
                        "OAuth 2.0 Device Authorization Endpoint (Common)": `${BASE_ENDPOINT}common/oauth2/v2.0/devicecode`,
                        "OAuth 2.0 Token Revocation Endpoint (Common)": `${BASE_ENDPOINT}common/oauth2/v2.0/logout`,
                        "OpenID Connect Discovery Document (Common)": `${BASE_ENDPOINT}common/v2.0/.well-known/openid-configuration`,
                        "OpenID Connect Metadata Document (Common)": `${BASE_ENDPOINT}common/v2.0/.well-known/openid-configuration?p=${item.appId}`,
                        "OpenID Connect Keys Document (Common)": `${BASE_ENDPOINT}common/discovery/v2.0/keys`
                    };
                    break;
                case "PersonalMicrosoftAccount":
                    endpoints = {
                        "OAuth 2.0 Authorization Endpoint (Consumers)": `${BASE_ENDPOINT}consumers/oauth2/v2.0/authorize`,
                        "OAuth 2.0 Token Endpoint (Consumers)": `${BASE_ENDPOINT}consumers/oauth2/v2.0/token`,
                        "OAuth 2.0 Device Authorization Endpoint (Consumers)": `${BASE_ENDPOINT}consumers/oauth2/v2.0/devicecode`,
                        "OAuth 2.0 Token Revocation Endpoint (Consumers)": `${BASE_ENDPOINT}consumers/oauth2/v2.0/logout`,
                        "OpenID Connect Discovery Document (Consumers)": `${BASE_ENDPOINT}consumers/v2.0/.well-known/openid-configuration`,
                        "OpenID Connect Metadata Document (Consumers)": `${BASE_ENDPOINT}consumers/v2.0/.well-known/openid-configuration?p=${item.appId}`,
                        "OpenID Connect Keys Document (Consumers)": `${BASE_ENDPOINT}consumers/discovery/v2.0/keys`
                    };
                    break;
                default:
                    endpoints = {};
                    break;
            }

            const newDocument = new class implements TextDocumentContentProvider {
                onDidChangeEmitter = new EventEmitter<Uri>();
                onDidChange = this.onDidChangeEmitter.event;
                provideTextDocumentContent(): string {
                    return JSON.stringify(endpoints, null, 4);
                }
            };

            const contentProvider = uuidv4();
            this.disposable.push(workspace.registerTextDocumentContentProvider(contentProvider, newDocument));
            const uri = Uri.parse(`${contentProvider}:Endpoints - ${item.label}.json`);
            workspace.openTextDocument(uri)
                .then(async (doc) => {
                    await window.showTextDocument(doc, { preview: false });
                    clearStatusBarMessage(status!);
                });
        } else {
            this.triggerOnError(result.error);
        }
    }

    // Opens the application manifest in a new editor window.
    async viewManifest(item: AppRegItem): Promise<void> {
        const status = this.indicateChange("Loading Application Manifest...");
        const result: GraphResult<Application> = await this.graphRepository.getApplicationDetailsFull(item.objectId!);
        if (result.success === true && result.value !== undefined) {
            const newDocument = new class implements TextDocumentContentProvider {
                onDidChangeEmitter = new EventEmitter<Uri>();
                onDidChange = this.onDidChangeEmitter.event;
                provideTextDocumentContent(): string {
                    return JSON.stringify(result.value, null, 4);
                }
            };

            const contentProvider = uuidv4();
            this.disposable.push(workspace.registerTextDocumentContentProvider(contentProvider, newDocument));
            const uri = Uri.parse(`${contentProvider}:Manifest - ${item.label}.json`);
            workspace.openTextDocument(uri)
                .then(async (doc) => {
                    await window.showTextDocument(doc, { preview: false });
                    clearStatusBarMessage(status!);
                });
        } else {
            this.triggerOnError(result.error);
        }
    }

    // Updates the application registration.
    private async updateApplication(id: string, application: Application, status: string | undefined = undefined): Promise<void> {
        const update: GraphResult<void> = await this.graphRepository.updateApplication(id, application);
        update.success === true ? this.triggerOnComplete(status) : this.triggerOnError(update.error);
    }

    // Copies the application Id to the clipboard.
    copyClientId(item: AppRegItem): void {
        env.clipboard.writeText(item.appId!);
    }

    // Opens the application registration in the Azure Portal.
    openInPortal(item: AppRegItem): void {
        env.openExternal(Uri.parse(`${PORTAL_APP_URI}${item.appId}`));
    }

    // Validates the display name of the application.
    private validateDisplayName(displayName: string, signInAudience: string): string | undefined {

        // Check the length of the application name.
        if (displayName.length < 1) {
            return "An application name must be at least one character.";
        }

        switch (signInAudience) {
            case "AzureADMyOrg":
            case "AzureADMultipleOrgs":
                if (displayName.length > 120) {
                    return "An application name cannot be longer than 120 characters.";
                }
                break;
            case "AzureADandPersonalMicrosoftAccount":
            case "PersonalMicrosoftAccount":
                if (displayName.length > 90) {
                    return "An application name cannot be longer than 90 characters.";
                }
                break;
            default:
                break;
        }

        return undefined;
    }

    // Validates the app id URI.
    private validateAppIdUri(uri: string, signInAudience: string): string | undefined {

        if (uri.endsWith("/") === true) {
            return "The Application ID URI cannot end with a trailing slash.";
        }

        switch (signInAudience) {
            case "AzureADMyOrg":
            case "AzureADMultipleOrgs":
                if (uri.includes("://") === false || uri.startsWith("://") === true) {
                    return "The Application ID URI is not valid. It must start with http://, https://, api://, MS-APPX://, or customScheme://.";
                }
                if (uri.includes("*") === true) {
                    return "Wildcards are not supported.";
                }
                if (uri.length > 255) {
                    return "The Application ID URI is not valid. A URI cannot be longer than 255 characters.";
                }
                break;
            case "AzureADandPersonalMicrosoftAccount":
            case "PersonalMicrosoftAccount":
                if (uri.includes("://") === false || (uri.startsWith("api://") === false && uri.startsWith("https://") === false && uri.startsWith("http://") === true)) {
                    return "The Application ID URI is not valid. It must start with http://, https://, or api://.";
                }
                if (uri.includes("?") === true || uri.includes("#") === true || uri.includes("*") === true) {
                    return "Wildcards, fragments, and query strings are not supported.";
                }
                if (uri.length > 120) {
                    return "The Application ID URI is not valid. A URI cannot be longer than 120 characters.";
                }
                break;
            default:
                break;
        }

        return undefined;
    }
}