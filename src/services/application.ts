/* eslint-disable @typescript-eslint/naming-convention */
import { AZURE_PORTAL_ROOT, ENTRA_PORTAL_ROOT, AZURE_AND_ENTRA_PORTAL_APP_PATH, SIGNIN_AUDIENCE_OPTIONS, BASE_ENDPOINT } from "../constants";
import { window, env, Uri, TextDocumentContentProvider, EventEmitter, workspace } from "vscode";
import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { ServiceBase } from "./service-base";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { Application } from "@microsoft/microsoft-graph-types";
import { GraphResult } from "../types/graph-result";
import { clearStatusBarMessage } from "../utils/status-bar";
import { v4 as uuidv4 } from "uuid";
import { validateAppIdUri, validateApplicationDisplayName, validateLogoutUrl } from "../utils/validation";
import { AccountProvider } from "../types/account-provider";

export class ApplicationService extends ServiceBase {

	// The account provider.
    private accountProvider: AccountProvider;

    // The constructor for the ApplicationService class.
    constructor(graphRepository: GraphApiRepository, treeDataProvider: AppRegTreeDataProvider, accountProvider: AccountProvider) {
        super(graphRepository, treeDataProvider);
        this.accountProvider = accountProvider;
    }

	// Creates a new application registration.
	async add(): Promise<void> {
		// Prompt the user for the sign in audience.
		const signInAudience = await window.showQuickPick(SIGNIN_AUDIENCE_OPTIONS, {
			placeHolder: "Select the sign in audience",
			title: "Add New Application (1/2)",
			ignoreFocusOut: true
		});

		if (signInAudience !== undefined) {
			// Prompt the user for the application name.
			const displayName = await this.inputDisplayNameForNew(signInAudience.value, validateApplicationDisplayName);

			// If the application name is not undefined then create the application.
			if (displayName !== undefined) {
				// Set the added trigger to the status bar message.
				const status = this.indicateChange("Creating Application Registration...");
				const update: GraphResult<Application> = await this.graphRepository.createApplication({ displayName: displayName, signInAudience: signInAudience.value });
				update.success === true && update.value !== undefined ? await this.triggerRefresh(status) : await this.handleError(update.error);
			}
		}
	}

	// Edit an application id URI.
	async editAppIdUri(item: AppRegItem): Promise<void> {
		const result: GraphResult<string> = await this.graphRepository.getSignInAudience(item.objectId!);
		if (result.success !== true || result.value === undefined) {
			await this.handleError(result.error);
			return;
		}

		// Prompt the user for the new uri.
		const uri = await this.inputAppIdUri(item, result.value, validateAppIdUri);

		// If the new application id uri is not undefined then update the application.
		if (uri !== undefined) {
			const status = this.indicateChange("Setting Application ID URI...", item);
			await this.updateApplication(item.objectId!, { identifierUris: [uri] }, status);
		}
	}

	// Edit a front-channel logout url.
	async editLogoutUrl(item: AppRegItem): Promise<void> {
		// Prompt the user for the new uri.
		const uri = await this.inputLogoutUrl(item, validateLogoutUrl);

		// If the new application id uri is not undefined then update the application.
		if (uri !== undefined) {
			const status = this.indicateChange("Setting Front-channel Logout URL...", item);
			await this.updateApplication(item.objectId!, { web: { logoutUrl: uri } }, status);
		}
	}

	// Renames an application registration.
	async rename(item: AppRegItem): Promise<void> {
		const result: GraphResult<string> = await this.graphRepository.getSignInAudience(item.objectId!);
		if (result.success !== true || result.value === undefined) {
			await this.handleError(result.error);
			return;
		}

		// Prompt the user for the new application name.
		const displayName = await this.inputDisplayNameForRename(item.value!, result.value, validateApplicationDisplayName);

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
			update.success === true ? await this.triggerRefresh(status) : await this.handleError(update.error);
		}
	}

	// Removes an App Id Uri.
	async removeAppIdUri(item: AppRegItem): Promise<void> {
		// Prompt the user to confirm the removal.
		const answer = await window.showWarningMessage("Do you want to remove the Application ID URI?", "Yes", "No");

		// If the user confirms the removal then remove it.
		if (answer === "Yes") {
			const status = this.indicateChange("Removing Application ID URI...", item);
			await this.updateApplication(item.objectId!, { identifierUris: [] }, status);
		}
	}

	// Removes a Logout Url
	async removeLogoutUrl(item: AppRegItem): Promise<void> {
		// Prompt the user to confirm the removal.
		const answer = await window.showWarningMessage("Do you want to remove the Front-channel Logout URL?", "Yes", "No");

		// If the user confirms the removal then remove it.
		if (answer === "Yes") {
			const status = this.indicateChange("Removing Front-channel Logout URL...", item);
			await this.updateApplication(item.objectId!, { web: { logoutUrl: null } }, status);
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
                    await this.accountProvider.getAccountInformation()
                        .then(async (accountInformation) => {
                            const response = accountInformation.tenantId.replace(/(\r\n)/gm, "");
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
                            await this.handleError(error);
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

			const newDocument = new (class implements TextDocumentContentProvider {
				onDidChangeEmitter = new EventEmitter<Uri>();
				onDidChange = this.onDidChangeEmitter.event;
				provideTextDocumentContent(): string {
					return JSON.stringify(endpoints, null, 4);
				}
			})();

			const contentProvider = uuidv4();
			this.disposable.push(workspace.registerTextDocumentContentProvider(contentProvider, newDocument));
			const uri = Uri.parse(`${contentProvider}:Endpoints - ${item.label}.json`);
			workspace.openTextDocument(uri).then(async (doc) => {
				await window.showTextDocument(doc, { preview: false });
				clearStatusBarMessage(status!);
			});
		} else {
			await this.handleError(result.error);
		}
	}

	// Opens the application manifest in a new editor window.
	async viewManifest(item: AppRegItem): Promise<void> {
		const status = this.indicateChange("Loading Application Manifest...");
		const result: GraphResult<Application> = await this.graphRepository.getApplicationDetailsFull(item.objectId!);
		if (result.success === true && result.value !== undefined) {
			const newDocument = new (class implements TextDocumentContentProvider {
				onDidChangeEmitter = new EventEmitter<Uri>();
				onDidChange = this.onDidChangeEmitter.event;
				provideTextDocumentContent(): string {
					return JSON.stringify(result.value, null, 4);
				}
			})();

			const contentProvider = uuidv4();
			this.disposable.push(workspace.registerTextDocumentContentProvider(contentProvider, newDocument));
			const uri = Uri.parse(`${contentProvider}:Manifest - ${item.label}.json`);
			workspace.openTextDocument(uri).then(async (doc) => {
				await window.showTextDocument(doc, { preview: false });
				clearStatusBarMessage(status!);
			});
		} else {
			await this.handleError(result.error);
		}
	}

	// Copies the application Id to the clipboard.
	copyClientId(item: AppRegItem): void {
		env.clipboard.writeText(item.appId!);
	}

    // Opens the application registration in the Azure Portal.
    async openInAzurePortal(item: AppRegItem): Promise<boolean> {
        return await this.openInPortal(AZURE_PORTAL_ROOT, item);
    }

    // Opens the application registration in the Entra Portal.
    async openInEntraPortal(item: AppRegItem): Promise<boolean> {
		return await this.openInPortal(ENTRA_PORTAL_ROOT, item);
    }

	// Opens the application registration in the Azure or Entra Portal.
    async openInPortal(portalRoot: string, item: AppRegItem): Promise<boolean>{
        // Determine if "omit tenant ID from portal requests" has been set.
        const omitTenantIdFromPortalRequests = workspace.getConfiguration("applicationRegistrations").get("omitTenantIdFromPortalRequests") as boolean;

        let uriText = "";
        if (omitTenantIdFromPortalRequests === false){
            const accountInformation = await this.accountProvider.getAccountInformation();
            if (accountInformation.tenantId){
                uriText = `${portalRoot}/${accountInformation.tenantId}${AZURE_AND_ENTRA_PORTAL_APP_PATH}${item.appId}`;
            }
        }

        // Check if uriText has been set to anything meaningful yet. If not, then set it w/o a tenant ID.
        if (typeof(uriText) === "string" && uriText.trim().length === 0){
            uriText = `${portalRoot}${AZURE_AND_ENTRA_PORTAL_APP_PATH}${item.appId}`;
        }

        return await env.openExternal(Uri.parse(uriText));
    }

	// Input box for creating a new application registration display name.
	async inputDisplayNameForNew(signInAudience: string, validation: (value: string, signInAudience: string) => string | undefined): Promise<string | undefined> {
		return await window.showInputBox({
			placeHolder: "Application name",
			prompt: "Create new Application Registration",
			title: "Add New Application (2/2)",
			ignoreFocusOut: true,
			validateInput: (value) => validation(value, signInAudience)
		});
	}

	// Input box for renaming an application registration display name.
	async inputDisplayNameForRename(appName: string, signInAudience: string, validation: (value: string, signInAudience: string) => string | undefined): Promise<string | undefined> {
		return await window.showInputBox({
			placeHolder: "Application name",
			prompt: "Rename application with new display name",
			value: appName,
			title: "Rename Application",
			ignoreFocusOut: true,
			validateInput: (value) => validation(value, signInAudience)
		});
	}

	// Input box for editing an application registration App Id URI.
	async inputAppIdUri(item: AppRegItem, signInAudience: string, validation: (value: string, signInAudience: string) => string | undefined): Promise<string | undefined> {
		return await window.showInputBox({
			placeHolder: "Application ID URI",
			prompt: "Set Application ID URI",
			value: item.value! === "Not set" ? `api://${item.appId!}` : item.value!,
			title: "Edit Application ID URI",
			ignoreFocusOut: true,
			validateInput: (value) => validation(value, signInAudience)
		});
	}

	// Input box for editing an application registration Logout URL.
	async inputLogoutUrl(item: AppRegItem, validation: (value: string) => string | undefined): Promise<string | undefined> {
		return await window.showInputBox({
			placeHolder: "Front-channel Logout URL",
			prompt: "Set Front-channel Logout URL",
			value: item.value! === "Not set" ? "" : item.value!,
			title: "Edit Front-channel Logout URL",
			ignoreFocusOut: true,
			validateInput: (value) => validation(value)
		});
	}
}
