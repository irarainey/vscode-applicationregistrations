import * as path from "path";
import { SIGNIN_COMMAND_TEXT, APPLICATION_SELECT_PROPERTIES } from "../constants";
import { workspace, window, ThemeIcon, ThemeColor, TreeDataProvider, TreeItem, Event, EventEmitter, ProviderResult, ConfigurationTarget } from "vscode";
import { Application, KeyCredential, PasswordCredential, User, AppRole, RequiredResourceAccess, PermissionScope, ServicePrincipal } from "@microsoft/microsoft-graph-types";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { AppRegItem } from "../models/app-reg-item";
import { sort } from "fast-sort";
import { addDays, format, isAfter } from "date-fns";
import { GraphResult } from "../types/graph-result";
import { clearStatusBarMessage, setStatusBarMessage } from "../utils/status-bar";
import { errorHandler } from "../error-handler";

// Application registration tree data provider for the tree view.
export class AppRegTreeDataProvider implements TreeDataProvider<AppRegItem> {
	// Private instance of the tree data
	private treeData: AppRegItem[] = [];

	// Private property to hold the list filter command
	private filterCommand: string | undefined = undefined;

	// Private property to hold the list filter plain text
	private filterText: string | undefined = undefined;

	// This is the event that is fired when the tree view is refreshed.
	private onDidChangeTreeDataEvent: EventEmitter<AppRegItem | undefined | null | void> = new EventEmitter<AppRegItem | undefined | null | void>();

	//Defines the event that is fired when the tree view is refreshed.
	public readonly onDidChangeTreeData: Event<AppRegItem | undefined | null | void> = this.onDidChangeTreeDataEvent.event;

	// A public get property to get the updating state
	public isUpdating: boolean = false;

	// A public get property to return if the tree is empty.
	public get isTreeEmpty(): boolean {
		// If the tree data is empty then return true.
		if (this.treeData.length === 0) {
			return true;
		}

		// If the tree data contains a single item with the context value of "EMPTY" then return true.
		if (this.treeData.length === 1 && this.treeData[0].contextValue === "EMPTY") {
			return true;
		}

		// Otherwise return false.
		return false;
	}

	// The constructor for the AppRegTreeDataProvider class.
	constructor(public graphRepository: GraphApiRepository) {
		this.graphRepository.initialiseTreeView = async (type: string) => {
			await this.render(undefined, type);
		};
	}

	// Initialises the tree view data based on the type of data to be displayed.
	async render(status: string | undefined = undefined, type: string = "APPLICATIONS"): Promise<void> {
		// Clear the tree data
		this.treeData = [];

		// Add the appropriate tree view item based on the type of data to be displayed.
		switch (type) {
			case "INITIALISING":
				this.treeData.push(
					new AppRegItem({
						label: "Initialising extension",
						context: "INITIALISING",
						iconPath: new ThemeIcon("loading~spin", new ThemeColor("editor.foreground"))
					})
				);
				this.onDidChangeTreeDataEvent.fire(undefined);
				break;
			case "AUTHENTICATING":
				this.treeData.push(
					new AppRegItem({
						label: "Waiting for authentication to complete",
						context: "AUTHENTICATING",
						iconPath: new ThemeIcon("loading~spin", new ThemeColor("editor.foreground"))
					})
				);
				this.onDidChangeTreeDataEvent.fire(undefined);
				break;
			case "EMPTY":
				this.treeData.push(
					new AppRegItem({
						label: "No applications found",
						context: "EMPTY",
						iconPath: new ThemeIcon("info", new ThemeColor("editor.foreground"))
					})
				);
				if (status !== undefined) {
					clearStatusBarMessage(status);
				}
				this.onDidChangeTreeDataEvent.fire(undefined);
				break;
			case "SIGN-IN":
				this.treeData.push(
					new AppRegItem({
						label: SIGNIN_COMMAND_TEXT,
						context: "SIGN-IN",
						iconPath: new ThemeIcon("sign-in", new ThemeColor("editor.foreground")),
						command: {
							command: "appRegistrations.signInToAzure",
							title: SIGNIN_COMMAND_TEXT
						}
					})
				);
				if (status !== undefined) {
					clearStatusBarMessage(status);
				}
				this.onDidChangeTreeDataEvent.fire(undefined);
				break;
			case "AUTHENTICATED":
				this.treeData.push(
					new AppRegItem({
						label: "Initialising extension",
						context: "INITIALISING",
						iconPath: new ThemeIcon("loading~spin", new ThemeColor("editor.foreground"))
					})
				);
				this.onDidChangeTreeDataEvent.fire(undefined);
			case "APPLICATIONS":
				await this.populateTreeData(status);
				break;
		}
	}

	// Filters the tree view.
	async filter() {
		// If the tree is currently updating then we don't want to do anything.
		if (this.isUpdating) {
			return;
		}

		// If the tree is currently empty then we don't want to do anything.
		if (this.isTreeEmpty === true && this.filterText === undefined) {
			return;
		}

		// Determine if eventual consistency is enabled.
		const useEventualConsistency = workspace.getConfiguration("applicationRegistrations").get("useEventualConsistency") as boolean;

		// If eventual consistency is disabled then we cannot apply the filter
		if (useEventualConsistency === false) {
			window.showInformationMessage("The application list cannot be filtered when not using eventual consistency. Please enable this in user settings first.", "OK");
			return;
		}

		// Prompt the user for the filter text.
		const newFilter = await window.showInputBox({
			placeHolder: "Name starts with",
			prompt: "Filter applications by display name",
			value: this.filterText,
			ignoreFocusOut: true
		});

		// Escape has been hit so we don't want to do anything.
		if (newFilter === undefined || (newFilter === "" && newFilter === (this.filterText ?? ""))) {
			return;
		} else if (newFilter === "" && this.filterText !== "") {
			this.filterText = undefined;
			this.filterCommand = undefined;
			await this.render(setStatusBarMessage("Loading Application Registrations..."));
		} else if (newFilter !== "" && newFilter !== this.filterText) {
			// If the filter text is not empty then set the filter command and filter text.
			this.filterText = newFilter!;
			this.filterCommand = `startswith(displayName, \'${newFilter.replace(/'/g, "''")}\')`;
			await this.render(setStatusBarMessage("Filtering Application Registrations..."));
		}
	}

	// Trigger the event to refresh the tree view
	triggerOnDidChangeTreeData(item?: AppRegItem) {
		this.onDidChangeTreeDataEvent.fire(item);
	}

	// Returns a single tree item.
	getTreeItem(element: AppRegItem): TreeItem | Thenable<TreeItem> {
		return element;
	}

	// Get the Application parent item of any given tree item.
	getTreeItemApplicationParent(element: AppRegItem): AppRegItem {
		return this.treeData.find((item) => item.contextValue === "APPLICATION" && item.objectId === element.objectId)!;
	}

	// Returns a tree item specified by the context value
	getTreeItemChildByContext(element: AppRegItem, context: string): AppRegItem | undefined {
		if (element.children !== undefined && element.children.length > 0) {
			for (let i = 0; i < element.children.length; i++) {
				if (element.children[i].contextValue === context) {
					return element.children[i];
				} else if (Array.isArray(element.children[i].children)) {
					const result = this.getTreeItemChildByContext(element.children[i], context);
					if (result !== undefined) {
						return result;
					}
				}
			}
		}
		return undefined;
	}

	// Returns the children for the given element or root (if no element is passed).
	getChildren(element?: AppRegItem | undefined): ProviderResult<AppRegItem[] | undefined> {
		// No element selected so return all top level applications to render static elements
		if (element === undefined) {
			return this.treeData;
		}

		// If an element is selected then return the children for that element
		switch (element.contextValue) {
			case "OWNERS":
				// Return the owners for the application
				return Promise.resolve(this.getApplicationOwners(element));
			case "WEB-REDIRECT":
				// Return the web redirect URIs for the application
				return this.graphRepository.getApplicationDetailsPartial(element.objectId!, "web").then(async (result: GraphResult<Application>) => {
					if (result.success === true && result.value !== undefined) {
						return await this.getApplicationRedirectUris(element, "WEB-REDIRECT-URI", result.value.web!.redirectUris!);
					} else {
						await this.handleError(result.error);
						return undefined;
					}
				});
			case "SPA-REDIRECT":
				// Return the SPA redirect URIs for the application
				return this.graphRepository.getApplicationDetailsPartial(element.objectId!, "spa").then(async (result: GraphResult<Application>) => {
					if (result.success === true && result.value !== undefined) {
						return await this.getApplicationRedirectUris(element, "SPA-REDIRECT-URI", result.value.spa!.redirectUris!);
					} else {
						await this.handleError(result.error);
						return undefined;
					}
				});
			case "NATIVE-REDIRECT":
				// Return the native redirect URIs for the application
				return this.graphRepository.getApplicationDetailsPartial(element.objectId!, "publicClient").then(async (result: GraphResult<Application>) => {
					if (result.success === true && result.value !== undefined) {
						return await this.getApplicationRedirectUris(element, "NATIVE-REDIRECT-URI", result.value.publicClient!.redirectUris!);
					} else {
						await this.handleError(result.error);
						return undefined;
					}
				});
			case "PASSWORD-CREDENTIALS":
				// Return the password credentials for the application
				return this.graphRepository.getApplicationDetailsPartial(element.objectId!, "passwordCredentials").then(async (result: GraphResult<Application>) => {
					if (result.success === true && result.value !== undefined) {
						return await this.getApplicationPasswordCredentials(element, result.value.passwordCredentials!);
					} else {
						await this.handleError(result.error);
						return undefined;
					}
				});
			case "CERTIFICATE-CREDENTIALS":
				// Return the key credentials for the application
				return this.graphRepository.getApplicationDetailsPartial(element.objectId!, "keyCredentials").then(async (result: GraphResult<Application>) => {
					if (result.success === true && result.value !== undefined) {
						return await this.getApplicationKeyCredentials(element, result.value.keyCredentials!);
					} else {
						await this.handleError(result.error);
						return undefined;
					}
				});
			case "API-PERMISSIONS":
				// Return the API permissions for the application
				return this.graphRepository.getApplicationDetailsPartial(element.objectId!, "requiredResourceAccess").then(async (result: GraphResult<Application>) => {
					if (result.success === true && result.value !== undefined) {
						return await this.getApplicationApiPermissions(element, result.value.requiredResourceAccess!);
					} else {
						await this.handleError(result.error);
						return undefined;
					}
				});
			case "EXPOSED-API-PERMISSIONS":
				// Return the exposed API permissions for the application
				return this.graphRepository.getApplicationDetailsPartial(element.objectId!, "api").then(async (result: GraphResult<Application>) => {
					if (result.success === true && result.value !== undefined) {
						return await this.getApplicationExposedApiPermissions(element, result.value.api?.oauth2PermissionScopes!);
					} else {
						await this.handleError(result.error);
						return undefined;
					}
				});
			case "APP-ROLES":
				// Return the app roles for the application
				return this.graphRepository.getApplicationDetailsPartial(element.objectId!, "appRoles").then(async (result: GraphResult<Application>) => {
					if (result.success === true && result.value !== undefined) {
						return await this.getApplicationAppRoles(element, result.value.appRoles!);
					} else {
						await this.handleError(result.error);
						return undefined;
					}
				});
			default:
				// Nothing specific so return the statically defined children
				return element.children;
		}
	}

	// Returns the sign-in audience description
	getSignInAudienceDescription(audience: string): string {
		switch (audience) {
			case "AzureADMyOrg":
				return "Single Tenant";
			case "AzureADMultipleOrgs":
				return "Multiple Tenants";
			case "AzureADandPersonalMicrosoftAccount":
				return "Multiple Tenants and Personal Microsoft Accounts";
			case "PersonalMicrosoftAccount":
				return "Personal Microsoft Accounts";
			default:
				return audience;
		}
	}

	// Populates the tree view data for the application registrations.
	private async populateTreeData(status?: string): Promise<void> {
		// If the tree view is already being updated then return.
		if (this.isUpdating) {
			if (status !== undefined) {
				clearStatusBarMessage(status);
			}
			return;
		}

		try {
			// Set the flag to indicate that the tree view is being updated.
			this.isUpdating = true;

			// Get the configuration settings.
			const useEventualConsistency = workspace.getConfiguration("applicationRegistrations").get("useEventualConsistency") as boolean;
			const showApplicationCountWarning = workspace.getConfiguration("applicationRegistrations").get("showApplicationCountWarning") as boolean;

			// Determine if the warning message should be displayed.
			if (showApplicationCountWarning === true) {
				let totalApplicationCount: number = 0;
				const showOwnedApplicationsOnly = workspace.getConfiguration("applicationRegistrations").get("showOwnedApplicationsOnly") as boolean;

				// Get the total number of applications in the tenant based on the filter settings.
				if (showOwnedApplicationsOnly) {
					const result: GraphResult<number> = await this.graphRepository.getApplicationCountOwned();
					if (result.success === true && result.value !== undefined) {
						totalApplicationCount = result.value;
					} else {
						this.isUpdating = false;
						await this.handleError(result.error);
						return;
					}
				} else {
					const result: GraphResult<number> = await this.graphRepository.getApplicationCountAll();
					if (result.success === true && result.value !== undefined) {
						totalApplicationCount = result.value;
					} else {
						this.isUpdating = false;
						await this.handleError(result.error);
						return;
					}
				}

				// If the total number of applications is less than 200 and eventual consistency is enabled then display a warning message.
				if (totalApplicationCount > 0 && totalApplicationCount <= 200 && useEventualConsistency === true) {
					window.showWarningMessage(`You have enabled eventual consistency for Graph API calls but only have ${totalApplicationCount} applications in your tenant. You would likely benefit from disabling eventual consistency in user settings. Would you like to do this now?`, "Yes", "No", "Disable Warning").then((result) => {
						// If the user selects "Disable Warning" then disable the warning message.
						if (result === "Disable Warning") {
							workspace.getConfiguration("applicationRegistrations").update("showApplicationCountWarning", false, ConfigurationTarget.Global);
							// If the user selects "Yes" then disable eventual consistency.
						} else if (result === "Yes") {
							workspace.getConfiguration("applicationRegistrations").update("useEventualConsistency", false, ConfigurationTarget.Global);
						}
					});
					// If the total number of applications is greater than 200 and eventual consistency is disabled then display a warning message.
				} else if (totalApplicationCount > 200 && useEventualConsistency === false) {
					window.showWarningMessage(`You do not have enabled eventual consistency enabled for Graph API calls and have ${totalApplicationCount} applications in your tenant. You would likely benefit from enabling eventual consistency in user settings. Would you like to do this now?`, "Yes", "No", "Disable Warning").then((result) => {
						// If the user selects "Disable Warning" then disable the warning message.
						if (result === "Disable Warning") {
							workspace.getConfiguration("applicationRegistrations").update("showApplicationCountWarning", false, ConfigurationTarget.Global);
							// If the user selects "Yes" then enable eventual consistency.
						} else if (result === "Yes") {
							workspace.getConfiguration("applicationRegistrations").update("useEventualConsistency", true, ConfigurationTarget.Global);
						}
					});
				}
			}

			// Get the application registrations from the graph client.
			const applicationList = await this.getApplicationList();

			// If the array is undefined then it'll be an Azure CLI authentication issue.
			if (applicationList === undefined) {
				this.isUpdating = false;
				return;
			}

			// If there are no application registrations then display an empty tree view.
			if (applicationList.length === 0) {
				this.render(status, "EMPTY");
				this.isUpdating = false;
				return;
			}

			// Create an array to store the application registrations.
			let allApplications: Application[] = [];

			// If we're not using eventual consistency then we need to sort the application registrations by display name.
			if (useEventualConsistency === false) {
				allApplications = sort(applicationList).asc((a) => a.displayName!.toLowerCase().replace(/[`~!@#$%^&*()_|+\-=?;:'",.<>\{\}\[\]\\\/]/gi, ""));
			} else {
				allApplications = applicationList;
			}

			let unsorted: Promise<AppRegItem | undefined>[] = [];
			const showDeletedApplications = workspace.getConfiguration("applicationRegistrations").get("showDeletedApplications") as boolean;
			if (showDeletedApplications === false) {
				// Iterate through the application registrations and create the tree view items.
				unsorted = allApplications.map(async (application, index) => {
					try {
						// Get the application details.
						const result: GraphResult<Application> = await this.graphRepository.getApplicationDetailsPartial(application.id!, APPLICATION_SELECT_PROPERTIES, true);
						if (result.success === true && result.value !== undefined) {
							const app: Application = result.value;

							// Create the tree view item.
							const appRegItem: AppRegItem = new AppRegItem({
								label: app.displayName!,
								value: app.displayName!,
								context: "APPLICATION",
								iconPath: path.join(__filename, "..", "..", "resources", "icons", "app.svg"),
								baseIcon: path.join(__filename, "..", "..", "resources", "icons", "app.svg"),
								objectId: app.id!,
								appId: app.appId!,
								tooltip: `Name: ${app.displayName!}\nClient Id: ${app.appId!}\nCreated: ${format(new Date(app.createdDateTime!), "yyyy-MM-dd")}${app.notes !== null ? "\nNotes: " + app.notes! : ""}`,
								order: index,
								children: []
							});

							// Application (Client) Id
							appRegItem.children!.push(
								new AppRegItem({
									label: "Client Id",
									context: "APPID-PARENT",
									iconPath: new ThemeIcon("preview"),
									baseIcon: new ThemeIcon("preview"),
									objectId: app.id!,
									tooltip: "The Application (Client) Id is used to identify the application to Azure AD.",
									children: [
										new AppRegItem({
											label: app.appId!,
											value: app.appId!,
											context: "COPY",
											iconPath: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
											baseIcon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
											objectId: app.id!,
											tooltip: "The Application (Client) Id is used to identify the application to Azure AD."
										})
									]
								})
							);

							// Application ID URI
							appRegItem.children!.push(
								new AppRegItem({
									label: "Application Id URI",
									context: "APPID-URI-PARENT",
									objectId: app.id!,
									appId: app.appId!,
									value: app.identifierUris![0] === undefined ? "Not set" : app.identifierUris![0],
									iconPath: new ThemeIcon("globe"),
									baseIcon: new ThemeIcon("globe"),
									tooltip: "The Application Id URI is a globally unique URI used to identify this web API. It is the prefix for scopes and in access tokens, it is the value of the audience claim. Also referred to as an identifier URI.",
									children: [
										new AppRegItem({
											label: app.identifierUris![0] === undefined ? "Not set" : app.identifierUris![0],
											value: app.identifierUris![0] === undefined ? "Not set" : app.identifierUris![0],
											appId: app.appId!,
											objectId: app.id!,
											context: app.identifierUris![0] !== undefined ? "APPID-URI" : "APPID-URI-EMPTY",
											iconPath: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
											baseIcon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
											tooltip: "The Application Id URI, this is set when an application is used as a resource app. The URI acts as the prefix for the scopes you'll reference in your API's code, and must be globally unique."
										})
									]
								})
							);

							// Sign In Audience
							appRegItem.children!.push(
								new AppRegItem({
									label: "Sign In Audience",
									context: "AUDIENCE-PARENT",
									iconPath: new ThemeIcon("account"),
									baseIcon: new ThemeIcon("account"),
									value: app.signInAudience!,
									objectId: app.id!,
									tooltip: "The Sign In Audience determines whether the application can be used by accounts in the same Azure AD tenant or accounts in any Azure AD tenant.",
									children: [
										new AppRegItem({
											label: this.getSignInAudienceDescription(app.signInAudience!),
											context: "AUDIENCE",
											iconPath: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
											baseIcon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
											objectId: app.id!,
											value: app.signInAudience!,
											tooltip: "The Sign In Audience determines whether the application can be used by accounts in the same Azure AD tenant or accounts in any Azure AD tenant."
										})
									]
								})
							);

							if (app.web?.redirectUris?.length !== 0 || app.spa?.redirectUris?.length !== 0) {
								// Logout URL
								appRegItem.children!.push(
									new AppRegItem({
										label: "Front-channel Logout URL",
										context: "LOGOUT-URL-PARENT",
										iconPath: new ThemeIcon("sign-out"),
										baseIcon: new ThemeIcon("sign-out"),
										objectId: app.id!,
										value: app.web!.logoutUrl! === undefined || app.web!.logoutUrl! === "" || app.web!.logoutUrl! === null ? "Not set" : app.web!.logoutUrl!,
										tooltip: "The URL to logout of the application.",
										children: [
											new AppRegItem({
												label: app.web!.logoutUrl! === undefined || app.web!.logoutUrl! === "" || app.web!.logoutUrl! === null ? "Not set" : app.web!.logoutUrl!,
												context: app.web!.logoutUrl! === undefined || app.web!.logoutUrl! === "" || app.web!.logoutUrl! === null ? "LOGOUT-URL-EMPTY" : "LOGOUT-URL",
												iconPath: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
												baseIcon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
												objectId: app.id!,
												value: app.web!.logoutUrl! === undefined || app.web!.logoutUrl! === "" || app.web!.logoutUrl! === null ? "Not set" : app.web!.logoutUrl!,
												tooltip: "The URL to logout of the application."
											})
										]
									})
								);
							}

							// Redirect URIs
							appRegItem.children!.push(
								new AppRegItem({
									label: "Redirect URIs",
									context: "REDIRECT-PARENT",
									iconPath: new ThemeIcon("go-to-file", new ThemeColor("editor.foreground")),
									baseIcon: new ThemeIcon("go-to-file", new ThemeColor("editor.foreground")),
									objectId: app.id!,
									tooltip: "The Redirect URIs are the endpoints where Azure AD will return responses containing tokens or errors.",
									children: [
										new AppRegItem({
											label: "Web",
											context: "WEB-REDIRECT",
											iconPath: new ThemeIcon("globe"),
											baseIcon: new ThemeIcon("globe"),
											objectId: app.id!,
											tooltip: "Redirect URIs for web applications.",
											children: app.web?.redirectUris?.length === 0 ? undefined : []
										}),
										new AppRegItem({
											label: "SPA",
											context: "SPA-REDIRECT",
											iconPath: new ThemeIcon("browser"),
											baseIcon: new ThemeIcon("browser"),
											objectId: app.id!,
											tooltip: "Redirect URIs for single page applications.",
											children: app.spa?.redirectUris?.length === 0 ? undefined : []
										}),
										new AppRegItem({
											label: "Mobile and Desktop",
											context: "NATIVE-REDIRECT",
											iconPath: new ThemeIcon("editor-layout"),
											baseIcon: new ThemeIcon("editor-layout"),
											objectId: app.id!,
											tooltip: "Redirect URIs for mobile and desktop applications.",
											children: app.publicClient?.redirectUris?.length === 0 ? undefined : []
										})
									]
								})
							);

							// Credentials
							appRegItem.children!.push(
								new AppRegItem({
									label: "Credentials",
									context: "PROPERTY-ARRAY",
									objectId: app.id!,
									iconPath: new ThemeIcon("shield", new ThemeColor("editor.foreground")),
									baseIcon: new ThemeIcon("shield", new ThemeColor("editor.foreground")),
									tooltip: "Credentials enable confidential applications to identify themselves to the authentication service when receiving tokens at a web addressable location (using an HTTPS scheme).",
									children: [
										new AppRegItem({
											label: "Client Secrets",
											context: "PASSWORD-CREDENTIALS",
											objectId: app.id!,
											iconPath: new ThemeIcon("key", new ThemeColor("editor.foreground")),
											baseIcon: new ThemeIcon("key", new ThemeColor("editor.foreground")),
											tooltip: "Client secrets are used to authenticate confidential applications to the authentication service when receiving tokens at a web addressable location (using an HTTPS scheme).",
											children: app.passwordCredentials!.length === 0 ? undefined : []
										}),
										new AppRegItem({
											label: "Certificates",
											context: "CERTIFICATE-CREDENTIALS",
											objectId: app.id!,
											iconPath: new ThemeIcon("gist-secret", new ThemeColor("editor.foreground")),
											baseIcon: new ThemeIcon("gist-secret", new ThemeColor("editor.foreground")),
											tooltip: "Certificates are used to authenticate confidential applications to the authentication service when receiving tokens at a web addressable location (using an HTTPS scheme).",
											children: app.keyCredentials!.length === 0 ? undefined : []
										})
									]
								})
							);

							// API Permissions
							appRegItem.children!.push(
								new AppRegItem({
									label: "API Permissions",
									context: "API-PERMISSIONS",
									objectId: app.id!,
									iconPath: new ThemeIcon("checklist", new ThemeColor("editor.foreground")),
									baseIcon: new ThemeIcon("checklist", new ThemeColor("editor.foreground")),
									tooltip: "API permissions define the access that an application has to an API.",
									children: app.requiredResourceAccess?.length === 0 ? undefined : []
								})
							);

							// Exposed API Permissions
							appRegItem.children!.push(
								new AppRegItem({
									label: "Exposed API Permissions",
									context: "EXPOSED-API-PERMISSIONS",
									objectId: app.id!,
									iconPath: new ThemeIcon("list-tree", new ThemeColor("editor.foreground")),
									baseIcon: new ThemeIcon("list-tree", new ThemeColor("editor.foreground")),
									tooltip: "Exposed API permissions define custom scopes to restrict access to data and functionality protected by this API.",
									children: app.api!.oauth2PermissionScopes!.length === 0 ? undefined : []
								})
							);

							// App Roles
							appRegItem.children!.push(
								new AppRegItem({
									label: "App Roles",
									context: "APP-ROLES",
									objectId: app.id!,
									iconPath: new ThemeIcon("note", new ThemeColor("editor.foreground")),
									baseIcon: new ThemeIcon("note", new ThemeColor("editor.foreground")),
									tooltip: "App roles define custom roles that can be assigned to users and groups, or assigned as application-only scopes to a client application.",
									children: app.appRoles!.length === 0 ? undefined : []
								})
							);

							// Owners
							appRegItem.children!.push(
								new AppRegItem({
									label: "Owners",
									context: "OWNERS",
									objectId: app.id!,
									iconPath: new ThemeIcon("organization", new ThemeColor("editor.foreground")),
									baseIcon: new ThemeIcon("organization", new ThemeColor("editor.foreground")),
									tooltip: "Owners are users who can manage the application.",
									children: app.owners!.length === 0 ? undefined : []
								})
							);

							return appRegItem;
						} else {
							this.isUpdating = false;
							await this.handleError(result.error);
							return undefined;
						}
					} catch (error: any) {
						// Need to catch here any attempt to get an application that has been deleted but eventual consistency is still listing it
						if (error.code !== undefined && error.code === "Request_ResourceNotFound") {
							console.log("Application no longer exists.");
						}
						// Return undefined which we will filter out later
						return undefined;
					}
				});
			} else {
				// Iterate through the application registrations and create the tree view items.
				unsorted = allApplications.map(async (application, index) => {
					return new AppRegItem({
						label: application.displayName!,
						value: application.displayName!,
						context: "APPLICATION-DELETED",
						iconPath: path.join(__filename, "..", "..", "resources", "icons", "app-deleted.svg"),
						baseIcon: path.join(__filename, "..", "..", "resources", "icons", "app-deleted.svg"),
						objectId: application.id!,
						appId: application.appId!,
						order: index,
						children: []
					});
				});
			}

			// Sort the applications by name and assign to the class-level array used to render the tree.
			this.treeData = sort((await Promise.all(unsorted.filter((a) => a !== undefined))) as AppRegItem[]).asc(async (a) => a.order);

			// Trigger the event to refresh the tree view
			this.triggerOnDidChangeTreeData();

			// Clear the status bar message
			if (status !== undefined) {
				clearStatusBarMessage(status);
			}

			// Set the flag to indicate that the tree is no longer updating
			this.isUpdating = false;
		} catch (error: any) {
			// Set the flag to indicate that the tree is no longer updating
			this.isUpdating = false;

			if (error.code !== undefined && error.code === "CredentialUnavailableError") {
				// Clear the status bar message
				if (status !== undefined) {
					clearStatusBarMessage(status);
				}

				// // Check to see if the user is signed in and if not then prompt them to sign in
				// this.graphRepository.initialise();
			} else {
				await this.handleError(error);
			}
		}
	}

	// Returns application list depending on the user setting
	private async getApplicationList(): Promise<Application[] | undefined> {
		// If the user has selected to show deleted applications then get the list of deleted applications
		const showDeletedApplications = workspace.getConfiguration("applicationRegistrations").get("showDeletedApplications") as boolean;
		if (showDeletedApplications === true) {
			const result: GraphResult<Application[]> = await this.graphRepository.getApplicationListDeleted(this.filterCommand);
			if (result.success === true && result.value !== undefined) {
				return result.value;
			} else {
				await this.handleError(result.error);
				return undefined;
			}
		}

		// Get the user setting to determine whether to show all applications or just the ones owned by the user
		const showOwnedApplicationsOnly = workspace.getConfiguration("applicationRegistrations").get("showOwnedApplicationsOnly") as boolean;
		if (showOwnedApplicationsOnly === true) {
			const result: GraphResult<Application[]> = await this.graphRepository.getApplicationListOwned(this.filterCommand);
			if (result.success === true && result.value !== undefined) {
				return result.value;
			} else {
				await this.handleError(result.error);
				return undefined;
			}
		} else {
			const result: GraphResult<Application[]> = await this.graphRepository.getApplicationListAll(this.filterCommand);
			if (result.success === true && result.value !== undefined) {
				return result.value;
			} else {
				await this.handleError(result.error);
				return undefined;
			}
		}
	}

	// Returns the application owners for the given application
	private async getApplicationOwners(element: AppRegItem): Promise<AppRegItem[] | undefined> {
		const result: GraphResult<User[]> = await this.graphRepository.getApplicationOwners(element.objectId!);
		if (result.success === true && result.value !== undefined) {
			return result.value.map((owner) => {
				return new AppRegItem({
					label: owner.displayName!,
					context: "OWNER",
					iconPath: new ThemeIcon("person", new ThemeColor("editor.foreground")),
					baseIcon: new ThemeIcon("person", new ThemeColor("editor.foreground")),
					objectId: element.objectId,
					userId: owner.id!,
					children: [
						new AppRegItem({
							label: owner.mail!,
							context: "COPY",
							objectId: element.objectId,
							value: owner.mail!,
							iconPath: new ThemeIcon("mail", new ThemeColor("editor.foreground")),
							baseIcon: new ThemeIcon("mail", new ThemeColor("editor.foreground")),
							tooltip: "Email address of user"
						}),
						new AppRegItem({
							label: owner.userPrincipalName!,
							context: "COPY",
							objectId: element.objectId,
							value: owner.userPrincipalName!,
							iconPath: new ThemeIcon("account", new ThemeColor("editor.foreground")),
							baseIcon: new ThemeIcon("account", new ThemeColor("editor.foreground")),
							tooltip: "User principal name of user"
						})
					]
				});
			});
		} else {
			await this.handleError(result.error);
			return undefined;
		}
	}

	// Returns the redirect URIs for the given application
	private async getApplicationRedirectUris(element: AppRegItem, context: string, redirects: string[]): Promise<AppRegItem[]> {
		return redirects.map((redirectUri) => {
			return new AppRegItem({
				label: redirectUri,
				value: redirectUri,
				context: context,
				iconPath: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
				baseIcon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
				objectId: element.objectId
			});
		});
	}

	// Returns the password credentials for the given application
	private async getApplicationPasswordCredentials(element: AppRegItem, passwords: PasswordCredential[]): Promise<AppRegItem[]> {
		return passwords.map((credential) => {
			const isExpired = isAfter(Date.now(), new Date(credential.endDateTime!));
			const isAboutToExpire = isAfter(addDays(Date.now(), 30), new Date(credential.endDateTime!));
			const iconColour = isExpired ? "list.errorForeground" : isAboutToExpire ? "list.warningForeground" : "editor.foreground";
			return new AppRegItem({
				label: credential.displayName!,
				context: "PASSWORD",
				iconPath: new ThemeIcon("symbol-key", new ThemeColor(iconColour)),
				baseIcon: new ThemeIcon("symbol-key", new ThemeColor(iconColour)),
				value: credential.keyId!,
				objectId: element.objectId,
				tooltip: isExpired === true ? "This credential has expired." : isAboutToExpire === true ? "This credential will expire soon." : "A secret string that the application uses to prove its identity when requesting a token. Also can be referred to as application password.",
				children: [
					new AppRegItem({
						label: `Value: ${credential.hint!}******************`,
						context: "PASSWORD-VALUE",
						iconPath: new ThemeIcon("symbol-field", new ThemeColor(iconColour)),
						baseIcon: new ThemeIcon("symbol-field", new ThemeColor(iconColour)),
						objectId: element.objectId,
						tooltip: "Client secret values cannot be viewed or accessed, except for immediately after creation."
					}),
					new AppRegItem({
						label: `Created: ${format(new Date(credential.startDateTime!), "yyyy-MM-dd")}`,
						context: "PASSWORD-VALUE",
						iconPath: new ThemeIcon("symbol-field", new ThemeColor(iconColour)),
						objectId: element.objectId,
						baseIcon: new ThemeIcon("symbol-field", new ThemeColor(iconColour))
					}),
					new AppRegItem({
						label: `Expires: ${format(new Date(credential.endDateTime!), "yyyy-MM-dd")}`,
						context: "PASSWORD-VALUE",
						iconPath: new ThemeIcon("symbol-field", new ThemeColor(iconColour)),
						objectId: element.objectId,
						baseIcon: new ThemeIcon("symbol-field", new ThemeColor(iconColour))
					})
				]
			});
		});
	}

	// Returns the password credentials for the given application
	private async getApplicationKeyCredentials(element: AppRegItem, keys: KeyCredential[]): Promise<AppRegItem[]> {
		return keys.map((credential) => {
			const isExpired = isAfter(Date.now(), new Date(credential.endDateTime!));
			const isAboutToExpire = isAfter(addDays(Date.now(), 30), new Date(credential.endDateTime!));
			const iconColour = isExpired ? "list.errorForeground" : isAboutToExpire ? "list.warningForeground" : "editor.foreground";
			return new AppRegItem({
				label: credential.displayName!,
				context: "CERTIFICATE",
				iconPath: new ThemeIcon("gist-secret", new ThemeColor(iconColour)),
				baseIcon: new ThemeIcon("gist-secret", new ThemeColor(iconColour)),
				objectId: element.objectId,
				keyId: credential.keyId!,
				tooltip: isExpired === true ? "This credential has expired." : isAboutToExpire === true ? "This credential will expire soon." : "A certificate that the application uses to prove its identity when requesting a token. Also can be referred to as application certificate or public key.",
				children: [
					new AppRegItem({
						label: `Type: ${credential.type!}`,
						context: "CERTIFICATE-VALUE",
						iconPath: new ThemeIcon("symbol-field", new ThemeColor(iconColour)),
						objectId: element.objectId,
						baseIcon: new ThemeIcon("symbol-field", new ThemeColor(iconColour))
					}),
					new AppRegItem({
						label: `Key Identifier: ${credential.customKeyIdentifier!}`,
						context: "COPY",
						value: credential.customKeyIdentifier!,
						iconPath: new ThemeIcon("symbol-field", new ThemeColor(iconColour)),
						objectId: element.objectId,
						baseIcon: new ThemeIcon("symbol-field", new ThemeColor(iconColour))
					}),
					new AppRegItem({
						label: `Created: ${format(new Date(credential.startDateTime!), "yyyy-MM-dd")}`,
						context: "CERTIFICATE-VALUE",
						iconPath: new ThemeIcon("symbol-field", new ThemeColor(iconColour)),
						objectId: element.objectId,
						baseIcon: new ThemeIcon("symbol-field", new ThemeColor(iconColour))
					}),
					new AppRegItem({
						label: `Expires: ${format(new Date(credential.endDateTime!), "yyyy-MM-dd")}`,
						context: "CERTIFICATE-VALUE",
						iconPath: new ThemeIcon("symbol-field", new ThemeColor(iconColour)),
						objectId: element.objectId,
						baseIcon: new ThemeIcon("symbol-field", new ThemeColor(iconColour))
					})
				]
			});
		});
	}

	// Returns the api permissions for the given application
	private async getApplicationApiPermissions(element: AppRegItem, permissions: RequiredResourceAccess[]): Promise<AppRegItem[]> {
		// Iterate through each permission and get the service principal app name
		const applicationNames = permissions.map(async (permission) => {
			const result: GraphResult<ServicePrincipal> = await this.graphRepository.findServicePrincipalByAppId(permission.resourceAppId!);
			if (result.success === true && result.value !== undefined) {
				return new AppRegItem({
					label: result.value.displayName!,
					context: "API-PERMISSIONS-APP",
					iconPath: new ThemeIcon("preview", new ThemeColor("editor.foreground")),
					baseIcon: new ThemeIcon("preview", new ThemeColor("editor.foreground")),
					objectId: element.objectId,
					resourceAppId: permission.resourceAppId,
					children: permission.resourceAccess!.map((resourceAccess) => {
						let scopeLabel = "";
						let tooltip = undefined;
						let scope = undefined;
						if (resourceAccess.type === "Scope") {
							const tmp = result.value!.oauth2PermissionScopes;
							const tmp2 = resourceAccess.id;
							const tmp3 = result.value!.oauth2PermissionScopes!.find((scope) => scope.id === resourceAccess.id);

							scope = result.value!.oauth2PermissionScopes!.find((scope) => scope.id === resourceAccess.id)!.value!;
							scopeLabel = `Delegated: ${scope}`;
							tooltip = result.value!.oauth2PermissionScopes!.find((scope) => scope.id === resourceAccess.id)!.adminConsentDescription;
						} else {
							scope = result.value!.appRoles!.find((scope) => scope.id === resourceAccess.id)!.value!;
							scopeLabel = `Application: ${scope}`;
							tooltip = result.value!.appRoles!.find((scope) => scope.id === resourceAccess.id)!.description;
						}

						return new AppRegItem({
							label: scopeLabel,
							context: "API-PERMISSIONS-SCOPE",
							objectId: element.objectId,
							resourceAppId: permission.resourceAppId,
							resourceScopeId: resourceAccess.id,
							value: scope,
							iconPath: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
							baseIcon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")),
							tooltip: tooltip!
						});
					})
				});
			} else {
				await this.handleError(result.error);
				return { label: "Error" };
			}
		});

		return Promise.all(applicationNames);
	}

	// Returns the exposed api permissions for the given application
	private async getApplicationExposedApiPermissions(element: AppRegItem, scopes: PermissionScope[]): Promise<AppRegItem[]> {
		return scopes.map((scope) => {
			const iconColour = scope.isEnabled! ? "editor.foreground" : "disabledForeground";
			return new AppRegItem({
				label: scope.adminConsentDisplayName!,
				context: `SCOPE-${scope.isEnabled! === true ? "ENABLED" : "DISABLED"}`,
				iconPath: new ThemeIcon("list-tree", new ThemeColor(iconColour)),
				baseIcon: new ThemeIcon("list-tree", new ThemeColor(iconColour)),
				objectId: element.objectId,
				value: scope.id!,
				state: scope.isEnabled!,
				children: [
					new AppRegItem({
						label: `Scope: ${scope.value!}`,
						context: "SCOPE-VALUE",
						objectId: element.objectId,
						value: scope.id!,
						iconPath: new ThemeIcon("symbol-field", new ThemeColor(iconColour)),
						baseIcon: new ThemeIcon("symbol-field", new ThemeColor(iconColour))
					}),
					new AppRegItem({
						label: `Description: ${scope.adminConsentDescription!}`,
						context: "SCOPE-DESCRIPTION",
						objectId: element.objectId,
						value: scope.id!,
						iconPath: new ThemeIcon("symbol-field", new ThemeColor(iconColour)),
						baseIcon: new ThemeIcon("symbol-field", new ThemeColor(iconColour))
					}),
					new AppRegItem({
						label: `Consent: ${scope.type! === "User" ? "Admins and Users" : "Admins Only"}`,
						context: "SCOPE-CONSENT",
						objectId: element.objectId,
						value: scope.id!,
						iconPath: new ThemeIcon("symbol-field", new ThemeColor(iconColour)),
						baseIcon: new ThemeIcon("symbol-field", new ThemeColor(iconColour))
					}),
					new AppRegItem({
						label: `Enabled: ${scope.isEnabled! ? "Yes" : "No"}`,
						context: "SCOPE-STATE",
						objectId: element.objectId,
						value: scope.id!,
						iconPath: new ThemeIcon("symbol-field", new ThemeColor(iconColour)),
						baseIcon: new ThemeIcon("symbol-field", new ThemeColor(iconColour))
					})
				]
			});
		});
	}

	// Returns the app roles for the given application
	private async getApplicationAppRoles(element: AppRegItem, roles: AppRole[]): Promise<AppRegItem[]> {
		return roles.map((role) => {
			const iconColour = role.isEnabled! ? "editor.foreground" : "disabledForeground";
			return new AppRegItem({
				label: role.displayName!,
				context: `ROLE-${role.isEnabled! === true ? "ENABLED" : "DISABLED"}`,
				iconPath: new ThemeIcon("person", new ThemeColor(iconColour)),
				baseIcon: new ThemeIcon("person", new ThemeColor(iconColour)),
				objectId: element.objectId,
				value: role.id!,
				state: role.isEnabled!,
				children: [
					new AppRegItem({
						label: `Value: ${role.value!}`,
						context: "ROLE-VALUE",
						objectId: element.objectId,
						value: role.id!,
						iconPath: new ThemeIcon("symbol-field", new ThemeColor(iconColour)),
						baseIcon: new ThemeIcon("symbol-field", new ThemeColor(iconColour))
					}),
					new AppRegItem({
						label: `Description: ${role.description!}`,
						context: "ROLE-DESCRIPTION",
						objectId: element.objectId,
						value: role.id!,
						iconPath: new ThemeIcon("symbol-field", new ThemeColor(iconColour)),
						baseIcon: new ThemeIcon("symbol-field", new ThemeColor(iconColour))
					}),
					new AppRegItem({
						label: `Allowed: ${role.allowedMemberTypes!.map((type) => (type === "Application" ? "Applications" : "Users/Groups")).join(", ")}`,
						context: "ROLE-ALLOWED",
						objectId: element.objectId,
						value: role.id!,
						iconPath: new ThemeIcon("symbol-field", new ThemeColor(iconColour)),
						baseIcon: new ThemeIcon("symbol-field", new ThemeColor(iconColour))
					}),
					new AppRegItem({
						label: `Enabled: ${role.isEnabled! ? "Yes" : "No"}`,
						context: "ROLE-STATE",
						objectId: element.objectId,
						value: role.id!,
						iconPath: new ThemeIcon("symbol-field", new ThemeColor(iconColour)),
						baseIcon: new ThemeIcon("symbol-field", new ThemeColor(iconColour))
					})
				]
			});
		});
	}

	// Trigger the event to indicate an error
	private async handleError(error?: Error) {
		await errorHandler({ error: error, treeDataProvider: this });
	}

	// Dispose of the event listener
	dispose(): void {
		this.onDidChangeTreeDataEvent.dispose();
	}
}
