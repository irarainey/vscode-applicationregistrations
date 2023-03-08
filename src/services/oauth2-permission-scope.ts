import { QuickPickItem, window } from "vscode";
import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { PermissionScope, ApiApplication, Application, NullableOption, ServicePrincipal } from "@microsoft/microsoft-graph-types";
import { v4 as uuidv4 } from "uuid";
import { ServiceBase } from "./service-base";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { debounce } from "ts-debounce";
import { GraphResult } from "../types/graph-result";
import { clearStatusBarMessage, setStatusBarMessage } from "../utils/status-bar";
import { validateScopeAdminDescription, validateScopeAdminDisplayName, validateScopeUserDisplayName, validateScopeValue } from "../utils/validation";
import { sort } from "fast-sort";

export class OAuth2PermissionScopeService extends ServiceBase {
	// The constructor for the OAuth2PermissionScopeService class.
	constructor(graphRepository: GraphApiRepository, treeDataProvider: AppRegTreeDataProvider) {
		super(graphRepository, treeDataProvider);
	}

	async addToExistingAuthorisedClient(item: AppRegItem, existingAppId?: NullableOption<string> | undefined, existingAppName?: string): Promise<void> {
		const startStep = existingAppId === undefined ? 0 : 2;
		const numberOfSteps = existingAppId === undefined ? 1 : 3;
		const apiAppName = existingAppId === undefined ? item.label : existingAppName;

		// Update the tree item icon to show the loading animation.
		const status = this.indicateChange(`Loading Scopes for ${apiAppName}...`, item);

		// Determine the application ID to use.
		const appIdToUse = existingAppId === undefined ? item.value : existingAppId;

		// Get the service principal for the application so we can get the scopes.
		const result: GraphResult<ServicePrincipal> = await this.graphRepository.findServicePrincipalByAppId(appIdToUse!);
		if (result.success === false || result.value === undefined) {
			await this.handleError(result.error);
			return;
		}

		const servicePrincipal: ServicePrincipal = result.value;

		// Get all the properties for the application.
		const properties = await this.getProperties(item.objectId!);

		// If the array is undefined then it'll be an Azure CLI authentication issue.
		if (properties === undefined) {
			return;
		}

		// Find the client app in the collection.
		const apiAppScopes = properties.api?.preAuthorizedApplications!.filter((r) => r.appId === item.resourceAppId!)[0];

		// Define a variable to hold the selected scope item.
		let scopeItem: any;

		// Remove any scopes that are already assigned.
		if (apiAppScopes !== undefined) {
			apiAppScopes.delegatedPermissionIds!.forEach((r) => {
				servicePrincipal.oauth2PermissionScopes = servicePrincipal.oauth2PermissionScopes!.filter((s) => s.id !== r);
			});
		}

		clearStatusBarMessage(status!);
		this.resetTreeItemIcon(item);

		// If there are no scopes available then drop out.
		if (servicePrincipal.oauth2PermissionScopes === undefined || servicePrincipal.oauth2PermissionScopes!.length === 0) {
			window.showInformationMessage("There are no scopes available to add to this application.", "OK");
			return;
		}

		// Prompt the user for the scope to add.
		const permissions = sort(servicePrincipal.oauth2PermissionScopes!)
			.asc((r) => r.value!)
			.map((r) => {
				return {
					label: r.value!,
					description: r.adminConsentDescription!,
					value: r.id!
				} as QuickPickItem;
			});
		scopeItem = await window.showQuickPick(permissions, {
			placeHolder: "Select a scope",
			title: `Add Scope (${startStep + 1}/${numberOfSteps})`,
			ignoreFocusOut: true
		});

		// If the user cancels the input then drop out.
		if (scopeItem === undefined) {
			return;
		}

		// Set the added trigger to the status bar message.
		const addStatus = this.indicateChange("Adding Authorized Client Scope...", item);

		if (apiAppScopes !== undefined) {
			// Add the new scope to the existing app.
			apiAppScopes.delegatedPermissionIds!.push(scopeItem.value);
		} else {
			// Add the new scope to a new api app.
			properties.api?.preAuthorizedApplications!.push({
				appId: appIdToUse!,
				delegatedPermissionIds: [ scopeItem.value ]
			});
		}

		//Update the application.
		await this.updateApplication(item.objectId!, { api: properties.api }, addStatus);
	}

	async addAuthorisedClientScope(item: AppRegItem): Promise<void> {
		
	}

	async removeAuthorisedClientScope(item: AppRegItem): Promise<void> {
		// Prompt the user to confirm the removal.
		const answer = await window.showWarningMessage(`Do you want to remove the Authorized Client Scope ${item.label}?`, "Yes", "No");

		// If the user confirms the removal then delete the role.
		if (answer === "Yes") {
			// Set the added trigger to the status bar message.
			const status = this.indicateChange("Removing Authorized Client Scope...", item);

			// Get all the properties for the application.
			const properties = await this.getProperties(item.objectId!);

			// If the array is undefined then it'll be an Azure CLI authentication issue.
			if (properties === undefined) {
				return;
			}

			// Find the client app in the collection.
			const apiAppScopes = properties.api?.preAuthorizedApplications!.filter((r) => r.appId === item.resourceAppId!)[0];

			// Remove the scope requested.
			apiAppScopes?.delegatedPermissionIds!.splice(
				apiAppScopes.delegatedPermissionIds!.findIndex((x) => x === item.value!),
				1
			);

			// If there are no more scopes for the api app then remove the api app from the collection.
			if (apiAppScopes?.delegatedPermissionIds!.length === 0) {
				properties.api?.preAuthorizedApplications!.splice(
					properties.api?.preAuthorizedApplications!.findIndex((r) => r.appId === item.resourceAppId!),
					1
				);
			}

			//Update the application.
			await this.updateApplication(item.objectId!, { api: properties.api }, status);
		}
	}

	async removeAuthorisedClient(item: AppRegItem): Promise<void> {
		// Prompt the user to confirm the removal.
		const answer = await window.showWarningMessage(`Do you want to remove the Authorized Client Application ${item.label}?`, "Yes", "No");

		// If the user confirms the removal then remove the client application.
		if (answer === "Yes") {
			// Set the added trigger to the status bar message.
			const status = this.indicateChange("Removing Authorized Client Application...", item);

			// Get all the properties for the application.
			const properties = await this.getProperties(item.objectId!);

			// If the array is undefined then it'll be an Azure CLI authentication issue.
			if (properties === undefined) {
				return;
			}

			// Find the requested app in the collection and remove it.
			properties.api?.preAuthorizedApplications!.splice(
				properties.api?.preAuthorizedApplications!.findIndex((r) => r.appId === item.resourceAppId!),
				1
			);

			//Update the application.
			await this.updateApplication(item.objectId!, { api: properties.api }, status);
		}
	}

	// Adds a new exposed api scope to an application registration.
	async add(item: AppRegItem): Promise<void> {
		// Show that we're doing something
		const check = setStatusBarMessage("Reading Scopes...");

		// Get the parent application so we can read the required propoerties.
		const properties = await this.getProperties(item.objectId!);

		// If the object is undefined then it'll be an Azure CLI authentication issue.
		if (properties === undefined) {
			return;
		}

		// Check to see if the application has an appIdURI. If it doesn't then we don't want to add a scope.
		const appIdUri = properties.identifierUris;

		if (appIdUri === undefined || appIdUri.length === 0) {
			window.showWarningMessage("This application does not have an Application ID URI. Please add one before adding a scope.", "OK");
			clearStatusBarMessage(check!);
			return;
		}

		// Clear off the status bar message.
		clearStatusBarMessage(check!);

		// Capture the new scope details by passing in an empty scope.
		const scope = await this.inputScopeDetails({}, item.objectId!, false, properties.signInAudience!, properties.api!);

		// If the user cancels the input then return undefined.
		if (scope === undefined) {
			return;
		}

		// Set the added trigger to the status bar message.
		const status = this.indicateChange("Adding Scope...", item);

		// Add the new scope to the existing scopes.
		properties.api!.oauth2PermissionScopes!.push(scope);

		// Update the application.
		await this.updateApplication(item.objectId!, { api: properties.api }, status, "EXPOSED-API-PERMISSIONS");
	}

	// Edits an exposed api scope from an application registration.
	async edit(item: AppRegItem): Promise<void> {
		// Show that we're doing something
		const check = setStatusBarMessage("Reading Scope...");

		// Get the parent application so we can read the required propoerties.
		const properties = await this.getProperties(item.objectId!);

		// If the object is undefined then it'll be an Azure CLI authentication issue.
		if (properties === undefined) {
			return;
		}

		// Check to see if the application has an appIdURI. If it doesn't then we don't want to add a scope.
		const appIdUri = properties.identifierUris;

		if (appIdUri === undefined || appIdUri.length === 0) {
			window.showWarningMessage("This application does not have an Application ID URI. Please add one before editing scopes.", "OK");
			clearStatusBarMessage(check!);
			return;
		}

		// Clear off the status bar message.
		clearStatusBarMessage(check!);

		// Capture the new app role details by passing in the existing role.
		const scope = await this.inputScopeDetails(
			properties.api!.oauth2PermissionScopes!.filter((r) => r.id === item.value!)[0],
			item.objectId!,
			true,
			properties.signInAudience!,
			properties.api!
		);

		// If the user cancels the input then return undefined.
		if (scope === undefined) {
			return;
		}

		// Set the added trigger to the status bar message.
		const status = this.indicateChange("Updating Scope...", item);

		// Update the application.
		await this.updateApplication(item.objectId!, { api: properties.api }, status, "EXPOSED-API-PERMISSIONS");
	}

	// Edits an exposed api scope value from an application registration.
	async editValue(item: AppRegItem): Promise<void> {
		// Show that we're doing something
		const check = setStatusBarMessage("Reading Scope...");

		// Get the parent application so we can read the required propoerties.
		const properties = await this.getProperties(item.objectId!);

		// If the object is undefined then it'll be an Azure CLI authentication issue.
		if (properties === undefined) {
			return;
		}

		// Check to see if the application has an appIdURI. If it doesn't then we don't want to add a scope.
		const appIdUri = properties.identifierUris;

		if (appIdUri === undefined || appIdUri.length === 0) {
			window.showWarningMessage("This application does not have an Application ID URI. Please add one before editing scopes.", "OK");
			clearStatusBarMessage(check!);
			return;
		}

		// Clear off the status bar message.
		clearStatusBarMessage(check!);

		// Get the existing scope.
		const scope = properties.api!.oauth2PermissionScopes!.filter((r) => r.id === item.value!)[0];

		switch (item.contextValue) {
			case "SCOPE-VALUE":
				const value = await this.inputValue(
					"Edit Exposed API Permission (1/1)",
					scope.value!,
					true,
					properties.signInAudience!,
					properties.api!,
					validateScopeValue
				);

				// If escape is pressed or the new name is empty then return undefined.
				if (value === undefined) {
					return undefined;
				}

				scope.value = value;
				break;
			case "SCOPE-NAME":
				// Prompt the user for the new admin consent display name.
				const adminConsentDisplayName = await this.inputAdminConsentDisplayName(
					"Edit Exposed API Permission (1/1)",
					scope.adminConsentDisplayName!,
					validateScopeAdminDisplayName
				);

				// If escape is pressed or the new display name is empty then return undefined.
				if (adminConsentDisplayName === undefined) {
					return undefined;
				}

				scope.adminConsentDisplayName = adminConsentDisplayName;
				break;
			case "SCOPE-DESCRIPTION":
				// Prompt the user for the new admin consent description.
				const adminConsentDescription = await this.inputAdminConsentDescription(
					"Edit Exposed API Permission (1/1)",
					scope.adminConsentDescription!,
					validateScopeAdminDescription
				);

				// If escape is pressed or the new description is empty then return undefined.
				if (adminConsentDescription === undefined) {
					return undefined;
				}

				scope.adminConsentDescription = adminConsentDescription;
				break;
			case "SCOPE-CONSENT":
				// Prompt the user for the new allowed member types.
				const consentType = await this.inputConsentType("Edit Exposed API Permission (1/1)");

				// If escape is pressed or the new allowed member types is empty then return undefined.
				if (consentType === undefined) {
					return undefined;
				}

				scope.type = consentType.value;
				break;
		}

		// Set the added trigger to the status bar message.
		const status = this.indicateChange("Updating Scope...", item);

		// Update the application.
		await this.updateApplication(item.objectId!, { api: properties.api }, status, "EXPOSED-API-PERMISSIONS");
	}

	// Changes the enabled state of an exposed api scope for an application registration.
	async changeState(item: AppRegItem, state: boolean): Promise<void> {
		// Set the added trigger to the status bar message.
		const status = this.indicateChange(state === true ? "Enabling Scope..." : "Disabling Scope...", item);

		// Get the parent application so we can read the scopes.
		const properties = await this.getProperties(item.objectId!);

		// If the array is undefined then it'll be an Azure CLI authentication issue.
		if (properties === undefined) {
			return;
		}

		// Toggle the state of the app role.
		properties.api!.oauth2PermissionScopes!.filter((r) => r.id === item.value!)[0].isEnabled = state;

		// Update the application.
		await this.updateApplication(item.objectId!, { api: properties.api }, status, "EXPOSED-API-PERMISSIONS");
	}

	// Deletes an exposed api scope from an application registration.
	async delete(item: AppRegItem, hideConfirmation: boolean = false): Promise<void> {
		// Get the parent application so we can read the scopes.
		const properties = await this.getProperties(item.objectId!);

		// If the array is undefined then it'll be an Azure CLI authentication issue.
		if (properties === undefined) {
			return;
		}

		if (properties.api?.preAuthorizedApplications!.some((app) => app.delegatedPermissionIds!.includes(item.value!)) === true) {
			await window.showWarningMessage(
				"This scope could not be deleted because it is assigned to an authorized client application. To delete it, please first remove this assignment.",
				"OK"
			);
			return;
		}

		if (item.state !== false && hideConfirmation === false) {
			const disableScope = await window.showWarningMessage(
				`The Scope ${item.label} cannot be deleted unless it is disabled. Do you want to disable the scope and then delete it?`,
				"Yes",
				"No"
			);
			if (disableScope === "Yes") {
				await this.changeState(item, false);
				await this.delete(item, true);
			}
			return;
		}

		if (hideConfirmation === false) {
			// If the user confirms the removal then delete the scope.
			const deleteScope = await window.showWarningMessage(`Do you want to delete the Scope ${item.label}?`, "Yes", "No");
			if (deleteScope === "No") {
				return;
			}
		}

		// Set the added trigger to the status bar message.
		const status = this.indicateChange("Deleting Scope...", item);

		// Remove the app role from the array.
		properties.api!.oauth2PermissionScopes!.splice(
			properties.api!.oauth2PermissionScopes!.findIndex((r) => r.id === item.value!),
			1
		);

		// Update the application.
		await this.updateApplication(item.objectId!, { api: properties.api }, status, "EXPOSED-API-PERMISSIONS");
	}

	// Captures the value for a scope.
	async inputValue(
		title: string,
		existingValue: string | undefined,
		isEditing: boolean,
		signInAudience: string,
		scopes: ApiApplication,
		validate: (
			value: string,
			isEditing: boolean,
			existingValue: string | undefined,
			signInAudience: string,
			scopes: ApiApplication
		) => string | undefined
	): Promise<string | undefined> {
		// Debounce the validation function to prevent multiple calls to the Graph API.
		const debouncedValidation = debounce(validate, 500);

		// Prompt the user for the new value.
		return await window.showInputBox({
			prompt: "Scope value",
			placeHolder: "Enter a scope value (e.g. Files.Read)",
			title: title,
			ignoreFocusOut: true,
			value: existingValue,
			validateInput: async (value) => debouncedValidation(value, isEditing, existingValue, signInAudience, scopes)
		});
	}

	// Captures the admin description for a scope.
	async inputAdminConsentDescription(
		title: string,
		existingValue: string | undefined,
		validate: (description: string) => string | undefined
	): Promise<string | undefined> {
		return await window.showInputBox({
			prompt: "Admin consent description",
			placeHolder: "Enter an admin consent description (e.g. Allows the app to read files on your behalf.)",
			title: title,
			ignoreFocusOut: true,
			value: existingValue,
			validateInput: async (value) => validate(value)
		});
	}

	// Captures the admin display name for a scope.
	async inputAdminConsentDisplayName(
		title: string,
		existingValue: string | undefined,
		validate: (displayName: string) => string | undefined
	): Promise<string | undefined> {
		return await window.showInputBox({
			prompt: "Admin consent display name",
			placeHolder: "Enter an admin consent display name (e.g. Read files)",
			title: title,
			ignoreFocusOut: true,
			value: existingValue,
			validateInput: async (value) => validate(value)
		});
	}

	// Captures the user display name for a scope.
	async inputUserConsentDisplayName(
		title: string,
		existingValue: string | undefined,
		validate: (displayName: string) => string | undefined
	): Promise<string | undefined> {
		return await window.showInputBox({
			prompt: "User consent display name",
			placeHolder: "Enter an optional user consent display name (e.g. Read your files)",
			title: title,
			ignoreFocusOut: true,
			value: existingValue,
			validateInput: async (value) => validate(value)
		});
	}

	// Captures the consent type for a scope.
	async inputConsentType(title: string): Promise<{ label: string; description: string; value: string } | undefined> {
		return await window.showQuickPick(
			[
				{
					label: "Administrators only",
					description: "Only administrators can consent to the scope",
					value: "Admin"
				},
				{
					label: "Administrators and users",
					description: "Administrators and users can consent to the scope",
					value: "User"
				}
			],
			{
				placeHolder: "Select who can consent to the scope",
				title: title,
				ignoreFocusOut: true
			}
		);
	}

	// Captures the details for a scope.
	async inputScopeDetails(
		scope: PermissionScope,
		id: string,
		isEditing: boolean,
		signInAudience: string,
		scopes: ApiApplication
	): Promise<PermissionScope | undefined> {
		const value = await this.inputValue(
			isEditing === true ? "Edit Exposed API Permission (1/7)" : "Add API Exposed Permission (1/7)",
			scope.value ?? undefined,
			isEditing,
			signInAudience,
			scopes,
			validateScopeValue
		);

		// If escape is pressed or the new name is empty then return undefined.
		if (value === undefined) {
			return undefined;
		}

		// Prompt the user for the new allowed member types.
		const consentType = await this.inputConsentType(
			isEditing === true ? "Edit Exposed API Permission (2/7)" : "Add API Exposed Permission (2/7)"
		);

		// If escape is pressed or the new allowed member types is empty then return undefined.
		if (consentType === undefined) {
			return undefined;
		}

		// Prompt the user for the new admin consent display name.
		const adminConsentDisplayName = await this.inputAdminConsentDisplayName(
			isEditing === true ? "Edit Exposed API Permission (3/7)" : "Add API Exposed Permission (3/7)",
			scope.adminConsentDisplayName ?? undefined,
			validateScopeAdminDisplayName
		);

		// If escape is pressed or the new display name is empty then return undefined.
		if (adminConsentDisplayName === undefined) {
			return undefined;
		}

		// Prompt the user for the new admin consent description.
		const adminConsentDescription = await this.inputAdminConsentDescription(
			isEditing === true ? "Edit Exposed API Permission (4/7)" : "Add API Exposed Permission (4/7)",
			scope.adminConsentDescription ?? undefined,
			validateScopeAdminDescription
		);

		// If escape is pressed or the new description is empty then return undefined.
		if (adminConsentDescription === undefined) {
			return undefined;
		}

		// Prompt the user for the new user consent display name.
		const userConsentDisplayName = await this.inputUserConsentDisplayName(
			isEditing === true ? "Edit Exposed API Permission (5/7)" : "Add API Exposed Permission (5/7)",
			scope.userConsentDisplayName ?? undefined,
			validateScopeUserDisplayName
		);

		// If escape is pressed or the new user consent display name is empty then return undefined.
		if (userConsentDisplayName === undefined) {
			return undefined;
		}

		// Prompt the user for the new user consent description.
		const userConsentDescription = await window.showInputBox({
			prompt: "User consent description",
			placeHolder: "Enter an optional user consent description (e.g. Allows the app to read your files.)",
			title: isEditing === true ? "Edit Exposed API Permission (6/7)" : "Add API Exposed Permission (6/7)",
			ignoreFocusOut: true,
			value: scope.userConsentDescription ?? undefined
		});

		// If escape is pressed or the new user consent description is empty then return undefined.
		if (userConsentDescription === undefined) {
			return undefined;
		}

		// Prompt the user for the new state.
		const state = await window.showQuickPick(
			[
				{
					label: "Enabled",
					value: true
				},
				{
					label: "Disabled",
					value: false
				}
			],
			{
				placeHolder: "Select scope state",
				title: isEditing === true ? "Edit Exposed API Permission (7/7)" : "Add API Exposed Permission (7/7)",
				ignoreFocusOut: true
			}
		);

		// If escape is pressed or the new state is empty then return undefined.
		if (state === undefined) {
			return undefined;
		}

		// Set the new values on the scope
		scope.adminConsentDescription = adminConsentDescription;
		scope.adminConsentDisplayName = adminConsentDisplayName;
		scope.id = scope.id ?? uuidv4();
		scope.isEnabled = state.value;
		scope.type = consentType.value;
		scope.userConsentDisplayName = userConsentDisplayName;
		scope.userConsentDescription = userConsentDescription;
		scope.value = value;

		return scope;
	}

	// Gets the required properties from the application registration.
	private async getProperties(id: string): Promise<Application | undefined> {
		const result: GraphResult<Application> = await this.graphRepository.getApplicationDetailsPartial(id, "api,identifierUris,signInAudience");
		if (result.success === true && result.value !== undefined) {
			return result.value;
		} else {
			await this.handleError(result.error);
			return undefined;
		}
	}
}
