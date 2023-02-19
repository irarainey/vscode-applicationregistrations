import { window } from "vscode";
import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { PermissionScope, ApiApplication, Application } from "@microsoft/microsoft-graph-types";
import { v4 as uuidv4 } from "uuid";
import { ServiceBase } from "./service-base";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { debounce } from "ts-debounce";
import { GraphResult } from "../types/graph-result";
import { clearStatusBarMessage, setStatusBarMessage } from "../utils/status-bar";
import { validateScopeAdminDescription, validateScopeAdminDisplayName, validateScopeUserDisplayName, validateScopeValue } from "../utils/validation";

export class OAuth2PermissionScopeService extends ServiceBase {
	// The constructor for the OAuth2PermissionScopeService class.
	constructor(graphRepository: GraphApiRepository, treeDataProvider: AppRegTreeDataProvider) {
		super(graphRepository, treeDataProvider);
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
		const scope = await this.inputScopeDetails({}, item.objectId!, false, properties.signInAudience!, properties.api!, validateScopeUserDisplayName);

		// If the user cancels the input then return undefined.
		if (scope === undefined) {
			return;
		}

		// Set the added trigger to the status bar message.
		const status = this.indicateChange("Adding Scope...", item);

		// Add the new scope to the existing scopes.
		properties.api!.oauth2PermissionScopes!.push(scope);

		// Update the application.
		await this.updateApplication(item.objectId!, { api: properties.api }, status);
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
		const scope = await this.inputScopeDetails(properties.api!.oauth2PermissionScopes!.filter((r) => r.id === item.value!)[0], item.objectId!, true, properties.signInAudience!, properties.api!, validateScopeUserDisplayName);

		// If the user cancels the input then return undefined.
		if (scope === undefined) {
			return;
		}

		// Set the added trigger to the status bar message.
		const status = this.indicateChange("Updating Scope...", item);

		// Update the application.
		await this.updateApplication(item.objectId!, { api: properties.api }, status);
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

		// Get the existsing scope.
		const scope = properties.api!.oauth2PermissionScopes!.filter((r) => r.id === item.value!)[0];

		switch (item.contextValue) {
			case "SCOPE-ENABLED":
			case "SCOPE-DISABLED":
				// Prompt the user for the new admin consent display name.
				const adminConsentDisplayName = await this.inputAdminConsentDisplayName("Edit Exposed API Permission (1/1)", scope.adminConsentDisplayName!, validateScopeAdminDisplayName);

				// If escape is pressed or the new display name is empty then return undefined.
				if (adminConsentDisplayName === undefined) {
					return undefined;
				}

				scope.adminConsentDisplayName = adminConsentDisplayName;
				break;
			case "SCOPE-VALUE":
				const value = await this.inputValue("Edit Exposed API Permission (1/1)", scope.value!, true, properties.signInAudience!, properties.api!, validateScopeValue);

				// If escape is pressed or the new name is empty then return undefined.
				if (value === undefined) {
					return undefined;
				}

				scope.value = value;
				break;
			case "SCOPE-DESCRIPTION":
				// Prompt the user for the new admin consent description.
				const adminConsentDescription = await this.inputAdminConsentDescription("Edit Exposed API Permission (1/1)", scope.adminConsentDescription!, validateScopeAdminDescription);

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
		await this.updateApplication(item.objectId!, { api: properties.api }, status);
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
		await this.updateApplication(item.objectId!, { api: properties.api }, status);
	}

	// Deletes an exposed api scope from an application registration.
	async delete(item: AppRegItem, hideConfirmation: boolean = false): Promise<void> {
		if (item.state !== false && hideConfirmation === false) {
			const disableScope = await window.showWarningMessage(`The Scope ${item.label} cannot be deleted unless it is disabled. Do you want to disable the scope and then delete it?`, "Yes", "No");
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

		// Get the parent application so we can read the app roles.
		const properties = await this.getProperties(item.objectId!);

		// If the array is undefined then it'll be an Azure CLI authentication issue.
		if (properties === undefined) {
			return;
		}

		// Remove the app role from the array.
		properties.api!.oauth2PermissionScopes!.splice(
			properties.api!.oauth2PermissionScopes!.findIndex((r) => r.id === item.value!),
			1
		);

		// Update the application.
		await this.updateApplication(item.objectId!, { api: properties.api }, status);
	}

	// Captures the value for a scope.
	async inputValue(title: string, existingValue: string | undefined, isEditing: boolean, signInAudience: string, scopes: ApiApplication, validate: (value: string, isEditing: boolean, existingValue: string | undefined, signInAudience: string, scopes: ApiApplication) => string | undefined): Promise<string | undefined> {
		// Debounce the validation function to prevent multiple calls to the Graph API.
		const debouncedValidation = debounce(validate, 500);

		// Prompt the user for the new value.
		return await window.showInputBox({
			prompt: "Scope name",
			placeHolder: "Enter a scope name (e.g. Files.Read)",
			title: title,
			ignoreFocusOut: true,
			value: existingValue,
			validateInput: async (value) => debouncedValidation(value, isEditing, existingValue, signInAudience, scopes)
		});
	}

	// Captures the admin description for a scope.
	async inputAdminConsentDescription(title: string, existingValue: string | undefined, validate: (description: string) => string | undefined): Promise<string | undefined> {
		return await window.showInputBox({
			prompt: "Admin consent description",
			placeHolder: "Enter an admin consent description (e.g. Allows the app to read files on your behalf.)",
			title: title,
			ignoreFocusOut: true,
			value: existingValue,
			validateInput: async (value) => validate(value)
		});
	}

	// Captures the user display name for a scope.
	async inputAdminConsentDisplayName(title: string, existingValue: string | undefined, validate: (displayName: string) => string | undefined): Promise<string | undefined> {
		return await window.showInputBox({
			prompt: "Admin consent display name",
			placeHolder: "Enter an admin consent display name (e.g. Read files)",
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
	async inputScopeDetails(scope: PermissionScope, id: string, isEditing: boolean, signInAudience: string, scopes: ApiApplication, validate: (displayName: string) => string | undefined): Promise<PermissionScope | undefined> {
		const value = await this.inputValue(isEditing === true ? "Edit Exposed API Permission (1/7)" : "Add API Exposed Permission (1/7)", scope.value ?? undefined, isEditing, signInAudience, scopes, validateScopeValue);

		// If escape is pressed or the new name is empty then return undefined.
		if (value === undefined) {
			return undefined;
		}

		// Prompt the user for the new allowed member types.
		const consentType = await this.inputConsentType(isEditing === true ? "Edit Exposed API Permission (2/7)" : "Add API Exposed Permission (2/7)");

		// If escape is pressed or the new allowed member types is empty then return undefined.
		if (consentType === undefined) {
			return undefined;
		}

		// Prompt the user for the new admin consent display name.
		const adminConsentDisplayName = await this.inputAdminConsentDisplayName(isEditing === true ? "Edit Exposed API Permission (3/7)" : "Add API Exposed Permission (3/7)", scope.adminConsentDisplayName ?? undefined, validateScopeAdminDisplayName);

		// If escape is pressed or the new display name is empty then return undefined.
		if (adminConsentDisplayName === undefined) {
			return undefined;
		}

		// Prompt the user for the new admin consent description.
		const adminConsentDescription = await this.inputAdminConsentDescription(isEditing === true ? "Edit Exposed API Permission (4/7)" : "Add API Exposed Permission (4/7)", scope.adminConsentDescription ?? undefined, validateScopeAdminDescription);

		// If escape is pressed or the new description is empty then return undefined.
		if (adminConsentDescription === undefined) {
			return undefined;
		}

		// Prompt the user for the new user consent display name.
		const userConsentDisplayName = await window.showInputBox({
			prompt: "User consent display name",
			placeHolder: "Enter an optional user consent display name (e.g. Read your files)",
			title: isEditing === true ? "Edit Exposed API Permission (5/7)" : "Add API Exposed Permission (5/7)",
			ignoreFocusOut: true,
			value: scope.userConsentDisplayName ?? undefined,
			validateInput: async (value) => validate(value)
		});

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
