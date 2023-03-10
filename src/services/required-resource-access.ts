import { window, QuickPickItem } from "vscode";
import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { ServiceBase } from "./service-base";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { RequiredResourceAccess, NullableOption, Application, ServicePrincipal } from "@microsoft/microsoft-graph-types";
import { sort } from "fast-sort";
import { debounce } from "ts-debounce";
import { GraphResult } from "../types/graph-result";
import { clearStatusBarMessage } from "../utils/status-bar";
import { validateDebouncedInput } from "../utils/validation";

export class RequiredResourceAccessService extends ServiceBase {
	// The constructor for the RequiredResourceAccessService class.
	constructor(graphRepository: GraphApiRepository, treeDataProvider: AppRegTreeDataProvider) {
		super(graphRepository, treeDataProvider);
	}

	// Adds the selected scope to an application registration.
	async add(item: AppRegItem): Promise<void> {
		// Debounce the validation function to prevent multiple calls to the Graph API.
		const debouncedValidation = debounce(validateDebouncedInput, 500);

		// Prompt the user for the new value.
		const apiAppSearch = await window.showInputBox({
			prompt: "Search for API Application",
			placeHolder: "Enter part of an API Application name to build a list of matching applications.",
			ignoreFocusOut: true,
			title: "Add API Permission (1/4)",
			validateInput: (value) => debouncedValidation(value)
		});

		// If the user cancels the input then return undefined.
		if (apiAppSearch === undefined) {
			return;
		}

		// Update the tree item icon to show the loading animation.
		const status = this.indicateChange("Loading API Applications...", item);

		// Get the service principals that match the search criteria.
		const result: GraphResult<ServicePrincipal[]> = await this.graphRepository.findServicePrincipalsByDisplayName(apiAppSearch);
		if (result.success === false || result.value === undefined) {
			await this.handleError(result.error);
			return;
		}

		const servicePrincipals: ServicePrincipal[] = result.value;

		// If there are no service principals found then drop out.
		if (servicePrincipals.length === 0) {
			window.showInformationMessage("No API Applications were found that match the search criteria.", "OK");
			clearStatusBarMessage(status!);
			this.resetTreeItemIcon(item);
			return;
		}

		// Sort the list of service principals by display name.
		const newList = sort(servicePrincipals)
			.asc((x) => x.appDisplayName)
			.map((r) => {
				return {
					label: r.appDisplayName!,
					description: r.appDescription!,
					value: r.appId
				};
			});

		clearStatusBarMessage(status!);
		this.resetTreeItemIcon(item);

		// Prompt the user for the new allowed member types.
		const allowed = await window.showQuickPick(newList, {
			placeHolder: "Select an API Application",
			title: "Add API Permission (2/4)",
			ignoreFocusOut: true
		});

		// If the user cancels the input then return undefined.
		if (allowed === undefined) {
			return;
		}

		//Now we have the API application ID we can call the addToExisting method.
		await this.addToExisting(item, allowed.value, allowed.label);
	}

	// Adds the selected scope from an existing API to an application registration.
	async addToExisting(item: AppRegItem, existingApiId?: NullableOption<string> | undefined, existingApiAppName?: string): Promise<void> {
		const startStep = existingApiId === undefined ? 1 : 3;
		const numberOfSteps = existingApiId === undefined ? 2 : 4;
		const apiAppName = existingApiId === undefined ? item.label : existingApiAppName;

		// Determine the type of permission.
		const type = await window.showQuickPick(
			[
				{
					label: "Delegated permissions",
					detail: "Your application needs to access the API as the signed-in user.",
					value: "Scope"
				},
				{
					label: "Application permissions",
					detail: "Your application runs as a background service or daemon without a signed-in user.",
					value: "Role"
				}
			],
			{
				placeHolder: "What type of permissions does your application require?",
				title: `Add API Permission (${startStep}/${numberOfSteps})`,
				ignoreFocusOut: true
			}
		);

		// If the user cancels the input then drop out.
		if (type === undefined) {
			return;
		}

		// Update the tree item icon to show the loading animation.
		const status = this.indicateChange(`Loading API Permissions for ${apiAppName}...`, item);

		// Determine the API application ID to use.
		const apiIdToUse = existingApiId === undefined ? item.resourceAppId : existingApiId;

		// Get the service principal for the application so we can get the scopes.
		const result: GraphResult<ServicePrincipal> = await this.graphRepository.findServicePrincipalByAppId(apiIdToUse!);
		if (result.success === false || result.value === undefined) {
			await this.handleError(result.error);
			return;
		}

		const servicePrincipal: ServicePrincipal = result.value;

		// Get all the existing scopes for the application.
		const allAssignedScopes = await this.getApiPermissions(item.objectId!);

		// If the array is undefined then it'll be an Azure CLI authentication issue.
		if (allAssignedScopes === undefined) {
			return;
		}

		// Find the api app in the collection.
		const apiAppScopes = allAssignedScopes.filter((r) => r.resourceAppId === apiIdToUse!)[0];

		// Define a variable to hold the selected scope item.
		let scopeItem: any;

		// Prompt the user for the scope or role to add from those available.
		if (type.value === "Scope") {
			// Remove any scopes that are already assigned.
			if (apiAppScopes !== undefined) {
				apiAppScopes.resourceAccess!.forEach((r) => {
					servicePrincipal.oauth2PermissionScopes = servicePrincipal.oauth2PermissionScopes!.filter((s) => s.id !== r.id);
				});
			}

			clearStatusBarMessage(status!);
			this.resetTreeItemIcon(item);

			// If there are no scopes available then drop out.
			if (servicePrincipal.oauth2PermissionScopes === undefined || servicePrincipal.oauth2PermissionScopes!.length === 0) {
				window.showInformationMessage("There are no user delegated permissions available to add to this application registration.", "OK");
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
				placeHolder: "Select a user delegated permission",
				title: `Add API Permission (${startStep + 1}/${numberOfSteps})`,
				ignoreFocusOut: true
			});

			// If the user cancels the input then drop out.
			if (scopeItem === undefined) {
				return;
			}
		} else {
			// Remove any scopes that are already assigned.
			if (apiAppScopes !== undefined) {
				apiAppScopes.resourceAccess!.forEach((r) => {
					servicePrincipal.appRoles = servicePrincipal.appRoles!.filter((s) => s.id !== r.id);
				});
			}

			clearStatusBarMessage(status!);
			this.resetTreeItemIcon(item);

			// If there are no scopes available then drop out.
			if (servicePrincipal.appRoles === undefined || servicePrincipal.appRoles!.length === 0) {
				window.showInformationMessage("There are no application permissions available to add to this application registration.", "OK");
				return;
			}

			// Prompt the user for the scope to add.
			const permissions = sort(servicePrincipal.appRoles!)
				.asc((r) => r.value!)
				.map((r) => {
					return {
						label: r.value!,
						description: r.description!,
						value: r.id!
					} as QuickPickItem;
				});
			scopeItem = await window.showQuickPick(permissions, {
				placeHolder: "Select an API Permission",
				title: `Add API Permission (${startStep + 1}/${numberOfSteps})`,
				ignoreFocusOut: true
			});

			// If the user cancels the input then drop out.
			if (scopeItem === undefined) {
				return;
			}
		}

		// Set the added trigger to the status bar message.
		const addStatus = this.indicateChange("Adding API Permission...", item);

		if (apiAppScopes !== undefined) {
			// Add the new scope to the existing app.
			apiAppScopes.resourceAccess!.push({
				id: scopeItem.value,
				type: type.value
			});
		} else {
			// Add the new scope to a new api app.
			allAssignedScopes.push({
				resourceAppId: apiIdToUse!,
				resourceAccess: [
					{
						id: scopeItem.value,
						type: type.value
					}
				]
			});
		}

		//Update the application.
		await this.updateApplication(item.objectId!, { requiredResourceAccess: allAssignedScopes }, addStatus);
	}

	// Removes the selected scope from an application registration.
	async remove(item: AppRegItem): Promise<void> {
		// Prompt the user to confirm the removal.
		const answer = await window.showWarningMessage(`Do you want to remove the API Permission ${item.value}?`, "Yes", "No");

		// If the user confirms the removal then delete the role.
		if (answer === "Yes") {
			// Set the added trigger to the status bar message.
			const status = this.indicateChange("Removing API Permission...", item);

			// Get all the scopes for the application.
			const allAssignedScopes = await this.getApiPermissions(item.objectId!);

			// If the array is undefined then it'll be an Azure CLI authentication issue.
			if (allAssignedScopes === undefined) {
				return;
			}

			// Find the api app in the collection.
			const apiAppScopes = allAssignedScopes.filter((r) => r.resourceAppId === item.resourceAppId!)[0];

			// Remove the scope requested.
			apiAppScopes.resourceAccess!.splice(
				apiAppScopes.resourceAccess!.findIndex((x) => x.id === item.resourceScopeId!),
				1
			);

			// If there are no more scopes for the api app then remove the api app from the collection.
			if (apiAppScopes.resourceAccess!.length === 0) {
				allAssignedScopes.splice(
					allAssignedScopes.findIndex((r) => r.resourceAppId === item.resourceAppId!),
					1
				);
			}

			//Update the application.
			await this.updateApplication(item.objectId!, { requiredResourceAccess: allAssignedScopes }, status);
		}
	}

	// Removes all scopes for an api app from an application registration.
	async removeApi(item: AppRegItem): Promise<void> {
		// Prompt the user to confirm the removal.
		const answer = await window.showWarningMessage(`Do you want to remove all API Permissions for ${item.label}?`, "Yes", "No");

		// If the user confirms the removal then remove the scopes.
		if (answer === "Yes") {
			// Set the added trigger to the status bar message.
			const status = this.indicateChange("Removing API Permissions...", item);

			// Get all the scopes for the application.
			const allAssignedScopes = await this.getApiPermissions(item.objectId!);

			// If the array is undefined then it'll be an Azure CLI authentication issue.
			if (allAssignedScopes === undefined) {
				return;
			}

			// Find the requested api app in the collection and remove it.
			allAssignedScopes.splice(
				allAssignedScopes.findIndex((r) => r.resourceAppId === item.resourceAppId!),
				1
			);

			//Update the application.
			await this.updateApplication(item.objectId!, { requiredResourceAccess: allAssignedScopes }, status);
		}
	}

	// Gets the api permissions for an application registration.
	private async getApiPermissions(id: string): Promise<RequiredResourceAccess[] | undefined> {
		const result: GraphResult<Application> = await this.graphRepository.getApplicationDetailsPartial(id, "requiredResourceAccess");
		if (result.success === true && result.value !== undefined) {
			return result.value.requiredResourceAccess!;
		} else {
			await this.handleError(result.error);
			return undefined;
		}
	}
}
