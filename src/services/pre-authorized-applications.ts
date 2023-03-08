import { QuickPickItem, window } from "vscode";
import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { Application, NullableOption, ServicePrincipal } from "@microsoft/microsoft-graph-types";
import { ServiceBase } from "./service-base";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { debounce } from "ts-debounce";
import { GraphResult } from "../types/graph-result";
import { clearStatusBarMessage } from "../utils/status-bar";
import { validateDebouncedInput } from "../utils/validation";
import { sort } from "fast-sort";

export class PreAuthorizedApplicationsService extends ServiceBase {
	// The constructor for the PreAuthorizedApplicationsService class.
	constructor(graphRepository: GraphApiRepository, treeDataProvider: AppRegTreeDataProvider) {
		super(graphRepository, treeDataProvider);
	}

	// Add a scope to a new authorised client application.
	async addAuthorisedClientScope(item: AppRegItem): Promise<void> {
		// Debounce the validation function to prevent multiple calls to the Graph API.
		const debouncedValidation = debounce(validateDebouncedInput, 500);

		// Prompt the user for the new value.
		const appSearch = await window.showInputBox({
			prompt: "Search for Application",
			placeHolder: "Enter part of an Application name to build a list of matching applications.",
			ignoreFocusOut: true,
			title: "Add Authorized Client Application (1/3)",
			validateInput: (value) => debouncedValidation(value)
		});

		// If the user cancels the input then return undefined.
		if (appSearch === undefined) {
			return;
		}

		// Update the tree item icon to show the loading animation.
		const status = this.indicateChange("Loading Applications...", item);

		// Get the service principals that match the search criteria.
		const result: GraphResult<ServicePrincipal[]> = await this.graphRepository.findServicePrincipalsByDisplayName(appSearch);
		if (result.success === false || result.value === undefined) {
			await this.handleError(result.error);
			return;
		}

		const servicePrincipals: ServicePrincipal[] = result.value;

		// If there are no service principals found then drop out.
		if (servicePrincipals.length === 0) {
			window.showInformationMessage("No Applications were found that match the search criteria.", "OK");
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
			placeHolder: "Select an Application",
			title: "Add Authorized Client Application (2/3)",
			ignoreFocusOut: true
		});

		// If the user cancels the input then return undefined.
		if (allowed === undefined) {
			return;
		}

		//Now we have the API application ID we can call the addToExisting method.
		await this.addToExistingAuthorisedClient(item, allowed.value);
	}

	// Add a new scope to an existing authorised client application.
	async addToExistingAuthorisedClient(item: AppRegItem, clientAppId?: NullableOption<string> | undefined): Promise<void> {
		const startStep = clientAppId === undefined ? 0 : 2;
		const numberOfSteps = clientAppId === undefined ? 1 : 3;

		// Update the tree item icon to show the loading animation.
		const status = this.indicateChange(`Loading available Scopes...`, item);

		// Get the service principal for the application so we can get the scopes.
		const result: GraphResult<ServicePrincipal> = await this.graphRepository.findServicePrincipalByAppId(item.value!);
		if (result.success === false || result.value === undefined) {
			await this.handleError(result.error);
			return;
		}

		const servicePrincipal: ServicePrincipal = result.value;

		// Get all the properties for the application.
		const properties = await this.getApi(item.objectId!);

		// If the array is undefined then it'll be an Azure CLI authentication issue.
		if (properties === undefined) {
			return;
		}

		// Find the client app in the collection.
		const appScopes = properties.api?.preAuthorizedApplications!.filter((r) => r.appId === item.resourceAppId!)[0];

		// Define a variable to hold the selected scope item.
		let scopeItem: any;

		// Remove any scopes that are already assigned.
		if (appScopes !== undefined) {
			appScopes.delegatedPermissionIds!.forEach((r) => {
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
			title: `Add Authorized Client Application (${startStep + 1}/${numberOfSteps})`,
			ignoreFocusOut: true
		});

		// If the user cancels the input then drop out.
		if (scopeItem === undefined) {
			return;
		}

		// Set the added trigger to the status bar message.
		const addStatus = this.indicateChange("Adding Authorized Client Scope...", item);

		if (appScopes !== undefined) {
			// Add the new scope to the existing app.
			appScopes.delegatedPermissionIds!.push(scopeItem.value);
		} else {
			// Add the new scope to a new api app.
			properties.api?.preAuthorizedApplications!.push({
				appId: clientAppId!,
				delegatedPermissionIds: [scopeItem.value]
			});
		}

		//Update the application.
		await this.updateApplication(item.objectId!, { api: properties.api }, addStatus);
	}

	// Remove a scope from an authorized client application.
	async removeAuthorisedClientScope(item: AppRegItem): Promise<void> {
		// Prompt the user to confirm the removal.
		const answer = await window.showWarningMessage(`Do you want to remove the Authorized Client Scope ${item.label}?`, "Yes", "No");

		// If the user confirms the removal then delete the role.
		if (answer === "Yes") {
			// Set the added trigger to the status bar message.
			const status = this.indicateChange("Removing Authorized Client Scope...", item);

			// Get all the properties for the application.
			const properties = await this.getApi(item.objectId!);

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

	// Remove an authorized client application.
	async removeAuthorisedClient(item: AppRegItem): Promise<void> {
		// Prompt the user to confirm the removal.
		const answer = await window.showWarningMessage(`Do you want to remove the Authorized Client Application ${item.label}?`, "Yes", "No");

		// If the user confirms the removal then remove the client application.
		if (answer === "Yes") {
			// Set the added trigger to the status bar message.
			const status = this.indicateChange("Removing Authorized Client Application...", item);

			// Get all the properties for the application.
			const properties = await this.getApi(item.objectId!);

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

	// Gets the required properties from the application registration.
	private async getApi(id: string): Promise<Application | undefined> {
		const result: GraphResult<Application> = await this.graphRepository.getApplicationDetailsPartial(id, "api");
		if (result.success === true && result.value !== undefined) {
			return result.value;
		} else {
			await this.handleError(result.error);
			return undefined;
		}
	}
}
