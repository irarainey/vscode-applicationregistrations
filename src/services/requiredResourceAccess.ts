import { window } from 'vscode';
import { AppRegTreeDataProvider } from '../data/appRegTreeDataProvider';
import { AppRegItem } from '../models/appRegItem';
import { ServiceBase } from './serviceBase';
import { GraphClient } from '../clients/graph';
import { RequiredResourceAccess } from '@microsoft/microsoft-graph-types';
import { sort } from 'fast-sort';

export class RequiredResourceAccessService extends ServiceBase {

    // The constructor for the RequiredResourceAccessService class.
    constructor(treeDataProvider: AppRegTreeDataProvider, graphClient: GraphClient) {
        super(treeDataProvider, graphClient);
    }

    // Adds the selected scope to an application registration.
    public async add(item: AppRegItem): Promise<void> {

        console.log("addScope");

        // const servicePrincipals = (await this.graphClient.getServicePrincipalByDisplayName("")).map(r => r.displayName!);

        // // Prompt the user for the new allowed member types.
        // const allowed = await window.showQuickPick(
        //     servicePrincipals,
        //     {
        //         placeHolder: "Select an API application",
        //         ignoreFocusOut: true
        //     });


        // // Debounce the validation function to prevent multiple calls to the Graph API.
        // const validation = async (value: string) => this.validateValue(value);
        // const debouncedValidation = debounce(validation, 500);

        // // Prompt the user for the new value.
        // const scope = await window.showInputBox({
        //     prompt: "Edit scope",
        //     placeHolder: "Enter a scope to add to the application registration",
        //     ignoreFocusOut: true,
        //     validateInput: async (value) => debouncedValidation(value)
        // });

        // // If the user cancels the input then return undefined.
        // if (scope === undefined) {
        //     return;
        // }

        // // Set the added trigger to the status bar message.
        // const previousIcon = item.iconPath;
        // const status = this.triggerTreeChange("Adding new api permission...", item);

        // // Get the existing scopes.
        // const allAssignedScopes = await this.getApiPermissions(item.objectId!);

        // Check to see if the api app is already in the collection.
        //const apiAppScopes = allAssignedScopes.filter(r => r.resourceAppId === item.resourceAppId!)[0];

        // if(apiAppScopes !== undefined) {
        //     // Add the new scope to the existing app.
        //     apiAppScopes.resourceAccess!.push(scope);
        // } else {
        //     // Add the new scope to a new api app.
        //     allAssignedScopes.push({
        //         resourceAppId: item.resourceAppId!,
        //         resourceAccess: [scope]
        //     });
        // }

        // // Update the application.
        // this.graphClient.updateApplication(item.objectId!, { requiredResourceAccess: allAssignedScopes })
        //     .then(() => {
        //         this.triggerOnComplete({ success: true, statusBarHandle: status });
        //     })
        //     .catch((error) => {
        //         this.triggerOnError({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon });
        //     });        
    }

    // Adds the selected scope from an existing API to an application registration.
    public async addToExisting(item: AppRegItem): Promise<void> {

        // Determine the type of permission.
        const type = await window.showQuickPick(
            [
                "Delegated permission",
                "Application permission",
            ],
            {
                placeHolder: "Select a permission type",
                ignoreFocusOut: true
            });

        // If the user cancels the input then drop out.
        if (type === undefined) {
            return;
        }

        // Get the service principal for the application so we can get the scopes.
        const servicePrincipals = await this.graphClient.getServicePrincipalByAppId(item.resourceAppId!);

        // Get all the existing scopes for the application.
        const allAssignedScopes = await this.getApiPermissions(item.objectId!);

        // Find the api app in the collection.
        const apiAppScopes = allAssignedScopes.filter(r => r.resourceAppId === item.resourceAppId!)[0];

        let scopeId: string;
        let scopeValue: string | undefined = undefined;

        // Prompt the user for the scope or role to add from those available.
        if (type === "Delegated permission") {
            const permissions = sort(servicePrincipals.oauth2PermissionScopes!).asc(r => r.value!).map(r => r.value!);
            scopeValue = await window.showQuickPick(
                permissions,
                {
                    placeHolder: "Select a user delegated permission",
                    ignoreFocusOut: true,
                });

            // If the user cancels the input then drop out.
            if (scopeValue === undefined) {
                return;
            }

            // Find the id of the scope selected
            scopeId = servicePrincipals.oauth2PermissionScopes!.filter(r => r.value === scopeValue)[0].id!;

        } else {
            const permissions = sort(servicePrincipals.appRoles!).asc(r => r.value!).map(r => r.value!);
            scopeValue = await window.showQuickPick(
                permissions,
                {
                    placeHolder: "Select an application permission",
                    ignoreFocusOut: true
                });

            // If the user cancels the input then drop out.
            if (scopeValue === undefined) {
                return;
            }

            // Find the id of the role selected
            scopeId = servicePrincipals.appRoles!.filter(r => r.value === scopeValue)[0].id!;
        }

        // Get the previous icon so we can revert if the update fails.
        const previousIcon = item.iconPath;

        // Set the added trigger to the status bar message.
        const status = this.triggerTreeChange("Adding api permission...", item);

        // Add the new scope to the existing app.
        apiAppScopes.resourceAccess!.push({
            id: scopeId,
            type: type === "Delegated permissions" ? "Scope" : "Role"
        });

        //Update the application.
        this.graphClient.updateApplication(item.objectId!, { requiredResourceAccess: allAssignedScopes })
            .then(() => {
                this.triggerOnComplete({ success: true, statusBarHandle: status });
            })
            .catch((error) => {
                this.triggerOnError({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon });
            });
    }

    // Removes the selected scope from an application registration.
    public async remove(item: AppRegItem): Promise<void> {

        // Prompt the user to confirm the removal.
        const answer = await window.showInformationMessage(`Do you want to remove the api permission ${item.value}?`, "Yes", "No");

        // If the user confirms the removal then delete the role.
        if (answer === "Yes") {
            // Set the added trigger to the status bar message.
            const previousIcon = item.iconPath;
            const status = this.triggerTreeChange("Removing api permission...", item);

            // Get all the scopes for the application.
            const allAssignedScopes = await this.getApiPermissions(item.objectId!);

            // Find the api app in the collection.
            const apiAppScopes = allAssignedScopes.filter(r => r.resourceAppId === item.resourceAppId!)[0];

            // Remove the scope requested.
            apiAppScopes.resourceAccess!.splice(apiAppScopes.resourceAccess!.findIndex(x => x.id === item.resourceScopeId!), 1);

            // If there are no more scopes for the api app then remove the api app from the collection.
            if (apiAppScopes.resourceAccess!.length === 0) {
                allAssignedScopes.splice(allAssignedScopes.findIndex(r => r.resourceAppId === item.resourceAppId!), 1);
            }

            //Update the application.
            this.graphClient.updateApplication(item.objectId!, { requiredResourceAccess: allAssignedScopes })
                .then(() => {
                    this.triggerOnComplete({ success: true, statusBarHandle: status });
                })
                .catch((error) => {
                    this.triggerOnError({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon });
                });
        }
    }

    // Removes all scopes for an api app from an application registration.
    public async removeApi(item: AppRegItem): Promise<void> {

        // Prompt the user to confirm the removal.
        const answer = await window.showInformationMessage(`Do you want to remove all api permissions for ${item.label}?`, "Yes", "No");

        // If the user confirms the removal then remove the scopes.
        if (answer === "Yes") {
            // Set the added trigger to the status bar message.
            const previousIcon = item.iconPath;
            const status = this.triggerTreeChange("Removing api permissions...", item);

            // Get all the scopes for the application.
            const allAssignedScopes = await this.getApiPermissions(item.objectId!);

            // Find the requested api app in the collection and remove it.
            allAssignedScopes.splice(allAssignedScopes.findIndex(r => r.resourceAppId === item.resourceAppId!), 1);

            //Update the application.
            this.graphClient.updateApplication(item.objectId!, { requiredResourceAccess: allAssignedScopes })
                .then(() => {
                    this.triggerOnComplete({ success: true, statusBarHandle: status });
                })
                .catch((error) => {
                    this.triggerOnError({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon });
                });
        }
    }

    // Gets the api permissions for an application registration.
    private async getApiPermissions(id: string): Promise<RequiredResourceAccess[]> {
        return (await this.dataProvider.getApplicationPartial(id, "requiredResourceAccess")).requiredResourceAccess!;
    }

    // Validates scope.
    private async validateValue(value: string): Promise<string | undefined> {

        // // Check the length of the value.
        // if (value.length > 250) {
        //     return "A value cannot be longer than 250 characters.";
        // }

        // Check the length of the display name.
        if (value.length < 1) {
            return "A value cannot be empty.";
        }

        // const roles = await this.getAppRoles(id);
        // if (roles!.find(r => r.value === value) !== undefined) {
        //     return "The value specified already exists.";
        // }

        return undefined;
    }
}