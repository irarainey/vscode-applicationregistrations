import { window, QuickPickItem } from 'vscode';
import { AppRegTreeDataProvider } from '../data/appRegTreeDataProvider';
import { AppRegItem } from '../models/appRegItem';
import { ServiceBase } from './serviceBase';
import { GraphClient } from '../clients/graph';
import { RequiredResourceAccess, NullableOption } from '@microsoft/microsoft-graph-types';
import { sort } from 'fast-sort';

export class RequiredResourceAccessService extends ServiceBase {

    // The constructor for the RequiredResourceAccessService class.
    constructor(treeDataProvider: AppRegTreeDataProvider, graphClient: GraphClient) {
        super(treeDataProvider, graphClient);
    }

    // Adds the selected scope to an application registration.
    public async add(item: AppRegItem): Promise<void> {

        // Prompt the user for the new value.
        const apiAppSearch = await window.showInputBox({
            prompt: "Search for API Application",
            placeHolder: "Enter the starting characters of the API application name to build a list of matching applications.",
            ignoreFocusOut: true,
            validateInput: (value) => {
                if (value.length < 3) {
                    return "You must enter at least partial name of the API application to filter the list. A minimum of 3 characters is required.";
                }
                return;
            }
        });

        // If the user cancels the input then return undefined.
        if (apiAppSearch === undefined) {
            return;
        }

        const servicePrincipals = await this.graphClient.getServicePrincipalByDisplayName(apiAppSearch);

        if(servicePrincipals.length === 0) {
            window.showInformationMessage("No API applications were found that match the search criteria.", "OK");
            return;
        }

        const newList = sort(servicePrincipals).asc(x => x.appDisplayName).map(r => {
            return {
                label: r.appDisplayName!,
                description: r.appDescription!,
                value: r.appId
            };
        });

        // Prompt the user for the new allowed member types.
        const allowed = await window.showQuickPick(
            newList,
            {
                placeHolder: "Select an API application",
                ignoreFocusOut: true
            });

        // If the user cancels the input then return undefined.
        if (allowed === undefined) {
            return;
        }

        //Now we have the API application ID we can call the addToExisting method.
        this.addToExisting(item, allowed.value);
    }

    // Adds the selected scope from an existing API to an application registration.
    public async addToExisting(item: AppRegItem, existingApiId?: NullableOption<string> | undefined): Promise<void> {

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
                ignoreFocusOut: true
            });

        // If the user cancels the input then drop out.
        if (type === undefined) {
            return;
        }

        // Get the previous icon so we can revert if the update fails.
        const previousIcon = item.iconPath;

        // Update the tree item icon to show the loading animation.
        const status = this.triggerTreeChange("Loading api permissions...", item);

        const apiIdToUse = existingApiId === undefined ? item.resourceAppId : existingApiId;

        // Get the service principal for the application so we can get the scopes.
        const servicePrincipals = await this.graphClient.getServicePrincipalByAppId(apiIdToUse!);

        // Get all the existing scopes for the application.
        const allAssignedScopes = await this.getApiPermissions(item.objectId!);

        // Find the api app in the collection.
        const apiAppScopes = allAssignedScopes.filter(r => r.resourceAppId === apiIdToUse!)[0];

        // Define a variable to hold the selected scope item.
        let scopeItem: any;

        // Prompt the user for the scope or role to add from those available.
        if (type.value === "Scope") {
            if (apiAppScopes !== undefined) {
                // Remove any scopes that are already assigned.
                apiAppScopes.resourceAccess!.forEach(r => {
                    servicePrincipals.oauth2PermissionScopes = servicePrincipals.oauth2PermissionScopes!.filter(s => s.id !== r.id);
                });

                // If there are no scopes available then drop out.
                if (servicePrincipals.oauth2PermissionScopes === undefined || servicePrincipals.oauth2PermissionScopes!.length === 0) {
                    status?.dispose();
                    this.setTreeItemIcon(item, previousIcon, false);
                    window.showInformationMessage("There are no unassigned user delegated permissions available to add to this application registration.", "OK");
                    return;
                }
            }

            status?.dispose();
            this.setTreeItemIcon(item, previousIcon, false);

            // Prompt the user for the scope to add.
            const permissions = sort(servicePrincipals.oauth2PermissionScopes!)
                .asc(r => r.value!)
                .map(r => {
                    return {
                        label: r.value!,
                        description: r.adminConsentDescription!,
                        value: r.id!
                    } as QuickPickItem;
                });
            scopeItem = await window.showQuickPick(
                permissions,
                {
                    placeHolder: "Select a user delegated permission",
                    ignoreFocusOut: true,
                });

            // If the user cancels the input then drop out.
            if (scopeItem === undefined) {
                return;
            }
        } else {
            if (apiAppScopes !== undefined) {
                // Remove any scopes that are already assigned.
                apiAppScopes.resourceAccess!.forEach(r => {
                    servicePrincipals.appRoles = servicePrincipals.appRoles!.filter(s => s.id !== r.id);
                });

                // If there are no scopes available then drop out.
                if (servicePrincipals.appRoles === undefined || servicePrincipals.appRoles!.length === 0) {
                    status?.dispose();
                    this.setTreeItemIcon(item, previousIcon, false);
                    window.showInformationMessage("There are no unassigned application permissions available to add to this application registration.", "OK");
                    return;
                }
            }

            status?.dispose();
            this.setTreeItemIcon(item, previousIcon, false);

            // Prompt the user for the scope to add.
            const permissions = sort(servicePrincipals.appRoles!)
                .asc(r => r.value!)
                .map(r => {
                    return {
                        label: r.value!,
                        description: r.description!,
                        value: r.id!
                    } as QuickPickItem;
                });
            scopeItem = await window.showQuickPick(
                permissions,
                {
                    placeHolder: "Select an application permission",
                    ignoreFocusOut: true
                });

            // If the user cancels the input then drop out.
            if (scopeItem === undefined) {
                return;
            }
        }

        // Set the added trigger to the status bar message.
        const statusAdd = this.triggerTreeChange("Adding api permission...", item);

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
                resourceAccess: [{
                    id: scopeItem.value,
                    type: type.value
                }]
            });
        }

        //Update the application.
        this.graphClient.updateApplication(item.objectId!, { requiredResourceAccess: allAssignedScopes })
            .then(() => {
                this.triggerOnComplete({ success: true, statusBarHandle: statusAdd });
            })
            .catch((error) => {
                this.triggerOnError({ success: false, statusBarHandle: statusAdd, error: error, treeViewItem: item, previousIcon: previousIcon });
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
}