import { window } from 'vscode';
import { AppRegTreeDataProvider } from '../data/appRegTreeDataProvider';
import { AppRegItem } from '../models/appRegItem';
import { ServiceBase } from './serviceBase';
import { GraphClient } from '../clients/graph';
import { RequiredResourceAccess } from '@microsoft/microsoft-graph-types';

export class RequiredResourceAccessService extends ServiceBase {

    // The constructor for the RequiredResourceAccessService class.
    constructor(treeDataProvider: AppRegTreeDataProvider, graphClient: GraphClient) {
        super(treeDataProvider, graphClient);
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