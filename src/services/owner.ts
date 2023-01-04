import { window, env, Uri } from 'vscode';
import { portalUserUri } from '../constants';
import { AppRegDataProvider } from '../data/applicationRegistration';
import { AppRegItem } from '../models/appRegItem';
import { ServiceBase } from './serviceBase';
import { GraphClient } from '../clients/graph';
import { User } from "@microsoft/microsoft-graph-types";

export class OwnerService extends ServiceBase {

    // The list of users in the directory.
    private _userList: any = undefined;

    // The constructor for the OwnerService class.
    constructor(dataProvider: AppRegDataProvider, graphClient: GraphClient) {
        super(dataProvider, graphClient);
    }

    // Adds a new owner to an application registration.
    public async add(item: AppRegItem): Promise<void> {

        // Get the existing owners.
        const existingOwners = (await this._graphClient.getApplicationOwners(item.objectId!)).value;

        // Prompt the user for the new owner.
        const owner = await window.showInputBox({
            placeHolder: "Enter user name or email address...",
            prompt: "Add new owner to application",
            ignoreFocusOut: true,
            validateInput: async (value) => await this.validateOwner(value, existingOwners)
        });

        // If the new owner name is not empty then add as an owner.
        if (owner !== undefined) {
            // Set the added trigger to the status bar message.
            const previousIcon = item.iconPath;
            const status = this.indicateChange("Adding owner...", item);
            this._graphClient.addApplicationOwner(item.objectId!, this._userList.value[0].id)
                .then(() => {
                    this._onComplete.fire({ success: true, statusBarHandle: status });
                })
                .catch((error) => {
                    this._onError.fire({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon });
                });
        }
    }

    // Removes an owner from an application registration.
    public async remove(item: AppRegItem): Promise<void> {

        // Prompt the user to confirm the removal.
        const response = await window.showInformationMessage(`Do you want to remove ${item.label} as an owner of this application?`, "Yes", "No");

        // If the user confirms the removal then remove the user.
        if (response === "Yes") {
            // Set the added trigger to the status bar message.
            const previousIcon = item.iconPath;
            const status = this.indicateChange("Removing owner...", item);
            this._graphClient.removeApplicationOwner(item.objectId!, item.userId!)
                .then(() => {
                    this._onComplete.fire({ success: true, statusBarHandle: status });
                })
                .catch((error) => {
                    this._onError.fire({ success: false, statusBarHandle: status, error: error, treeViewItem: item, previousIcon: previousIcon });
                });
        }
    }

    // Opens the user in the Azure Portal.
    public openInPortal(item: AppRegItem): void {
        env.openExternal(Uri.parse(`${portalUserUri}${item.userId}`));
    }

    // Validates the owner name or email address.
    private async validateOwner(owner: string, existing: User[]): Promise<string | undefined> {

        this._userList = undefined;
        let identifier: string = "";
        if (owner.indexOf('@') > -1) {
            // Try to find the user by email.
            this._userList = await this._graphClient.findUserByEmail(owner);
            identifier = "user with an email address";
        } else {
            // Try to find the user by name.
            this._userList = await this._graphClient.findUserByName(owner);
            identifier = "name";
        }

        if (this._userList.value.length === 0) {
            // User not found
            return `No ${identifier} beginning with ${owner} can be found in your directory.`;
        } else if (this._userList.value.length > 1) {
            // More than one user found
            return `More than one user with the ${identifier} beginning with ${owner} exists in your directory.`;
        }

        // Check if the user is already an owner.
        for (let i = 0; i < existing.length; i++) {
            if (existing[i].id === this._userList.value[0].id) {
                return `${owner} is already an owner of this application.`;
            }
        }

        return undefined;
    }
}