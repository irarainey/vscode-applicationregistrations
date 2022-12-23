import { window, ThemeIcon, env, Uri, Disposable } from 'vscode';
import { portalUserUri } from '../constants';
import { GraphClient } from '../clients/graph';
import { User } from "@microsoft/microsoft-graph-types";
import { AppRegDataProvider } from '../data/applicationRegistration';
import { AppRegItem } from '../models/appRegItem';

export class OwnerService {

    // A private instance of the GraphClient class.
    private graphClient: GraphClient;

    // A private instance of the AppRegDataProvider class.
    private dataProvider: AppRegDataProvider;

    // The constructor for the ApplicationRegistrations class.
    constructor(graphClient: GraphClient, dataProvider: AppRegDataProvider) {
        this.graphClient = graphClient;
        this.dataProvider = dataProvider;
    }

    // Adds a new owner to an application registration.
    public async add(item: AppRegItem): Promise<Disposable | undefined> {

        // Set the created trigger default to undefined.
        let added = undefined;

        // Prompt the user for the new owner.
        const newOwner = await window.showInputBox({
            placeHolder: "Enter user name or email address...",
            prompt: "Add new owner to application"
        });

        // If the new owner name is not empty then add as an owner.
        if (newOwner !== undefined) {
            let userList: User[] = [];
            let identifier: string = "";
            if (newOwner.indexOf('@') > -1) {
                // Try to find the user by email.
                userList = await this.graphClient.findUserByEmail(newOwner);
                identifier = "email address";
            } else {
                // Try to find the user by name.
                userList = await this.graphClient.findUserByName(newOwner);
                identifier = "name";
            }

            if (userList.length === 0) {
                // User not found
                window.showErrorMessage(`No user with the ${identifier} ${newOwner} was found in your directory.`);
            } else if (userList.length > 1) {
                // More than one user found
                window.showErrorMessage(`More than one user with the ${identifier} ${newOwner} has been found in your directory.`);
            } else {
                // Sweet spot
                added = window.setStatusBarMessage("$(loading~spin) Adding owner...");
                item.iconPath = new ThemeIcon("loading~spin");
                this.dataProvider.triggerOnDidChangeTreeData();
                await this.graphClient.addApplicationOwner(item.objectId!, userList[0].id!)
                    .catch((error) => {
                        console.error(error);
                    });
            }
        }

        return added;
    }

    // Removes an owner from an application registration.
    public async remove(item: AppRegItem): Promise<Disposable | undefined> {

        // Set the removed trigger default to undefined.
        let removed = undefined;

        // Prompt the user to confirm the removal.
        const answer = await window.showInformationMessage(`Do you want to remove ${item.label} as an owner of this application?`, "Yes", "No");

        // If the user confirms the removal then remove the user.
        if (answer === "Yes") {
            removed = window.setStatusBarMessage("$(loading~spin) Removing owner...");
            item.iconPath = new ThemeIcon("loading~spin");
            this.dataProvider.triggerOnDidChangeTreeData();
            await this.graphClient.removeApplicationOwner(item.objectId!, item.userId!)
                .catch((error) => {
                    console.error(error);
                });
        }

        return removed;
    }

    // Opens the user in the Azure Portal.
    public openInPortal(user: AppRegItem): void {
        env.openExternal(Uri.parse(`${portalUserUri}${user.userId}`));
    }
}