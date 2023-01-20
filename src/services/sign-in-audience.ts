import { window } from "vscode";
import { SIGNIN_AUDIENCE_OPTIONS } from "../constants";
import { AppRegTreeDataProvider } from "../data/app-reg-tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { ServiceBase } from "./service-base";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { GraphResult } from "../types/graph-result";

export class SignInAudienceService extends ServiceBase {

    // The constructor for the SignInAudienceService class.
    constructor(graphRepository: GraphApiRepository, treeDataProvider: AppRegTreeDataProvider) {
        super(graphRepository, treeDataProvider);
    }

    // Edits the application sign in audience.
    async edit(item: AppRegItem): Promise<void> {

        const audience = await window.showQuickPick(
            SIGNIN_AUDIENCE_OPTIONS,
            {
                placeHolder: "Select the Sign In Audience",
                title: "Edit Sign In Audience",
                ignoreFocusOut: true
            });

        if (audience !== undefined) {
            let status: string | undefined; 
            if (item.contextValue! === "AUDIENCE-PARENT") {
                status = this.indicateChange("Updating Sign In Audience...", item.children![0]);
            } else {
                status = this.indicateChange("Updating Sign In Audience...", item);
            }
            const update: GraphResult<void> = await this.graphRepository.updateApplication(item.objectId!, { signInAudience: audience.value });
            update.success === true ? this.triggerOnComplete(status) : this.triggerOnError(update.error);
        }
    }
}