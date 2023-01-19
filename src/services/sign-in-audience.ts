import { window, ThemeIcon, ThemeColor } from "vscode";
import { SIGNIN_AUDIENCE_OPTIONS } from "../constants";
import { AppRegTreeDataProvider } from "../data/app-reg-tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { ServiceBase } from "./service-base";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { GraphResult } from "../types/graph-result";
import { Application } from "@microsoft/microsoft-graph-types";

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
            if (item.contextValue! === "AUDIENCE-PARENT") {
                item.children![0].iconPath = new ThemeIcon("loading~spin");
            } else {
                item.iconPath = new ThemeIcon("loading~spin");
            }
            this.treeDataProvider.triggerOnDidChangeTreeData(item);
            const status = window.setStatusBarMessage(`$(loading~spin) Updating Sign In Audience`);
            const update: GraphResult<void> = await this.graphRepository.updateApplication(item.objectId!, { signInAudience: audience.value });
            if (update.success === true) {
                this.triggerOnComplete({ success: true, statusBarHandle: status });
            } else {
                this.triggerOnError({ success: false, statusBarHandle: status, error: update.error, treeViewItem: item.contextValue! === "AUDIENCE-PARENT" ? item.children![0] : item, previousIcon: new ThemeIcon("symbol-field", new ThemeColor("editor.foreground")), treeDataProvider: this.treeDataProvider });
            }
        }
    }
}