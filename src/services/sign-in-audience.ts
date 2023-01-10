import { window, ThemeIcon, env, Uri, ThemeColor } from "vscode";
import { SIGNIN_AUDIENCE_DOCUMENTATION_URI, SIGNIN_AUDIENCE_OPTIONS } from "../constants";
import { AppRegTreeDataProvider } from "../data/app-reg-tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { ServiceBase } from "./service-base";
import { GraphClient } from "../clients/graph-client";

export class SignInAudienceService extends ServiceBase {

    // The constructor for the SignInAudienceService class.
    constructor(graphClient: GraphClient, treeDataProvider: AppRegTreeDataProvider) {
        super(graphClient, treeDataProvider);
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
            this.treeDataProvider!.triggerOnDidChangeTreeData(item);
            const status = window.setStatusBarMessage(`$(loading~spin) Updating Sign In Audience`);

            this.graphClient.updateApplication(item.objectId!, { signInAudience: audience.value })
                .then(() => {
                    this.triggerOnComplete({ success: true, statusBarHandle: status });
                })
                .catch(() => {
                    if (item.contextValue! === "AUDIENCE-PARENT") {
                        item.children![0].iconPath = new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"));
                    } else {
                        item.iconPath = new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"));
                    }
                    this.treeDataProvider!.triggerOnDidChangeTreeData(item);
                    status.dispose();

                    window.showErrorMessage(
                        `An error occurred while attempting to change the Sign In Audience. This is likely because some properties of the application are not supported by the new sign in audience. Please consult the Azure AD documentation for more information at ${SIGNIN_AUDIENCE_DOCUMENTATION_URI}.`,
                        ...["OK", "Open Documentation"]
                    )
                        .then((answer) => {
                            if (answer === "Open Documentation") {
                                env.openExternal(Uri.parse(SIGNIN_AUDIENCE_DOCUMENTATION_URI));
                            }
                        });
                });
        }
    }
}