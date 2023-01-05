import { window, ThemeIcon, env, Uri, ThemeColor } from 'vscode';
import { signInAudienceDocumentation, signInAudienceOptions } from '../constants';
import { AppRegTreeDataProvider } from '../data/appRegTreeDataProvider';
import { AppRegItem } from '../models/appRegItem';
import { ServiceBase } from './serviceBase';
import { GraphClient } from '../clients/graph';

export class SignInAudienceService extends ServiceBase {

    // The constructor for the SignInAudienceService class.
    constructor(treeDataProvider: AppRegTreeDataProvider, graphClient: GraphClient) {
        super(treeDataProvider, graphClient);
    }

    // Edits the application sign in audience.
    public async edit(item: AppRegItem): Promise<void> {

        const audience = await window.showQuickPick(
            signInAudienceOptions,
            {
                placeHolder: "Select the sign in audience...",
                ignoreFocusOut: true
            });

        if (audience !== undefined) {
            if (item.contextValue! === "AUDIENCE-PARENT") {
                item.children![0].iconPath = new ThemeIcon("loading~spin");
            } else {
                item.iconPath = new ThemeIcon("loading~spin");
            }
            this.dataProvider.triggerOnDidChangeTreeData(item);
            const status = window.setStatusBarMessage(`$(loading~spin) Updating sign in audience...`);

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
                    this.dataProvider.triggerOnDidChangeTreeData(item);
                    status.dispose();

                    window.showErrorMessage(
                        `An error occurred while attempting to change the sign in audience. This is likely because some properties of the application are not supported by the new sign in audience. Please consult the Azure AD documentation for more information at ${signInAudienceDocumentation}.`,
                        ...["OK", "Open Documentation"]
                    )
                        .then((answer) => {
                            if (answer === "Open Documentation") {
                                env.openExternal(Uri.parse(signInAudienceDocumentation));
                            }
                        });
                });
        }
    }
}