import { window, ThemeIcon, env, Uri, ThemeColor } from 'vscode';
import { signInAudienceOptions, signInAudienceDocumentation } from '../constants';
import { AppRegDataProvider } from '../data/applicationRegistration';
import { AppRegItem } from '../models/appRegItem';
import { convertSignInAudience } from '../utils/signInAudienceUtils';
import { ServiceBase } from './serviceBase';

export class SignInAudienceService extends ServiceBase {

    // The constructor for the SignInAudienceService class.
    constructor(dataProvider: AppRegDataProvider) {
        super(dataProvider);
    }

    // Edits the application sign in audience.
    public async edit(item: AppRegItem): Promise<void> {

        const audience = await window.showQuickPick(signInAudienceOptions, {
            placeHolder: "Select the sign in audience...",
            ignoreFocusOut: true
        });

        if (audience !== undefined) {
            if (item.contextValue! === "AUDIENCE-PARENT") {
                item.children![0].iconPath = new ThemeIcon("loading~spin");
            } else {
                item.iconPath = new ThemeIcon("loading~spin");
            }
            this._dataProvider.triggerOnDidChangeTreeData();
            const status = window.setStatusBarMessage(`$(loading~spin) Updating sign in audience...`)

            this._graphClient.updateApplication(item.objectId!, { signInAudience: convertSignInAudience(audience) })
                .then(() => {
                    this._onComplete.fire({ success: true, statusBarHandle: status });
                })
                .catch(() => {
                    if (item.contextValue! === "AUDIENCE-PARENT") {
                        item.children![0].iconPath = new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"));
                    } else {
                        item.iconPath = new ThemeIcon("symbol-field", new ThemeColor("editor.foreground"));
                    }
                    this._dataProvider.triggerOnDidChangeTreeData();
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