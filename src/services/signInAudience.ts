import { Disposable, window, ThemeIcon, env, Uri } from 'vscode';
import { signInAudienceOptions, signInAudienceDocumentation } from '../constants';
import { GraphClient } from '../clients/graph';
import { AppRegDataProvider } from '../data/applicationRegistration';
import { AppRegItem } from '../models/appRegItem';
import { convertSignInAudience } from '../utils/signInAudienceUtils';

export class SignInAudienceService {

    // A private instance of the GraphClient class.
    private graphClient: GraphClient;

    // A private instance of the AppRegDataProvider class.
    private dataProvider: AppRegDataProvider;

    // The constructor for the ApplicationRegistrations class.
    constructor(graphClient: GraphClient, dataProvider: AppRegDataProvider) {
        this.graphClient = graphClient;
        this.dataProvider = dataProvider;
    }

    // Edits the application sign in audience.
    public async edit(item: AppRegItem): Promise<Disposable | undefined> {

        // Set the edited trigger default to undefined.
        let edited = undefined;

        const audience = await window.showQuickPick(signInAudienceOptions, {
            placeHolder: "Select the sign in audience...",
        });

        if (audience !== undefined) {
            // Update the application.
            edited = window.setStatusBarMessage("$(loading~spin) Updating sign in audience...");
            if (item.contextValue! === "AUDIENCE-PARENT") {
                item.children![0].iconPath = new ThemeIcon("loading~spin");
            } else {
                item.iconPath = new ThemeIcon("loading~spin");
            }
            this.dataProvider.triggerOnDidChangeTreeData();
            await this.graphClient.updateApplication(item.objectId!, { signInAudience: convertSignInAudience(audience) })
                .catch(() => {
                    // If the application is not updated then show an error message and a link to the documentation.
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

        return edited;
    }
}