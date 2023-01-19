import { window, env, Uri } from "vscode";
import { SIGNIN_AUDIENCE_DOCUMENTATION_URI } from "./constants";
import { ActivityResult } from "./types/activity-result";

// Define the error handler function.
export const errorHandler = async (result: ActivityResult) => {

	if (result.statusBarHandle !== undefined) {
		// Clear any status bar messages.
		result.statusBarHandle!.dispose();
	}

	if (result.item !== undefined && result.treeDataProvider !== undefined) {
		// Restore the original icon.
		result.item.iconPath = result.item.baseIcon;
		result.treeDataProvider.triggerOnDidChangeTreeData(result.item);
	}

	if (result.error !== undefined) {
		// Log the error.
		console.error(result.error);

		// Determine if the error is due to the user not being logged in.
		if (result.error.message.includes("az login") || result.error.message.includes("az account set")) {
			if (result.treeDataProvider !== undefined) {
				window.showErrorMessage("You are not logged in to the Azure CLI. Please click the option to sign in, or run 'az login' in a terminal window.", "OK");
				await result.treeDataProvider.initialiseGraphClient();
				return;
			} else {
				window.showErrorMessage("You are not logged in to the Azure CLI. Please run 'az login' in a terminal window.", "OK");
				return;
			}
		}

		// Determine if the error is due to trying to change the sign in audience.
		if (result.error.message.includes("signInAudience")) {
			const result = await window.showErrorMessage(
				`An error occurred while attempting to change the Sign In Audience. This is likely because some properties of the application are not supported by the new sign in audience. Please consult the Azure AD documentation for more information at ${SIGNIN_AUDIENCE_DOCUMENTATION_URI}.`,
				...["OK", "Open Documentation"]
			);

			if (result === "Open Documentation") {
				env.openExternal(Uri.parse(SIGNIN_AUDIENCE_DOCUMENTATION_URI));
				return;
			}

			return;
		}

		// Display an error message.
		window.showErrorMessage(`An error occurred trying to complete your task: ${result.error!.message}.`, "OK");
	}
};
