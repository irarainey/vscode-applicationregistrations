import { window } from "vscode";
import { execShellCmd } from "../utils/exec-shell-cmd";
import { setStatusBarMessage } from "../utils/status-bar";
import { errorHandler } from "../error-handler";
import { AppRegTreeDataProvider } from "../data/tree-data-provider";

// Invokes the Azure CLI sign-in command to authenticate the user.
export const signInToCli = async (treeDataProvider: AppRegTreeDataProvider) => {
	await treeDataProvider.render(undefined, "AUTHENTICATING");
	const status = await authenticate();

	// The user pressed cancel.
	if (status === undefined) {
		// Set the tree view to initialising.
		await treeDataProvider.render(undefined, "INITIALISING");

		// Initialise the graph repository.
		const status = await treeDataProvider.graphRepository.initialise();

		// Were we able to get an access token?
		if (status !== true) {
			// If not then the user needs to authenticate.
			await treeDataProvider.render(undefined, "SIGN-IN");
		} else {
			// If so then the user is authenticated so load the tree view.
			await treeDataProvider.render(setStatusBarMessage("Loading Application Registrations..."));
		}

		return;
	}

	// If there was a problem authenticating then show the sign in option.
	if (status !== true) {
		await treeDataProvider.render(undefined, "SIGN-IN");
		return;
	}

	// If the user was successfully authenticated then load the list
	await treeDataProvider.render(setStatusBarMessage("Loading Application Registrations..."), "AUTHENTICATED");
};

// Invokes the Azure CLI sign-out command to sign the user out.
export const signOutFromCli = async (treeDataProvider: AppRegTreeDataProvider) => {
	const status = await execShellCmd("az logout")
		.then(async () => {
			await treeDataProvider.render(undefined, "SIGN-IN");
		})
		.catch(async (error) => {
			await errorHandler(error);
		});
};

// Invokes the Azure CLI sign-in command to authenticate the user.
const authenticate = async (): Promise<boolean | undefined> => {
	// Prompt the user for the tenant name or Id.
	const tenant = await window.showInputBox({
		placeHolder: "Tenant Name or ID",
		prompt: "Enter the tenant name or ID, or leave blank for the default tenant",
		title: "Azure CLI Sign-In Tenant",
		ignoreFocusOut: true
	});

	// If the tenant is undefined then we don't want to do anything because they pressed cancel.
	if (tenant === undefined) {
		return undefined;
	}

	// Build the command to invoke the Azure CLI sign-in command.
	let command = "az login";
	if (tenant.length > 0) {
		command += ` --tenant ${tenant}`;
	}

	// Execute the command.
	const status = await execShellCmd(command)
		.then(() => {
			return true;
		})
		.catch(() => {
			return false;
		});

	return status;
};
