import { window } from "vscode";
import { setStatusBarMessage } from "./status-bar";
import { errorHandler } from "../error-handler";
import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import { AccountProvider } from "../types/account-provider";

// Authenticate the user.
export const signInUser = async (treeDataProvider: AppRegTreeDataProvider, accountProvider: AccountProvider) => {
	await treeDataProvider.render(undefined, "AUTHENTICATING");
	const status = await authenticate(accountProvider);

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

// Uses the AccountProvider instance to sign out the current user.
export const signOutUser = async (treeDataProvider: AppRegTreeDataProvider, accountProvider: AccountProvider) => {
	await accountProvider
		.logoutUser()
		.then(async () => {
			await treeDataProvider.render(undefined, "SIGN-IN");
		})
		.catch(async (error) => {
			await errorHandler(error);
		});
};

// Uses the AccountProvider instance to authenticate the user.
const authenticate = async (accountProvider: AccountProvider): Promise<boolean | undefined> => {
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

	const status = await accountProvider.loginUser(tenant);
	return status;
};
