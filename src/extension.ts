import * as vscode from 'vscode';
import { view } from './constants';
import { AppRegDataProvider } from './appRegDataProvider';
import { GraphClient } from './graphClient';
import { ApplicationRegistrations } from './applicationRegistrations';

// This method is called when the extension is activated.
export async function activate(context: vscode.ExtensionContext) {

	// Create a new instance of the GraphClient class.
	const graphClient =	new GraphClient();

	// Create a new instance of the ApplicationDataProvider class.
	const dataProvider = new AppRegDataProvider();

	// Create a new instance of the ApplicationRegistrations class.
	const appReg = new ApplicationRegistrations(graphClient, dataProvider, context);

	// Register the commands.
	vscode.commands.registerCommand(`${view}.addApp`, () => appReg.addApp());
	vscode.commands.registerCommand(`${view}.deleteApp`, app => appReg.deleteApp(app));
	vscode.commands.registerCommand(`${view}.renameApp`, app => appReg.renameApp(app));
	vscode.commands.registerCommand(`${view}.refreshApps`, () => appReg.populateTreeView());
	vscode.commands.registerCommand(`${view}.filterApps`, () => appReg.filterApps());
	vscode.commands.registerCommand(`${view}.viewAppManifest`, app => appReg.viewAppManifest(app));
	vscode.commands.registerCommand(`${view}.copyAppId`, app => appReg.copyAppId(app));
	vscode.commands.registerCommand(`${view}.openAppInPortal`, app => appReg.openAppInPortal(app));
	vscode.commands.registerCommand(`${view}.editAudience`, app => appReg.editAudience(app));
	vscode.commands.registerCommand(`${view}.copyValue`, app => appReg.copyValue(app));
	vscode.commands.registerCommand(`${view}.signInToAzure`, () => appReg.invokeSignIn());
	vscode.commands.registerCommand(`${view}.openUserInPortal`, user => appReg.openUserInPortal(user));
	vscode.commands.registerCommand(`${view}.addOwner`, user => appReg.addOwner(user));
	vscode.commands.registerCommand(`${view}.removeOwner`, user => appReg.removeOwner(user));
	vscode.commands.registerCommand(`${view}.addRedirectUri`, app => appReg.addRedirectUri(app));

	// Register the tree view data provider.
	vscode.window.registerTreeDataProvider(view, dataProvider);

	// Set the initial state of the tree view.
	dataProvider.initialise("LOADING");
}
