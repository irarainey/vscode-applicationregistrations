import { commands, ExtensionContext, window } from 'vscode';
import { view } from './constants';
import { AppRegDataProvider } from './data/applicationRegistration';
import { GraphClient } from './clients/graph';
import { AppReg } from './applicationRegistrations';

// This method is called when the extension is activated.
export async function activate(context: ExtensionContext) {

	// Create a new instance of the GraphClient class.
	const graphClient = new GraphClient();

	// Create a new instance of the ApplicationDataProvider class.
	const dataProvider = new AppRegDataProvider();

	// Create a new instance of the ApplicationRegistrations class.
	const appReg = new AppReg(graphClient, dataProvider, context);

	// Register the commands.
	commands.registerCommand(`${view}.signInToAzure`, () => appReg.cliSignIn());
	commands.registerCommand(`${view}.refreshApps`, () => appReg.loadTreeView(window.setStatusBarMessage("$(loading~spin) Refreshing Application Registrations...")));
	commands.registerCommand(`${view}.filterApps`, () => appReg.filterTreeView());
	commands.registerCommand(`${view}.addApp`, () => appReg.addApp());
	commands.registerCommand(`${view}.deleteApp`, app => appReg.deleteApp(app));
	commands.registerCommand(`${view}.renameApp`, app => appReg.renameApp(app));
	commands.registerCommand(`${view}.copyAppId`, app => appReg.copyAppId(app));
	commands.registerCommand(`${view}.openAppInPortal`, app => appReg.openAppInPortal(app));
	commands.registerCommand(`${view}.viewAppManifest`, app => appReg.viewAppManifest(app));
	commands.registerCommand(`${view}.editAudience`, app => appReg.editAudience(app));
	commands.registerCommand(`${view}.addOwner`, user => appReg.addOwner(user));
	commands.registerCommand(`${view}.removeOwner`, user => appReg.removeOwner(user));
	commands.registerCommand(`${view}.openUserInPortal`, user => appReg.openUserInPortal(user));
	commands.registerCommand(`${view}.addRedirectUri`, app => appReg.addRedirectUri(app));
	commands.registerCommand(`${view}.editRedirectUri`, app => appReg.editRedirectUri(app));
	commands.registerCommand(`${view}.deleteRedirectUri`, app => appReg.deleteRedirectUri(app));
	commands.registerCommand(`${view}.copyValue`, app => appReg.copyValue(app));

	// Register the tree view data provider.
	window.registerTreeDataProvider(view, dataProvider);

	// Set the initial state of the tree view.
	dataProvider.initialise("LOADING", window.setStatusBarMessage("$(loading~spin) Initialising Application Registration Extension..."));
}
