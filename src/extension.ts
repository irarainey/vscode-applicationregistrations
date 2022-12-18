import * as vscode from 'vscode';
import { ApplicationRegistrations } from './applicationRegistrations';

export async function activate(context: vscode.ExtensionContext) {
	new ApplicationRegistrations(context);
}
