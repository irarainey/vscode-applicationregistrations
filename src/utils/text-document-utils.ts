import * as vscode from "vscode";
import { window, Uri, workspace, Disposable, TextDocumentContentProvider } from "vscode";
import { clearStatusBarMessage } from "./status-bar";
import { v4 as uuidv4 } from "uuid";

export const showJsonDocument = async (title: string, content: any, status?: string | undefined): Promise<Disposable> => {
	const newDocument = new (class implements TextDocumentContentProvider {
		onDidChangeEmitter = new vscode.EventEmitter<Uri>();
		onDidChange = this.onDidChangeEmitter.event;
		provideTextDocumentContent(): string {
			return content;
		}
	})();

	const contentProvider = uuidv4();
	const providerHandle = workspace.registerTextDocumentContentProvider(contentProvider, newDocument);
	const uri = Uri.parse(`${contentProvider}:${title}.json`);
	workspace.openTextDocument(uri).then(async (doc) => {
		await window.showTextDocument(doc, { preview: false });
		if (status !== undefined) {
			clearStatusBarMessage(status!);
		}
	});

	return providerHandle;
};
