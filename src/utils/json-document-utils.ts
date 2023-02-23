import { window, Uri, workspace, Disposable } from "vscode";
import { clearStatusBarMessage } from "./status-bar";
import { v4 as uuidv4 } from "uuid";
import { JsonDocument } from "../models/json-document";

export const showJsonDocument = async (title: string, content: any, status?: string | undefined): Promise<Disposable> => {
	// Set a unique content provider for each document to avoid conflicts
	const contentProvider = uuidv4();

	// Register the content provider and open the document
	const providerHandle = workspace.registerTextDocumentContentProvider(contentProvider, new JsonDocument(content));

	// Open the document
	const document = await workspace.openTextDocument(Uri.parse(`${contentProvider}:${title}.json`));

	// Show the document
	await window.showTextDocument(document, { preview: false }).then(() => {
		// Clear the status bar message if it was set
		if (status !== undefined) {
			clearStatusBarMessage(status!);
		}
	});

	// Return the provider handle to allow the caller to dispose of it
	return providerHandle;
};
