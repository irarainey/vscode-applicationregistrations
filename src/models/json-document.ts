import { TextDocumentContentProvider } from "vscode";

export class JsonDocument implements TextDocumentContentProvider {
	constructor(private content: any) {}

	provideTextDocumentContent(): string {
		return JSON.stringify(this.content, null, 4);
	}
}