/* eslint-disable @typescript-eslint/naming-convention */

export const window = {
	setStatusBarMessage: jest.fn(() => {
		return { dispose: jest.fn() };
	}),
	showOpenDialog: jest.fn(),
	showInformationMessage: jest.fn(),
	showErrorMessage: jest.fn(),
	showWarningMessage: jest.fn(),
	showQuickPick: jest.fn(),
	showTextDocument: async () => { return true; },
	showInputBox: jest.fn(() => {
		return {
			validateInput: jest.fn()
		};
	})
};

export const workspace = {
	getConfiguration: (_value: string) => ({
		get: (key: string) => {
			switch (key) {
				case "applicationListView":
					return "Owned Applications";
				case "useEventualConsistency":
					return false;
				case "showApplicationCountWarning":
					return false;
				case "includeEntraPortal":
					return true;
				case "omitTenantIdFromPortalRequests":
					return false;
				default:
					return undefined;
			}
		}
	}),
	TextDocument: () => { return true; },
	openTextDocument: async (_options?: { _language?: string; _content?: string }) => { return true; },
	registerTextDocumentContentProvider: () => {
		return new Disposable();
	}
};

export class EventEmitter {
	event = jest.fn();
	fire() {
		return jest.fn();
	}
	dispose() {
		return jest.fn();
	}
}

export const Uri = {
	parse: (value: string, _strict?: boolean) => {
		return { path: value, scheme: "https", fragment: "", query: "", fsPath: value };
	},
	file: (value: string) => {
		return { path: value, scheme: "file", fragment: "", query: "", fsPath: value };
	}
};

let clipboard: any;

export const env = {
	openExternal: async (_value: string) => {
		return true;
	},
	clipboard: {
		writeText: (item: string) => {
			clipboard = item;
		},
		readText: () => {
			return clipboard;
		}
	}
};

export const Event = jest.fn();

export const ThemeColor = jest.fn();

export const ThemeIcon = jest.fn();

export const TreeItem = jest.fn();

export const TreeItemCollapsibleState = jest.fn();

export const TreeDataProvider = {
	onDidChangeTreeDataEvent: new EventEmitter()
};

export const ProviderResult = jest.fn();

export const ConfigurationTarget = jest.fn();

export class Disposable {
	dispose(): any {
		return jest.fn();
	}
}

export class TextDocumentContentProvider {
	onDidChangeEmitter = new EventEmitter();
	onDidChange = this.onDidChangeEmitter.event;
	provideTextDocumentContent(): string {
		return "content";
	}
}

export const MessageOptions = jest.fn();

export const MessageItem = jest.fn();

export const vscode = {
	window,
	workspace,
	Event,
	EventEmitter,
	Disposable,
	ThemeColor,
	ThemeIcon,
	TreeItem,
	TreeItemCollapsibleState,
	TreeDataProvider,
	ProviderResult,
	ConfigurationTarget,
	Uri,
	TextDocumentContentProvider,
	MessageOptions,
	MessageItem
};
