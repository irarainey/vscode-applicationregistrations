/* eslint-disable @typescript-eslint/naming-convention */

export const window = {
  setStatusBarMessage: jest.fn(() => {
    return { dispose: jest.fn() };
  }),
  showInformationMessage: jest.fn(),
  showErrorMessage: jest.fn(),
  showWarningMessage: jest.fn(),
  showQuickPick: jest.fn(),
  showTextDocument: jest.fn()
};

export const workspace = {
  getConfiguration: (value: string) => ({
    get: (key: string) => {
      switch (key) {
        case "useEventualConsistency":
          return true;
        case "showApplicationCountWarning":
          return false;
        case "showOwnedApplicationsOnly":
          return true;
        default:
          return undefined;
      }
    }
  }),
  TextDocument: jest.fn(),
  openTextDocument: () => {
    return new Promise(workspace.TextDocument);
  },
  registerTextDocumentContentProvider: jest.fn()
};

export class EventEmitter {
  fire() { return jest.fn(); }
};

export const Uri = {
  parse: (value: string, strict?: boolean) => {
    jest.fn();
  }
};

let clipboard: any;

export const env = {
  openExternal: (value: string) => {
    jest.fn();
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
  onDidChangeTreeDataEvent: () => ({
    dispose: () => jest.fn()
  })
};

export const ProviderResult = jest.fn();

export const ConfigurationTarget = jest.fn();

export const Disposable = jest.fn();

export class TextDocumentContentProvider {
  provideTextDocumentContent() { return "{ \"test\": \"test\" }"; }
}

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
  TextDocumentContentProvider
};
