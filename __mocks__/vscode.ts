/* eslint-disable @typescript-eslint/naming-convention */
export const window = {
  setStatusBarMessage: jest.fn(() => {
    return { dispose: jest.fn() };
  }),
  showInformationMessage: jest.fn(),
  showErrorMessage: jest.fn(),
  showWarningMessage: jest.fn(),
  showQuickPick: jest.fn()
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
  })
};

export class EventEmitter {
  fire() { return jest.fn(); }
};

export const Event = jest.fn();

export const ThemeColor = jest.fn();

export const ThemeIcon = jest.fn();

export const TreeItem = jest.fn();

export const TreeItemCollapsibleState = jest.fn();

export const TreeDataProvider = jest.fn();

export const ProviderResult = jest.fn();

export const ConfigurationTarget = jest.fn();

export const Disposable = jest.fn();

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
  ConfigurationTarget
};
