/* eslint-disable @typescript-eslint/naming-convention */
export const window = {
  createStatusBarItem: jest.fn(() => ({
    show: jest.fn()
  })),
  setStatusBarMessage: jest.fn(),
  showInformationMessage: jest.fn().mockResolvedValue("YES"),
  showErrorMessage: jest.fn(),
  showWarningMessage: jest.fn(),
  showQuickPick: jest.fn().mockResolvedValue({
    label: "Single Tenant",
    description: "Accounts in this organizational directory only.",
    value: "AzureADMyOrg"
  })
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

export const vscode = {
  window,
  workspace,
  Event,
  EventEmitter,
  ThemeColor,
  ThemeIcon,
  TreeItem,
  TreeItemCollapsibleState,
  TreeDataProvider,
  ProviderResult,
  ConfigurationTarget
};
