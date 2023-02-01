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
  getConfiguration: jest.fn()
};

export class EventEmitter {
  constructor() {}
  fire() { return jest.fn(); }
};

export const Event = jest.fn();

export const ThemeIcon = jest.fn();

export const TreeItem = jest.fn();

export const TreeItemCollapsibleState = jest.fn();

export const vscode = {
  window,
  workspace,
  Event,
  EventEmitter,
  ThemeIcon,
  TreeItem,
  TreeItemCollapsibleState
};