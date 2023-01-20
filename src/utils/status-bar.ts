import { Disposable, window } from "vscode";
import { v4 as uuidv4 } from "uuid";

export let statusBarItems: [string, Disposable][] = [];

export const setStatusBarMessage = (message: string, spinner: boolean = true): string => {
    const handle = window.setStatusBarMessage(`${spinner === true ? "$(loading~spin)  " : ""}${message}`);
    const id = uuidv4();
    statusBarItems.push([id, handle]);
    return id;
};

export const clearStatusBarMessage = (id: string): void => {
    const index = statusBarItems.findIndex((item) => item[0] === id);
    if (index !== -1) {
        statusBarItems[index][1].dispose();
        statusBarItems.splice(index, 1);
    }
};

export const clearAllStatusBarMessages = (): void => {
    statusBarItems.forEach((item) => item[1].dispose());
    statusBarItems = [];
};