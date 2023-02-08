import { env } from "vscode";
import { AppRegItem } from "../models/app-reg-item";

// Define the copy value function.
export const copyValue = (item: AppRegItem) => {
    env.clipboard.writeText(
        item.contextValue === "COPY"
            || item.contextValue === "WEB-REDIRECT-URI"
            || item.contextValue === "SPA-REDIRECT-URI"
            || item.contextValue === "NATIVE-REDIRECT-URI"
            || item.contextValue === "APPID-URI"
            || item.contextValue === "LOGOUT-URL"
            ? item.value!
            : item.children![0].value!);
};