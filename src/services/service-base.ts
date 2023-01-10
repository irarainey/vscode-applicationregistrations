import { Event, EventEmitter, Disposable, ThemeIcon, window, Uri } from "vscode";
import { GraphClient } from "../clients/graph-client";
import { AppRegTreeDataProvider } from "../data/app-reg-tree-data-provider";
import { ActivityResult } from "../interfaces/activity-result";
import { AppRegItem } from "../models/app-reg-item";

export class ServiceBase {

    // A protected array of disposables for the service.
    protected readonly disposable: Disposable[] = [];

    // A private instance of the AppRegTreeDataProvider class.
    protected readonly treeDataProvider: AppRegTreeDataProvider | undefined;

    // A protected instance of the GraphClient class.
    protected readonly graphClient: GraphClient;

    // A protected instance of the EventEmitter class to handle error events.
    private onErrorEvent: EventEmitter<ActivityResult> = new EventEmitter<ActivityResult>();

    // A protected instance of the EventEmitter class to handle complete events.
    private onCompleteEvent: EventEmitter<ActivityResult> = new EventEmitter<ActivityResult>();

    // A public readonly property to expose the error event.
    public readonly onError: Event<ActivityResult> = this.onErrorEvent.event;

    // A public readonly property to expose the complete event.
    public readonly onComplete: Event<ActivityResult> = this.onCompleteEvent.event;

    // The constructor for the OwnerService class.
    constructor(graphClient: GraphClient, treeDataProvider: AppRegTreeDataProvider | undefined = undefined) {
        this.graphClient = graphClient;
        if(treeDataProvider === undefined) {
            this.treeDataProvider = treeDataProvider;
        }
    }

    // Trigger the event to indicate an error
    protected triggerOnError(item: ActivityResult) {
        this.onErrorEvent.fire(item);
    }

    // Trigger the event to indicate completion
    protected triggerOnComplete(item: ActivityResult) {
        this.onCompleteEvent.fire(item);
    }

    // Initiates the visual change of the tree view
    protected triggerTreeChange(statusBarMessage?: string, item?: AppRegItem): Disposable | undefined {
        if (item !== undefined) {
            item.iconPath = new ThemeIcon("loading~spin");
            this.treeDataProvider!.triggerOnDidChangeTreeData(item);
        }

        if (statusBarMessage !== undefined) {
            return window.setStatusBarMessage(`$(loading~spin) ${statusBarMessage}`);
        }
    }

    // Sets the icon for a tree item
    protected setTreeItemIcon(item: AppRegItem, icon?: string | Uri | {
        light: string | Uri;
        dark: string | Uri;
    } | ThemeIcon | undefined, spinner?: boolean) {

        if (spinner) {
            icon = new ThemeIcon("loading~spin");
        }
        item.iconPath = icon;
        this.treeDataProvider!.triggerOnDidChangeTreeData(item);
    }

    // Dispose of anything that needs to be disposed of.
    dispose(): void {
        this.onErrorEvent.dispose();
        this.onCompleteEvent.dispose();
        while (this.disposable.length) {
            const x = this.disposable.pop();
            if (x) {
                x.dispose();
            }
        }
    }
}