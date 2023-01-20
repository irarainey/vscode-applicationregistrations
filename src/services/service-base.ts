import { Event, EventEmitter, Disposable, ThemeIcon, window } from "vscode";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { AppRegTreeDataProvider } from "../data/app-reg-tree-data-provider";
import { ActivityResult } from "../types/activity-result";
import { AppRegItem } from "../models/app-reg-item";

export class ServiceBase {

    // A protected array of disposables for the service.
    protected readonly disposable: Disposable[] = [];

    // A private instance of the AppRegTreeDataProvider class.
    protected readonly treeDataProvider: AppRegTreeDataProvider;

    // A protected instance of the Graph Api Repository.
    protected readonly graphRepository: GraphApiRepository;

    // A protected instance of the tree item.
    protected item: AppRegItem | undefined = undefined;

    // A protected instance of the previous status bar handle.
    protected statusBarHandle: Disposable | undefined = undefined;

    // A protected instance of the EventEmitter class to handle error events.
    private onErrorEvent: EventEmitter<ActivityResult> = new EventEmitter<ActivityResult>();

    // A protected instance of the EventEmitter class to handle complete events.
    private onCompleteEvent: EventEmitter<ActivityResult> = new EventEmitter<ActivityResult>();

    // A public readonly property to expose the error event.
    public readonly onError: Event<ActivityResult> = this.onErrorEvent.event;

    // A public readonly property to expose the complete event.
    public readonly onComplete: Event<ActivityResult> = this.onCompleteEvent.event;

    // The constructor for the OwnerService class.
    constructor(graphRepository: GraphApiRepository, treeDataProvider: AppRegTreeDataProvider) {
        this.graphRepository = graphRepository;
        this.treeDataProvider = treeDataProvider;
    }

    // Trigger the event to indicate an error
    protected triggerOnError(error?: Error) {
        this.onErrorEvent.fire({ success: false, error: error, item: this.item, statusBarHandle: this.statusBarHandle, treeDataProvider: this.treeDataProvider });
    }

    // Trigger the event to indicate completion
    protected triggerOnComplete() {
        this.onCompleteEvent.fire({ success: true, statusBarHandle: this.statusBarHandle, treeDataProvider: this.treeDataProvider });
    }

    // Initiates the visual change of the tree view
    protected triggerTreeChange(statusBarMessage?: string, item?: AppRegItem): void {
        if (item !== undefined) {
            this.item = item;
            item.iconPath = new ThemeIcon("loading~spin");
            this.treeDataProvider.triggerOnDidChangeTreeData(item);
        }

        if (statusBarMessage !== undefined) {
            this.statusBarHandle = window.setStatusBarMessage(`$(loading~spin) ${statusBarMessage}`);
        }
    }

    // Resets the icon for a tree item
    protected resetTreeItemIcon(item: AppRegItem) {
        item.iconPath = item.baseIcon;
        this.treeDataProvider.triggerOnDidChangeTreeData(item);
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