import { Event, EventEmitter, Disposable, ThemeIcon } from "vscode";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { AppRegTreeDataProvider } from "../data/app-reg-tree-data-provider";
import { ErrorResult } from "../types/error-result";
import { AppRegItem } from "../models/app-reg-item";
import { setStatusBarMessage, clearAllStatusBarMessages } from "../utils/status-bar";

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
    private onErrorEvent: EventEmitter<ErrorResult> = new EventEmitter<ErrorResult>();

    // A protected instance of the EventEmitter class to handle complete events.
    private onCompleteEvent: EventEmitter<string | undefined> = new EventEmitter<string | undefined>();

    // A public readonly property to expose the error event.
    public readonly onError: Event<ErrorResult> = this.onErrorEvent.event;

    // A public readonly property to expose the complete event.
    public readonly onComplete: Event<string | undefined> = this.onCompleteEvent.event;

    // The constructor for the OwnerService class.
    constructor(graphRepository: GraphApiRepository, treeDataProvider: AppRegTreeDataProvider) {
        this.graphRepository = graphRepository;
        this.treeDataProvider = treeDataProvider;
    }

    // Trigger the event to indicate an error
    protected triggerOnError(error?: Error) {
        this.onErrorEvent.fire({ error: error, item: this.item, treeDataProvider: this.treeDataProvider });
    }

    // Trigger the event to indicate completion
    protected triggerOnComplete(statusId: string | undefined = undefined) {
        this.onCompleteEvent.fire(statusId);
    }

    // Initiates the visual change of the tree view
    protected indicateChange(statusMessage?: string, item?: AppRegItem): string | undefined {

        let id: string | undefined = undefined;

        if (statusMessage !== undefined) {
            id = setStatusBarMessage(statusMessage);
        }

        if (item !== undefined) {
            this.item = item;
            item.iconPath = new ThemeIcon("loading~spin");
            this.treeDataProvider.triggerOnDidChangeTreeData(item);
        }

        return id;
    }

    // Resets the icon for a tree item
    protected resetTreeItemIcon(item: AppRegItem) {
        item.iconPath = item.baseIcon;
        this.treeDataProvider.triggerOnDidChangeTreeData(item);
    }

    // Dispose of anything that needs to be disposed of.
    dispose(): void {
        clearAllStatusBarMessages();
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