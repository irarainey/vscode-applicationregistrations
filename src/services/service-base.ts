import { Disposable, ThemeIcon } from "vscode";
import { GraphApiRepository } from "../repositories/graph-api-repository";
import { AppRegTreeDataProvider } from "../data/tree-data-provider";
import { AppRegItem } from "../models/app-reg-item";
import { setStatusBarMessage, clearAllStatusBarMessages } from "../utils/status-bar";
import { errorHandler } from "../error-handler";
import { Application } from "@microsoft/microsoft-graph-types";
import { GraphResult } from "../types/graph-result";

export class ServiceBase {
	// A protected array of disposables for the service.
	protected readonly disposable: Disposable[] = [];

	// A protected instance of the tree item.
	protected item: AppRegItem | undefined = undefined;

	// A protected instance of the previous status bar handle.
	protected statusBarHandle: Disposable | undefined = undefined;

	// The constructor for the OwnerService class.
	constructor(protected graphRepository: GraphApiRepository, protected treeDataProvider: AppRegTreeDataProvider) {
	}

	// Handle an error
	protected async handleError(error?: Error, source?: string | undefined) {
		await errorHandler({ error: error, item: this.item, treeDataProvider: this.treeDataProvider, source: source });
	}

	// Trigger completion by refreshing the tree
	protected async triggerRefresh(statusId: string | undefined = undefined) {
		await this.treeDataProvider.render(statusId);
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

	// Updates the application registration.
	protected async updateApplication(id: string, application: Application, status: string | undefined = undefined, source?: string | undefined): Promise<void> {
		const update: GraphResult<void> = await this.graphRepository.updateApplication(id, application);
		update.success === true ? await this.triggerRefresh(status) : await this.handleError(update.error, source);
	}

	// Resets the icon for a tree item
	protected resetTreeItemIcon(item: AppRegItem) {
		item.iconPath = item.baseIcon;
		this.treeDataProvider.triggerOnDidChangeTreeData(item);
	}

	// Dispose of anything that needs to be disposed of.
	dispose(): void {
		clearAllStatusBarMessages();
		while (this.disposable.length) {
			const x = this.disposable.pop();
			if (x) {
				x.dispose();
			}
		}
	}
}
