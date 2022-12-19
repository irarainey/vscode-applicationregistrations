# Azure Application Registration Management for VS Code

![Extension Banner](resources/images/readme_banner.png)

This Visual Studio Code extension provides an easy way to view and manage Azure Application Registrations outside of the Azure Portal.

It allows for the viewing, copying, and editing of most of the core application properties, such as Application Id; Sign In Audience; Redirect URIs; API Permissions; Exposed API Permissions; App Roles; and Owners.

It also allows for the simple creation of new applications, quickly viewing of the full application manifest, and has the ability to open the application registration directly in the Azure Portal when you need full editing control. Once loaded, the application registration list can also be filtered by display name.

![Application Registration List](resources/images/application_view.png)

## Authentication
This extension uses the `AzureCliCredential` to authenticate the user and gain an access token required to manage applications. This means it does not use the Azure Account identity, but rather the Azure CLI. This is a workaround due to a [known issue](https://learn.microsoft.com/en-us/javascript/api/overview/azure/identity-readme?view=azure-node-latest#note-about-visualstudiocodecredential) with the `VisualStudioCodeCredential` and Azure Account extension >= v0.10.0.

Please ensure your Azure CLI is authenticated to the correct tenant using `az login --tenant <tenant_name or tenant_id>`. If you are not authenticated using the Azure CLI the extension will offer the opportunity to sign in to the tenant of your choice.

![VS Code Sign In](resources/images/signin.png)

The access token used for this extension uses the scope `Directory.AccessAsUser.All`. This means that it will use the Azure RBAC directory roles assigned to the authenticated user, and hence requires a role that allows for application management.

## Notes
This extension was created both as a learning exercise, and to address the common annoyances of managing Application Registrations. It is not officially supported and you use it at your own risk.

It has a dependency on the [Azure Tools extension pack](https://marketplace.visualstudio.com/items?itemName=ms-vscode.vscode-node-azure-pack), but only because it places the application registrations view into the Azure view container.
