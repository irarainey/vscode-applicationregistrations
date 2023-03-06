## 0.12.0

- Restricted when command palette commands are shown to only show when extension panel is visible.
- Added keyboard shortcuts for all top-level commands.

## 0.11.0

- Added functionality to manage multiple Application Id URIs.

## 0.10.0

- Added functionality to view and manage deleted applications.
- Added new command to the command palette and main view toolbar to change the list view.
- Amended logic showing filter and add application functionality to ensure it is only shown when relevant.

## 0.9.0

- Enhanced application tooltip to show additional information including creation date.
- Added visual indication of expired credentials and credentials expiring soon.
- Changed the visual appearance of disabled app roles and exposed api permissions to maintain consistent icons.
- Refactored use of sign in audience to improve performance when performing various validations.

## 0.8.0

- Added functionality to open the user or application in the Entra portal.
- Added the ability to open in the portal using an explicit tenant id.
- Added a warning if a duplicate app role or exposed api scope is detected.
- Refactored redirect uri editing and creation to improve performance.
- Fixed an issue adding API scopes from a new application.
- Fixed an issue validating expiry dates for password credentials.
- Fixed an issue validating App Id Uris for consumer accounts.

## 0.7.1

- Fixed an issue copying the front-channel logout url.

## 0.7.0

- Added commands to sign in and sign out of Azure CLI from the command palette.
- Added functionality to manage the front-channel logout URL of an application.
- Added command filters to ensure those that aren't relevant aren't shown in the palette.
- Refactored tree rendering trigger to facilitate better unit testing.
- Fixed an issue where the wrong error message was passed.

## 0.6.0

- Added functionality to edit individual app role and exposed api scope properties.
- Improved performance of creating and editing exposed api scopes.
- Improved validation around app roles and exposed scopes.
  
## 0.5.0

- Added user and directory role assignment information to tenant information screen.
- Added the ability to delete all redirect uris of a single type at once.
- Added functionality to show application endpoints based upon sign in audience.
- Added the ability to disable and delete app roles and scopes at the same time.
- Changed the order of parameters when creating a new application registration to check name length.
- Improved validation of app roles and scopes when creating and editing.
- Improved various validation based upon different sign in audiences.
- Fixed an issue where the manifest shown might not always been up to date.
- Fixed an issue where the sign in audience might not be displayed correctly.
- Fixed an issue that allowed duplicate app roles to be created.

## 0.4.0

- Added additional error handling to capture Azure CLI being signed out.
- Added a dispose call for the organization service.
- Refactored error and action completion triggers.
- Refactored statusbar message handling and disposal.
- Refactored Azure CLI authentication process.
- Fixed an issue where duplicate redirect uris could be added.
- Fixed an issue where the filter became inactive if the list was filtered empty.
- Other minor refactoring.

## 0.3.0

- Added the ability to show current tenant information.
- Fixed a bug where the application icon wasn't reset after opening the manifest.
- Added better error handling to detect user not being logged in.

## 0.2.0

- Included functionality to delete and upload certificate credentials.

## 0.1.0

- Initial release.
