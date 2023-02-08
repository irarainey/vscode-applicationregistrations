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

- ## 0.3.0

- Added the ability to show current tenant information.
- Fixed a bug where the application icon wasn't reset after opening the manifest.
- Added better error handling to detect user not being logged in.

## 0.2.0

- Included functionality to delete and upload certificate credentials.

## 0.1.0

- Initial release.
