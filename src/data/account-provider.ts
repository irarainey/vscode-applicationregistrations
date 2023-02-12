import { AccountInformation } from "../types/account-information";

// Provides functionality related to the currently signed-in account
export interface AccountProvider{
    /**
     * Log in a user to the indicated tenant (default tenant if blank)
     */
    loginUser(tenant: string): Promise<boolean>;

    /**
     * Log out the current user
     */
    logoutUser() : Promise<void>

    /**
     * Retrieve account information for the currently logged in user
     */
    getAccountInformation(): Promise<AccountInformation>;
}