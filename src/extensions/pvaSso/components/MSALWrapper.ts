import {
  PublicClientApplication,
  AuthenticationResult,
  Configuration,
  InteractionRequiredAuthError,
} from "@azure/msal-browser";

export class MSALWrapper {
  private msalConfig: Configuration;
  private msalInstance: PublicClientApplication;

  constructor(clientId: string, authority: string) {
    this.msalConfig = {
      auth: {
        clientId: clientId,
        authority: authority,
      },
      cache: {
        cacheLocation: "localStorage", // Persistent cache for tokens
      },
    };

    this.msalInstance = new PublicClientApplication(this.msalConfig);
  }

  /**
   * Handle a logged-in user by attempting to retrieve an access token silently.
   * @param scopes - The scopes for which access is being requested.
   * @param userEmail - The email address of the user.
   * @returns An AuthenticationResult or null if no token is available.
   */
  public async handleLoggedInUser(
    scopes: string[],
    userEmail: string
  ): Promise<AuthenticationResult | null> {
    try {
      const accounts = this.msalInstance.getAllAccounts();

      if (!accounts || accounts.length === 0) {
        console.log("No users are signed in.");
        return null;
      }

      const userAccount =
        accounts.length === 1
          ? accounts[0]
          : this.msalInstance.getAccountByUsername(userEmail);

      if (!userAccount) {
        console.log(`No account found for user: ${userEmail}`);
        return null;
      }

      const accessTokenRequest = {
        scopes: scopes,
        account: userAccount,
      };

      return await this.msalInstance.acquireTokenSilent(accessTokenRequest);
    } catch (error) {
      console.error("Error in handleLoggedInUser:", error);
      return null;
    }
  }

  /**
   * Acquire an access token interactively if needed.
   * @param scopes - The scopes for which access is being requested.
   * @param userEmail - The email address of the user.
   * @returns An AuthenticationResult or null if token acquisition fails.
   */
  public async acquireAccessToken(
    scopes: string[],
    userEmail: string
  ): Promise<AuthenticationResult | null> {
    const accessTokenRequest = {
      scopes: scopes,
      loginHint: userEmail,
    };

    try {
      // Try to acquire the token silently
      return await this.msalInstance.ssoSilent(accessTokenRequest);
    } catch (silentError) {
      console.warn("Silent token acquisition failed:", silentError);

      // If interaction is required, fallback to a popup
      if (silentError instanceof InteractionRequiredAuthError) {
        try {
          return await this.msalInstance.loginPopup(accessTokenRequest);
        } catch (popupError) {
          console.error("Popup login failed:", popupError);
        }
      }

      return null;
    }
  }
}

export default MSALWrapper;
