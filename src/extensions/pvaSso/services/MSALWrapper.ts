import {
  PublicClientApplication,
  AuthenticationResult,
  BrowserAuthOptions,
  Configuration,
  InteractionRequiredAuthError,
  BrowserSystemOptions,
  AccountInfo,
  SilentRequest,
  PopupRequest
} from "@azure/msal-browser";

export class MSALWrapper {
  private msalInstance: PublicClientApplication;
  private initializePromise: Promise<void> | null = null;
  private static readonly TIMEOUT = 10000;

  constructor(clientId: string, authority: string) {
    const authConfig: BrowserAuthOptions = {
      clientId,
      authority,
      redirectUri: window.location.origin,
    };

    const systemConfig: BrowserSystemOptions = {
      loggerOptions: {
        logLevel: 3,
        loggerCallback: (level: number, message: string) => {
          if (level > 2) {
            if (process.env.NODE_ENV !== 'production') {
              console.warn(`MSAL: ${message}`);
            }
          }
        }
      }
    };

    const msalConfig: Configuration = {
      auth: authConfig,
      system: systemConfig,
      cache: {
        cacheLocation: 'sessionStorage',
        storeAuthStateInCookie: false,
      }
    };

    this.msalInstance = new PublicClientApplication(msalConfig);
    this.initializePromise = this.initializeMsal();
  }

  private async initializeMsal(): Promise<void> {
    try {
      await Promise.race([
        this.msalInstance.initialize(),
        new Promise((_, reject) => 
          setTimeout(() => reject(new Error('MSAL initialization timeout')), 
          MSALWrapper.TIMEOUT)
        )
      ]);
      
      this.msalInstance.addEventCallback((event) => {
        if (event.error) {
          console.error('MSAL Event Error:', event.error);
        }
      });
      
    } catch (error) {
      console.error("MSAL Initialization Error:", error);
      throw error;
    }
  }

  private async ensureInitialized(): Promise<void> {
    if (this.initializePromise) {
      await this.initializePromise;
    }
  }

  public async handleLoggedInUser(
    scopes: string[], 
    userEmail: string
  ): Promise<AuthenticationResult | null> {
    try {
      await this.ensureInitialized();
      
      const accounts = this.msalInstance.getAllAccounts();
      if (!accounts || accounts.length === 0) {
        console.log("No accounts found");
        return null;
      }

      let targetAccount: AccountInfo | null = null;
      if (accounts.length > 1) {
        targetAccount = this.msalInstance.getAccountByUsername(userEmail);
      } else {
        targetAccount = accounts[0];
      }

      if (!targetAccount) {
        console.log("No matching account found");
        return null;
      }

      const silentRequest: SilentRequest = {
        scopes,
        account: targetAccount,
        forceRefresh: false
      };

      try {
        return await Promise.race([
          this.msalInstance.acquireTokenSilent(silentRequest),
          new Promise<never>((_, reject) => 
            setTimeout(() => reject(new Error('Token acquisition timeout')), 
            MSALWrapper.TIMEOUT)
          )
        ]) as AuthenticationResult;
      } catch (error) {
        console.warn("Silent token acquisition failed:", error);
        return null;
      }
    } catch (error) {
      console.error("Error in handleLoggedInUser:", error);
      return null;
    }
  }

  public async acquireAccessToken(
    scopes: string[], 
    userEmail: string
  ): Promise<AuthenticationResult | null> {
    try {
      await this.ensureInitialized();

      const silentRequest: SilentRequest = {
        scopes,
        forceRefresh: false
      };

      try {
        return await this.msalInstance.ssoSilent(silentRequest);
      } catch (error) {
        if (error instanceof InteractionRequiredAuthError) {
          const popupRequest: PopupRequest = {
            scopes,
            loginHint: userEmail
          };

          try {
            return await Promise.race([
              this.msalInstance.loginPopup(popupRequest),
              new Promise<never>((_, reject) => 
                setTimeout(() => reject(new Error('Login popup timeout')), 
                MSALWrapper.TIMEOUT)
              )
            ]) as AuthenticationResult;
          } catch (popupError) {
            console.error("Popup login failed:", popupError);
            return null;
          }
        }
        console.error("SSO Silent failed:", error);
        return null;
      }
    } catch (error) {
      console.error("Error in acquireAccessToken:", error);
      return null;
    }
  }

  public async logout(): Promise<void> {
    try {
      await this.ensureInitialized();
      const currentAccount = this.msalInstance.getAllAccounts()[0];
      if (currentAccount) {
        await this.msalInstance.logoutPopup({
          account: currentAccount
        });
      }
    } catch (error) {
      console.error("Logout failed:", error);
      throw error;
    }
  }

  public getCurrentAccount(): AccountInfo | null {
    const accounts = this.msalInstance.getAllAccounts();
    return accounts?.[0] || null;
  }
}

export default MSALWrapper;