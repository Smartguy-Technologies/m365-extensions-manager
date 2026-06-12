import {
  PublicClientApplication,
  InteractionRequiredAuthError,
  type AccountInfo,
} from "@azure/msal-browser";
import type { AppSettings } from "./config";

export const GRAPH_SCOPES = ["User.ReadWrite.All"];

let pca: PublicClientApplication | null = null;

export async function getMsal(settings: AppSettings): Promise<PublicClientApplication> {
  if (pca) return pca;
  pca = new PublicClientApplication({
    auth: {
      clientId: settings.clientId.trim(),
      authority: `https://login.microsoftonline.com/${settings.tenantId.trim()}`,
      redirectUri: window.location.origin,
    },
    cache: {
      cacheLocation: "sessionStorage",
    },
  });
  await pca.initialize();
  return pca;
}

/** Drop the cached instance so new settings take effect. */
export function resetMsal(): void {
  pca = null;
}

export function getActiveAccount(app: PublicClientApplication): AccountInfo | null {
  return app.getActiveAccount() ?? app.getAllAccounts()[0] ?? null;
}

export async function signIn(app: PublicClientApplication): Promise<AccountInfo> {
  const result = await app.loginPopup({ scopes: GRAPH_SCOPES, prompt: "select_account" });
  app.setActiveAccount(result.account);
  return result.account;
}

export async function signOut(app: PublicClientApplication): Promise<void> {
  const account = getActiveAccount(app);
  await app.logoutPopup(account ? { account } : undefined);
}

export async function getToken(app: PublicClientApplication): Promise<string> {
  const account = getActiveAccount(app);
  if (!account) throw new Error("Not signed in");
  try {
    const result = await app.acquireTokenSilent({ scopes: GRAPH_SCOPES, account });
    return result.accessToken;
  } catch (e) {
    if (e instanceof InteractionRequiredAuthError) {
      const result = await app.acquireTokenPopup({ scopes: GRAPH_SCOPES, account });
      return result.accessToken;
    }
    throw e;
  }
}
