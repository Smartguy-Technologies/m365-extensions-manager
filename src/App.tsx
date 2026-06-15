import { useCallback, useEffect, useState } from "react";
import type { AccountInfo, PublicClientApplication } from "@azure/msal-browser";
import {
  loadSettings,
  saveSettings,
  settingsAreComplete,
  type AppSettings,
} from "./config";
import { getMsal, getActiveAccount, resetMsal, signIn, signOut } from "./auth";
import SettingsPanel from "./components/SettingsPanel";
import UsersTab from "./components/UsersTab";
import AttributeQuery from "./components/AttributeQuery";
import AllUsersTab from "./components/AllUsersTab";
import BulkOps from "./components/BulkOps";
import ExpressionBuilder from "./components/ExpressionBuilder";

type Tab = "users" | "query" | "all" | "bulk" | "expression" | "settings";

export default function App() {
  const [settings, setSettings] = useState<AppSettings>(() => loadSettings());
  const [msal, setMsal] = useState<PublicClientApplication | null>(null);
  const [account, setAccount] = useState<AccountInfo | null>(null);
  const [tab, setTab] = useState<Tab>(settingsAreComplete(loadSettings()) ? "users" : "settings");
  const [authError, setAuthError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    if (!settingsAreComplete(settings)) return;
    let cancelled = false;
    getMsal(settings)
      .then((app) => {
        if (cancelled) return;
        setMsal(app);
        setAccount(getActiveAccount(app));
      })
      .catch((e) => setAuthError(e instanceof Error ? e.message : String(e)));
    return () => {
      cancelled = true;
    };
  }, [settings]);

  const handleSaveSettings = useCallback((s: AppSettings) => {
    saveSettings(s);
    resetMsal();
    setMsal(null);
    setAccount(null);
    setSettings(s);
    if (settingsAreComplete(s)) setTab("users");
  }, []);

  const handleSignIn = useCallback(async () => {
    if (!msal) return;
    setBusy(true);
    setAuthError(null);
    try {
      setAccount(await signIn(msal));
    } catch (e) {
      setAuthError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  }, [msal]);

  const handleSignOut = useCallback(async () => {
    if (!msal) return;
    setBusy(true);
    try {
      await signOut(msal);
    } catch {
      /* user closed the popup */
    } finally {
      setAccount(getActiveAccount(msal));
      setBusy(false);
    }
  }, [msal]);

  const signedIn = msal !== null && account !== null;

  return (
    <div className="app">
      <header className="topbar">
        <h1>Entra Extension Attribute Manager</h1>
        <div className="topbar-right">
          {signedIn ? (
            <>
              <span className="account-chip" title={account.username}>
                {account.name ?? account.username}
              </span>
              <button onClick={handleSignOut} disabled={busy}>
                Sign out
              </button>
            </>
          ) : (
            <button
              className="primary"
              onClick={handleSignIn}
              disabled={busy || !msal}
              title={!msal ? "Configure tenant and client ID in Settings first" : ""}
            >
              {busy ? "Signing in…" : "Sign in to Entra"}
            </button>
          )}
        </div>
      </header>

      {authError && <div className="banner error">{authError}</div>}

      <nav className="tabs">
        <button className={tab === "users" ? "active" : ""} onClick={() => setTab("users")}>
          Users
        </button>
        <button className={tab === "query" ? "active" : ""} onClick={() => setTab("query")}>
          Query by attribute
        </button>
        <button className={tab === "all" ? "active" : ""} onClick={() => setTab("all")}>
          All users
        </button>
        <button className={tab === "bulk" ? "active" : ""} onClick={() => setTab("bulk")}>
          Bulk operations
        </button>
        <button
          className={tab === "expression" ? "active" : ""}
          onClick={() => setTab("expression")}
        >
          Expression builder
        </button>
        <button
          className={tab === "settings" ? "active" : ""}
          onClick={() => setTab("settings")}
        >
          Settings
        </button>
      </nav>

      <main className="content">
        {tab === "settings" && (
          <SettingsPanel settings={settings} onSave={handleSaveSettings} />
        )}
        {tab !== "settings" && !settingsAreComplete(settings) && (
          <div className="empty-state">
            Configure your Entra tenant ID and app registration client ID in{" "}
            <button className="link" onClick={() => setTab("settings")}>
              Settings
            </button>{" "}
            to get started.
          </div>
        )}
        {tab !== "settings" && settingsAreComplete(settings) && !signedIn && (
          <div className="empty-state">
            Sign in with an account that has <code>User.ReadWrite.All</code> permission to
            manage extension attributes.
          </div>
        )}
        {signedIn && tab === "users" && <UsersTab msal={msal} settings={settings} />}
        {signedIn && tab === "query" && <AttributeQuery msal={msal} settings={settings} />}
        {signedIn && tab === "all" && <AllUsersTab msal={msal} settings={settings} />}
        {signedIn && tab === "bulk" && <BulkOps msal={msal} settings={settings} />}
        {signedIn && tab === "expression" && (
          <ExpressionBuilder msal={msal} settings={settings} />
        )}
      </main>
    </div>
  );
}
