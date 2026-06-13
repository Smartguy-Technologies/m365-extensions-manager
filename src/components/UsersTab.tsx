import { useState } from "react";
import type { PublicClientApplication } from "@azure/msal-browser";
import type { AppSettings } from "../config";
import { searchUsers, getUser, type GraphUser } from "../graph";
import { ATTRIBUTE_NAMES } from "../attributes";
import UserEditor from "./UserEditor";

interface Props {
  msal: PublicClientApplication;
  settings: AppSettings;
}

function countSetAttributes(u: GraphUser): number {
  return ATTRIBUTE_NAMES.filter((n) => u.onPremisesExtensionAttributes[n]).length;
}

export default function UsersTab({ msal, settings }: Props) {
  const [query, setQuery] = useState("");
  const [results, setResults] = useState<GraphUser[]>([]);
  const [selected, setSelected] = useState<GraphUser | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [searched, setSearched] = useState(false);

  async function runSearch() {
    setLoading(true);
    setError(null);
    try {
      let users: GraphUser[];
      if (query.includes("@")) {
        // Looks like a UPN/email — try an exact lookup first, fall back to search.
        try {
          users = [await getUser(msal, query)];
        } catch {
          users = await searchUsers(msal, query);
        }
      } else {
        users = await searchUsers(msal, query);
      }
      setResults(users);
      setSearched(true);
      if (users.length === 1) setSelected(users[0]);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setLoading(false);
    }
  }

  return (
    <div className="two-col">
      <div className="panel">
        <h2>Find users</h2>
        <form
          className="row"
          onSubmit={(e) => {
            e.preventDefault();
            void runSearch();
          }}
        >
          <input
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="Name, UPN or email — empty lists first 25"
            spellCheck={false}
          />
          <button className="primary" type="submit" disabled={loading}>
            {loading ? "Searching…" : "Search"}
          </button>
        </form>
        {error && <div className="banner error">{error}</div>}
        {searched && results.length === 0 && !error && (
          <div className="empty-state">No users found.</div>
        )}
        {results.length > 0 && (
          <div className="table-scroll">
          <table className="user-table">
            <thead>
              <tr>
                <th>Name</th>
                <th>UPN</th>
                <th>Source</th>
                <th>Attrs set</th>
              </tr>
            </thead>
            <tbody>
              {results.map((u) => (
                <tr
                  key={u.id}
                  className={selected?.id === u.id ? "selected" : ""}
                  onClick={() => setSelected(u)}
                >
                  <td>{u.displayName}</td>
                  <td className="mono">{u.userPrincipalName}</td>
                  <td>
                    {u.onPremisesSyncEnabled ? (
                      <span className="tag warn" title="Synced from on-premises AD — attributes are read-only in Entra">
                        AD synced
                      </span>
                    ) : (
                      <span className="tag ok">Cloud</span>
                    )}
                  </td>
                  <td className="center">{countSetAttributes(u)}</td>
                </tr>
              ))}
            </tbody>
          </table>
          </div>
        )}
      </div>
      <div className="panel">
        {selected ? (
          <UserEditor
            key={selected.id}
            msal={msal}
            settings={settings}
            user={selected}
            onUserUpdated={(u) => {
              setSelected(u);
              setResults((prev) => prev.map((r) => (r.id === u.id ? u : r)));
            }}
          />
        ) : (
          <div className="empty-state">Select a user to view and edit extensionAttribute1–15.</div>
        )}
      </div>
    </div>
  );
}
