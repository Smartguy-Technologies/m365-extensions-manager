import { useMemo, useState } from "react";
import type { PublicClientApplication } from "@azure/msal-browser";
import type { AppSettings } from "../config";
import { listAllUsers, type GraphUser } from "../graph";
import { ATTRIBUTE_NAMES, parseItems } from "../attributes";
import UserEditor from "./UserEditor";

interface Props {
  msal: PublicClientApplication;
  settings: AppSettings;
}

function setAttributeNames(u: GraphUser): string[] {
  return ATTRIBUTE_NAMES.filter((n) => u.onPremisesExtensionAttributes[n]);
}

function csvEscape(value: string): string {
  return /[",\n\r]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
}

function buildCsv(users: GraphUser[]): string {
  const header = [
    "displayName",
    "userPrincipalName",
    "mail",
    "source",
    ...ATTRIBUTE_NAMES,
  ];
  const rows = users.map((u) =>
    [
      u.displayName,
      u.userPrincipalName,
      u.mail ?? "",
      u.onPremisesSyncEnabled ? "AD synced" : "Cloud",
      ...ATTRIBUTE_NAMES.map((n) => u.onPremisesExtensionAttributes[n] ?? ""),
    ]
      .map(csvEscape)
      .join(","),
  );
  return [header.join(","), ...rows].join("\r\n");
}

function downloadCsv(users: GraphUser[]) {
  const blob = new Blob([buildCsv(users)], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `extension-attributes-${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

export default function AllUsersTab({ msal, settings }: Props) {
  const [users, setUsers] = useState<GraphUser[] | null>(null);
  const [selected, setSelected] = useState<GraphUser | null>(null);
  const [loading, setLoading] = useState(false);
  const [progress, setProgress] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [onlyWithValues, setOnlyWithValues] = useState(true);
  const [textFilter, setTextFilter] = useState("");

  async function load() {
    setLoading(true);
    setError(null);
    setSelected(null);
    try {
      setUsers(await listAllUsers(msal, (n) => setProgress(`Fetched ${n} users…`)));
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setProgress(null);
      setLoading(false);
    }
  }

  const filtered = useMemo(() => {
    if (!users) return [];
    const needle = textFilter.trim().toLowerCase();
    return users.filter((u) => {
      if (onlyWithValues && setAttributeNames(u).length === 0) return false;
      if (!needle) return true;
      if (u.displayName.toLowerCase().includes(needle)) return true;
      if (u.userPrincipalName.toLowerCase().includes(needle)) return true;
      return ATTRIBUTE_NAMES.some((n) =>
        (u.onPremisesExtensionAttributes[n] ?? "").toLowerCase().includes(needle),
      );
    });
  }, [users, onlyWithValues, textFilter]);

  return (
    <div className="two-col">
      <div className="panel">
        <h2>All users</h2>
        <div className="row">
          <button className="primary" onClick={load} disabled={loading}>
            {loading ? "Loading…" : users ? "Reload all users" : "Load all users"}
          </button>
          {users && (
            <button onClick={() => downloadCsv(filtered)} disabled={filtered.length === 0}>
              ⬇ Export CSV ({filtered.length})
            </button>
          )}
          {progress && <span className="hint">{progress}</span>}
        </div>
        {error && <div className="banner error">{error}</div>}

        {users && (
          <>
            <div className="row" style={{ margin: "0.7rem 0" }}>
              <input
                value={textFilter}
                onChange={(e) => setTextFilter(e.target.value)}
                placeholder="Filter by name, UPN or attribute value…"
                spellCheck={false}
              />
              <label className="inline-check">
                <input
                  type="checkbox"
                  checked={onlyWithValues}
                  onChange={(e) => setOnlyWithValues(e.target.checked)}
                />
                Only users with attributes set
              </label>
            </div>
            <div className="hint">
              Showing {filtered.length} of {users.length} loaded user(s).
            </div>
            <table className="user-table">
              <thead>
                <tr>
                  <th>Name</th>
                  <th>UPN</th>
                  <th>Source</th>
                  <th>Extension attributes</th>
                </tr>
              </thead>
              <tbody>
                {filtered.map((u) => (
                  <tr
                    key={u.id}
                    className={selected?.id === u.id ? "selected" : ""}
                    onClick={() => setSelected(u)}
                  >
                    <td>{u.displayName}</td>
                    <td className="mono">{u.userPrincipalName}</td>
                    <td>
                      {u.onPremisesSyncEnabled ? (
                        <span className="tag warn">AD synced</span>
                      ) : (
                        <span className="tag ok">Cloud</span>
                      )}
                    </td>
                    <td>
                      {setAttributeNames(u).length === 0 && (
                        <span className="hint">none</span>
                      )}
                      {setAttributeNames(u).map((name) => (
                        <div key={name} className="attr-value-line">
                          <span className="attr-label">{name}</span>
                          <span className="chips compact">
                            {parseItems(
                              u.onPremisesExtensionAttributes[name],
                              settings.delimiter,
                            ).map((item, i) => (
                              <span key={i} className="chip">
                                {item}
                              </span>
                            ))}
                          </span>
                        </div>
                      ))}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </>
        )}
        {!users && !loading && (
          <div className="empty-state">
            Load every user in the tenant (paged through Graph) to review and export their
            extension attributes.
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
              setUsers((prev) => prev?.map((r) => (r.id === u.id ? u : r)) ?? prev);
            }}
          />
        ) : (
          <div className="empty-state">Select a user to edit their attributes.</div>
        )}
      </div>
    </div>
  );
}
