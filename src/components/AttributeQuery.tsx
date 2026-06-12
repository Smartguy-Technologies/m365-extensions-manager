import { useState } from "react";
import type { PublicClientApplication } from "@azure/msal-browser";
import type { AppSettings } from "../config";
import {
  queryByAttribute,
  listUsersWithAttribute,
  listAllUsers,
  type GraphUser,
} from "../graph";
import { ATTRIBUTE_NAMES, parseItems } from "../attributes";
import UserEditor from "./UserEditor";

interface Props {
  msal: PublicClientApplication;
  settings: AppSettings;
}

type Mode = "equals" | "containsItem" | "hasAnyValue";

const ALL = "all";

export default function AttributeQuery({ msal, settings }: Props) {
  const [attribute, setAttribute] = useState<string>(ATTRIBUTE_NAMES[0]);
  const [mode, setMode] = useState<Mode>("containsItem");
  const [value, setValue] = useState("");
  const [results, setResults] = useState<GraphUser[]>([]);
  const [selected, setSelected] = useState<GraphUser | null>(null);
  const [loading, setLoading] = useState(false);
  const [progress, setProgress] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [searched, setSearched] = useState(false);

  /** Attribute names a user matches the current query on. */
  function matchedAttributes(u: GraphUser, names: string[]): string[] {
    const needle = value.trim().toLowerCase();
    return names.filter((name) => {
      const raw = u.onPremisesExtensionAttributes[name];
      if (!raw) return false;
      switch (mode) {
        case "hasAnyValue":
          return true;
        case "equals":
          return raw.trim().toLowerCase() === needle;
        case "containsItem":
          return parseItems(raw, settings.delimiter).some(
            (item) => item.toLowerCase() === needle,
          );
      }
    });
  }

  async function runQuery() {
    setLoading(true);
    setError(null);
    setSelected(null);
    try {
      let users: GraphUser[];
      if (attribute === ALL) {
        // No server-side filter spans all 15 attributes — fetch everyone
        // (paged) and match locally.
        users = (
          await listAllUsers(msal, (n) => setProgress(`Fetched ${n} users…`))
        ).filter((u) => matchedAttributes(u, ATTRIBUTE_NAMES).length > 0);
      } else if (mode === "equals") {
        users = await queryByAttribute(msal, attribute, value);
      } else {
        // Graph can only filter exact matches on this property, so for item
        // membership we fetch everyone with a value and filter client-side.
        users = await listUsersWithAttribute(msal, attribute, 999);
        if (mode === "containsItem") {
          users = users.filter((u) => matchedAttributes(u, [attribute]).length > 0);
        }
      }
      setResults(users);
      setSearched(true);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setProgress(null);
      setLoading(false);
    }
  }

  const queriedNames = attribute === ALL ? ATTRIBUTE_NAMES : [attribute];

  return (
    <div className="two-col">
      <div className="panel">
        <h2>Query by attribute</h2>
        <form
          className="query-form"
          onSubmit={(e) => {
            e.preventDefault();
            void runQuery();
          }}
        >
          <select value={attribute} onChange={(e) => setAttribute(e.target.value)}>
            <option value={ALL}>All attributes (1–15)</option>
            {ATTRIBUTE_NAMES.map((n) => (
              <option key={n} value={n}>
                {n}
              </option>
            ))}
          </select>
          <select value={mode} onChange={(e) => setMode(e.target.value as Mode)}>
            <option value="containsItem">contains item</option>
            <option value="equals">equals exact value</option>
            <option value="hasAnyValue">has any value</option>
          </select>
          {mode !== "hasAnyValue" && (
            <input
              value={value}
              onChange={(e) => setValue(e.target.value)}
              placeholder={mode === "equals" ? "Full stored value" : "Item to look for"}
              spellCheck={false}
            />
          )}
          <button
            className="primary"
            type="submit"
            disabled={loading || (mode !== "hasAnyValue" && !value.trim())}
          >
            {loading ? "Querying…" : "Run query"}
          </button>
        </form>
        {progress && <div className="hint">{progress}</div>}
        {error && <div className="banner error">{error}</div>}
        {searched && !error && (
          <div className="hint">
            {results.length} user(s) matched on{" "}
            <code>{attribute === ALL ? "any attribute" : attribute}</code>.
          </div>
        )}
        {results.length > 0 && (
          <table className="user-table">
            <thead>
              <tr>
                <th>Name</th>
                <th>UPN</th>
                <th>{attribute === ALL ? "Matched attributes" : attribute}</th>
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
                    {matchedAttributes(u, queriedNames).map((name) => (
                      <div key={name} className="attr-value-line">
                        {attribute === ALL && <span className="attr-label">{name}</span>}
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
          <div className="empty-state">Select a result to edit that user's attributes.</div>
        )}
      </div>
    </div>
  );
}
