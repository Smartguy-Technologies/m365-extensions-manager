import { useState } from "react";
import type { PublicClientApplication } from "@azure/msal-browser";
import type { AppSettings } from "../config";
import { queryByAttribute, listUsersWithAttribute, type GraphUser } from "../graph";
import { ATTRIBUTE_NAMES, parseItems } from "../attributes";
import UserEditor from "./UserEditor";

interface Props {
  msal: PublicClientApplication;
  settings: AppSettings;
}

type Mode = "equals" | "containsItem" | "hasAnyValue";

export default function AttributeQuery({ msal, settings }: Props) {
  const [attribute, setAttribute] = useState(ATTRIBUTE_NAMES[0]);
  const [mode, setMode] = useState<Mode>("containsItem");
  const [value, setValue] = useState("");
  const [results, setResults] = useState<GraphUser[]>([]);
  const [selected, setSelected] = useState<GraphUser | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [searched, setSearched] = useState(false);

  async function runQuery() {
    setLoading(true);
    setError(null);
    setSelected(null);
    try {
      let users: GraphUser[];
      if (mode === "equals") {
        users = await queryByAttribute(msal, attribute, value);
      } else {
        // Graph can only filter exact matches on this property, so for item
        // membership we fetch everyone with a value and filter client-side.
        users = await listUsersWithAttribute(msal, attribute, 999);
        if (mode === "containsItem") {
          const needle = value.trim().toLowerCase();
          users = users.filter((u) =>
            parseItems(u.onPremisesExtensionAttributes[attribute], settings.delimiter).some(
              (item) => item.toLowerCase() === needle,
            ),
          );
        }
      }
      setResults(users);
      setSearched(true);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setLoading(false);
    }
  }

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
        {error && <div className="banner error">{error}</div>}
        {searched && !error && (
          <div className="hint">
            {results.length} user(s) matched on <code>{attribute}</code>.
          </div>
        )}
        {results.length > 0 && (
          <table className="user-table">
            <thead>
              <tr>
                <th>Name</th>
                <th>UPN</th>
                <th>{attribute}</th>
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
                    <div className="chips compact">
                      {parseItems(
                        u.onPremisesExtensionAttributes[attribute],
                        settings.delimiter,
                      ).map((item, i) => (
                        <span key={i} className="chip">
                          {item}
                        </span>
                      ))}
                    </div>
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
