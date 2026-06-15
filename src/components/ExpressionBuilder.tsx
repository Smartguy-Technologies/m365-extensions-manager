import { useMemo, useState } from "react";
import type { PublicClientApplication } from "@azure/msal-browser";
import type { AppSettings } from "../config";
import { queryUsersByFilter, type GraphUser } from "../graph";
import { ATTRIBUTE_NAMES, parseItems } from "../attributes";
import UserEditor from "./UserEditor";

interface Props {
  msal: PublicClientApplication;
  settings: AppSettings;
}

type Operator = "eq" | "ne" | "startsWith" | "hasValue" | "isEmpty";

interface Condition {
  id: number;
  attribute: string;
  operator: Operator;
  value: string;
}

const OPERATORS: { value: Operator; label: string; needsValue: boolean }[] = [
  { value: "eq", label: "equals", needsValue: true },
  { value: "ne", label: "not equals", needsValue: true },
  { value: "startsWith", label: "starts with", needsValue: true },
  { value: "hasValue", label: "has any value", needsValue: false },
  { value: "isEmpty", label: "is empty", needsValue: false },
];

function operatorNeedsValue(op: Operator): boolean {
  return OPERATORS.find((o) => o.value === op)?.needsValue ?? true;
}

/** Escape a string literal for an OData filter (single quotes are doubled). */
function escapeOData(value: string): string {
  return value.replace(/'/g, "''");
}

/** Build the OData clause for one condition, or null if it's still incomplete. */
function buildClause(c: Condition): string | null {
  const field = `onPremisesExtensionAttributes/${c.attribute}`;
  const v = c.value.trim();
  switch (c.operator) {
    case "hasValue":
      return `${field} ne null`;
    case "isEmpty":
      return `${field} eq null`;
    case "eq":
      return v ? `${field} eq '${escapeOData(v)}'` : null;
    case "ne":
      return v ? `${field} ne '${escapeOData(v)}'` : null;
    case "startsWith":
      return v ? `startswith(${field}, '${escapeOData(v)}')` : null;
  }
}

let conditionId = 0;
function blankCondition(): Condition {
  return { id: ++conditionId, attribute: ATTRIBUTE_NAMES[0], operator: "eq", value: "" };
}

export default function ExpressionBuilder({ msal, settings }: Props) {
  const [conditions, setConditions] = useState<Condition[]>(() => [blankCondition()]);
  const [combinator, setCombinator] = useState<"and" | "or">("and");
  const [results, setResults] = useState<GraphUser[]>([]);
  const [selected, setSelected] = useState<GraphUser | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [searched, setSearched] = useState(false);
  const [copied, setCopied] = useState(false);

  // Assemble the $filter from every condition that is complete enough to emit.
  const { filter, completeCount } = useMemo(() => {
    const clauses = conditions
      .map(buildClause)
      .filter((c): c is string => c !== null);
    const joiner = combinator === "and" ? " and " : " or ";
    return { filter: clauses.join(joiner), completeCount: clauses.length };
  }, [conditions, combinator]);

  // Attributes referenced by the conditions, in canonical order, for the table.
  const referencedAttributes = useMemo(() => {
    const used = new Set(conditions.map((c) => c.attribute));
    return ATTRIBUTE_NAMES.filter((n) => used.has(n));
  }, [conditions]);

  function updateCondition(id: number, patch: Partial<Condition>) {
    setConditions((prev) => prev.map((c) => (c.id === id ? { ...c, ...patch } : c)));
  }

  function addCondition() {
    setConditions((prev) => [...prev, blankCondition()]);
  }

  function removeCondition(id: number) {
    setConditions((prev) => (prev.length > 1 ? prev.filter((c) => c.id !== id) : prev));
  }

  async function copyFilter() {
    try {
      await navigator.clipboard.writeText(filter);
      setCopied(true);
      setTimeout(() => setCopied(false), 1500);
    } catch {
      /* clipboard unavailable (e.g. insecure context) */
    }
  }

  async function runQuery() {
    if (!filter) return;
    setLoading(true);
    setError(null);
    setSelected(null);
    try {
      const users = await queryUsersByFilter(msal, filter, 999);
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
        <h2>Expression builder</h2>
        <p className="hint">
          Compose a filter across multiple extension attributes, preview the generated
          OData <code>$filter</code>, and run it. Combine the conditions to match all
          (AND) or any (OR).
        </p>

        {conditions.length > 1 && (
          <label className="inline-check">
            Match
            <select
              className="mini"
              value={combinator}
              onChange={(e) => setCombinator(e.target.value as "and" | "or")}
            >
              <option value="and">all</option>
              <option value="or">any</option>
            </select>
            of these conditions:
          </label>
        )}

        <div className="attr-list">
          {conditions.map((c, i) => (
            <div key={c.id}>
              {i > 0 && <div className="cond-joiner hint">{combinator}</div>}
              <div className="row">
                <select
                  value={c.attribute}
                  onChange={(e) => updateCondition(c.id, { attribute: e.target.value })}
                >
                  {ATTRIBUTE_NAMES.map((n) => (
                    <option key={n} value={n}>
                      {n}
                    </option>
                  ))}
                </select>
                <select
                  value={c.operator}
                  onChange={(e) =>
                    updateCondition(c.id, { operator: e.target.value as Operator })
                  }
                >
                  {OPERATORS.map((o) => (
                    <option key={o.value} value={o.value}>
                      {o.label}
                    </option>
                  ))}
                </select>
                {operatorNeedsValue(c.operator) && (
                  <input
                    value={c.value}
                    onChange={(e) => updateCondition(c.id, { value: e.target.value })}
                    placeholder="value"
                    spellCheck={false}
                  />
                )}
                <button
                  type="button"
                  className="mini danger"
                  onClick={() => removeCondition(c.id)}
                  disabled={conditions.length === 1}
                  title="Remove condition"
                >
                  Remove
                </button>
              </div>
            </div>
          ))}
        </div>

        <div className="add-item">
          <button type="button" className="mini" onClick={addCondition}>
            + Add condition
          </button>
        </div>

        <div className="filter-preview">
          <div className="filter-bar">
            <span className="hint">Generated $filter</span>
            <div className="row">
              <button
                type="button"
                className="mini"
                onClick={() => void copyFilter()}
                disabled={!filter}
              >
                {copied ? "Copied!" : "Copy"}
              </button>
              <button
                type="button"
                className="primary mini"
                onClick={() => void runQuery()}
                disabled={loading || completeCount === 0}
              >
                {loading ? "Querying…" : "Run query"}
              </button>
            </div>
          </div>
          <pre className={filter ? "filter-code" : "filter-code empty"}>
            {filter || "Add a condition to build an expression."}
          </pre>
        </div>

        {error && <div className="banner error">{error}</div>}
        {searched && !error && (
          <div className="hint">{results.length} user(s) matched.</div>
        )}

        {results.length > 0 && (
          <div className="table-scroll">
            <table className="user-table">
              <thead>
                <tr>
                  <th>Name</th>
                  <th>UPN</th>
                  {referencedAttributes.map((n) => (
                    <th key={n}>{n}</th>
                  ))}
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
                    {referencedAttributes.map((n) => (
                      <td key={n}>
                        <span className="chips compact">
                          {parseItems(
                            u.onPremisesExtensionAttributes[n],
                            settings.delimiter,
                          ).map((item, i) => (
                            <span key={i} className="chip">
                              {item}
                            </span>
                          ))}
                        </span>
                      </td>
                    ))}
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
          <div className="empty-state">Select a result to edit that user's attributes.</div>
        )}
      </div>
    </div>
  );
}
