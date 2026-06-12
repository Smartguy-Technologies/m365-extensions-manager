import { useState } from "react";
import type { PublicClientApplication } from "@azure/msal-browser";
import type { AppSettings } from "../config";
import {
  getUser,
  batchUpdateExtensionAttributes,
  type BatchResult,
  type GraphUser,
} from "../graph";
import {
  ATTRIBUTE_NAMES,
  addItems,
  parseInputValues,
  parseItems,
  removeItems,
  serializeItems,
} from "../attributes";

interface Props {
  msal: PublicClientApplication;
  settings: AppSettings;
}

type Operation = "add" | "remove" | "overwrite" | "clear";

/** Sentinel for "write to the user's next empty attribute" (Add only). */
const AUTO = "auto";

interface PlannedChange {
  user: GraphUser;
  /** Attribute this change targets (resolved per user in Auto mode). */
  attribute: string;
  before: string | null;
  after: string | null;
  note: string;
  skip: boolean;
}

export default function BulkOps({ msal, settings }: Props) {
  const [upnText, setUpnText] = useState("");
  const [attribute, setAttribute] = useState<string>(AUTO);
  const [operation, setOperation] = useState<Operation>("add");
  const [valuesText, setValuesText] = useState("");
  const [plan, setPlan] = useState<PlannedChange[] | null>(null);
  const [unresolved, setUnresolved] = useState<{ upn: string; error: string }[]>([]);
  const [results, setResults] = useState<BatchResult[] | null>(null);
  const [progress, setProgress] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  const { delimiter } = settings;

  /** Add into the user's first empty attribute, deduping against ALL 15 attributes. */
  function planAutoAdd(user: GraphUser, values: string[]): PlannedChange {
    const everywhere = ATTRIBUTE_NAMES.flatMap((n) =>
      parseItems(user.onPremisesExtensionAttributes[n], delimiter),
    );
    const { items: merged, skipped } = addItems(everywhere, values);
    const newOnes = merged.slice(everywhere.length);
    if (newOnes.length === 0) {
      return {
        user,
        attribute: "—",
        before: null,
        after: null,
        note: "nothing to add (already present in some attribute)",
        skip: true,
      };
    }
    const target = ATTRIBUTE_NAMES.find((n) => !user.onPremisesExtensionAttributes[n]);
    if (!target) {
      return {
        user,
        attribute: "—",
        before: null,
        after: null,
        note: "no empty attribute available (all 15 in use)",
        skip: true,
      };
    }
    return {
      user,
      attribute: target,
      before: null,
      after: serializeItems(newOnes, delimiter),
      note: `+${newOnes.length} item(s)${
        skipped.length ? `, ${skipped.length} duplicate(s) skipped` : ""
      }`,
      skip: false,
    };
  }

  function planFor(user: GraphUser): PlannedChange {
    const values = parseInputValues(valuesText, delimiter);

    if (user.onPremisesSyncEnabled) {
      return {
        user,
        attribute: attribute === AUTO ? "—" : attribute,
        before: null,
        after: null,
        note: "AD-synced user — read-only, skipped",
        skip: true,
      };
    }

    if (operation === "add" && attribute === AUTO) {
      return planAutoAdd(user, values);
    }

    const before = user.onPremisesExtensionAttributes[attribute] ?? null;
    const existing = parseItems(before, delimiter);

    switch (operation) {
      case "add": {
        const { items, skipped } = addItems(existing, values);
        const added = values.length - skipped.length;
        return {
          user,
          attribute,
          before,
          after: serializeItems(items, delimiter),
          note:
            added === 0
              ? "nothing to add (all duplicates)"
              : `+${added} item(s)${skipped.length ? `, ${skipped.length} duplicate(s) skipped` : ""}`,
          skip: added === 0,
        };
      }
      case "remove": {
        const { items, removed } = removeItems(existing, values);
        return {
          user,
          attribute,
          before,
          after: serializeItems(items, delimiter),
          note: removed.length === 0 ? "no matching items" : `-${removed.length} item(s)`,
          skip: removed.length === 0,
        };
      }
      case "overwrite": {
        const after = serializeItems(values, delimiter);
        return { user, attribute, before, after, note: "value replaced", skip: after === before };
      }
      case "clear":
        return { user, attribute, before, after: null, note: "cleared", skip: before === null };
    }
  }

  async function buildPlan() {
    setBusy(true);
    setError(null);
    setResults(null);
    setPlan(null);
    try {
      const upns = Array.from(
        new Set(
          upnText
            .split(/\r?\n|,/)
            .map((s) => s.trim())
            .filter(Boolean),
        ),
      );
      if (upns.length === 0) throw new Error("Enter at least one user (UPN or object ID).");
      if ((operation === "add" || operation === "remove" || operation === "overwrite") &&
          parseInputValues(valuesText, delimiter).length === 0) {
        throw new Error("Enter at least one value.");
      }
      const resolved: GraphUser[] = [];
      const failed: { upn: string; error: string }[] = [];
      // Resolve sequentially-ish in small parallel chunks to stay under throttling limits.
      const CHUNK = 10;
      for (let i = 0; i < upns.length; i += CHUNK) {
        const chunk = upns.slice(i, i + CHUNK);
        const settled = await Promise.allSettled(chunk.map((upn) => getUser(msal, upn)));
        settled.forEach((r, j) => {
          if (r.status === "fulfilled") resolved.push(r.value);
          else
            failed.push({
              upn: chunk[j],
              error: r.reason instanceof Error ? r.reason.message : String(r.reason),
            });
        });
        setProgress(`Resolving users… ${Math.min(i + CHUNK, upns.length)}/${upns.length}`);
      }
      setUnresolved(failed);
      setPlan(resolved.map(planFor));
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setProgress(null);
      setBusy(false);
    }
  }

  async function applyPlan() {
    if (!plan) return;
    const changes = plan.filter((p) => !p.skip);
    if (changes.length === 0) return;
    setBusy(true);
    setError(null);
    try {
      const res = await batchUpdateExtensionAttributes(
        msal,
        changes.map((c) => ({ userId: c.user.id, attrs: { [c.attribute]: c.after } })),
        (done, total) => setProgress(`Applying… ${done}/${total}`),
      );
      setResults(res);
      setPlan(null);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setProgress(null);
      setBusy(false);
    }
  }

  const applicable = plan?.filter((p) => !p.skip).length ?? 0;

  return (
    <div className="panel">
      <h2>Bulk operations</h2>
      <div className="bulk-grid">
        <label>
          Users (UPN or object ID, one per line — paste a CSV column)
          <textarea
            rows={8}
            value={upnText}
            onChange={(e) => setUpnText(e.target.value)}
            placeholder={"alice@contoso.com\nbob@contoso.com"}
            spellCheck={false}
          />
        </label>
        <div>
          <label>
            Operation
            <select
              value={operation}
              onChange={(e) => {
                const op = e.target.value as Operation;
                setOperation(op);
                // Auto-targeting only makes sense for Add.
                if (op !== "add" && attribute === AUTO) setAttribute(ATTRIBUTE_NAMES[0]);
              }}
            >
              <option value="add">Add items (skips duplicates)</option>
              <option value="remove">Remove items</option>
              <option value="overwrite">Overwrite value</option>
              <option value="clear">Clear attribute</option>
            </select>
          </label>
          <label>
            Attribute{operation === "add" ? " (optional)" : ""}
            <select value={attribute} onChange={(e) => setAttribute(e.target.value)}>
              {operation === "add" && (
                <option value={AUTO}>Auto — next available attribute</option>
              )}
              {ATTRIBUTE_NAMES.map((n) => (
                <option key={n} value={n}>
                  {n}
                </option>
              ))}
            </select>
            {operation === "add" && attribute === AUTO && (
              <span className="hint">
                Writes to each user's first empty attribute (1→15); duplicates are checked
                across all 15 attributes.
              </span>
            )}
          </label>
          {operation !== "clear" && (
            <label>
              Values (one per line or delimited)
              <textarea
                rows={4}
                value={valuesText}
                onChange={(e) => setValuesText(e.target.value)}
                placeholder={"ProjectX\nVIP"}
                spellCheck={false}
              />
            </label>
          )}
        </div>
      </div>

      <div className="row">
        <button className="primary" onClick={buildPlan} disabled={busy}>
          Preview changes
        </button>
        {plan && (
          <button className="danger" onClick={applyPlan} disabled={busy || applicable === 0}>
            Apply to {applicable} user(s)
          </button>
        )}
        {progress && <span className="hint">{progress}</span>}
      </div>

      {error && <div className="banner error">{error}</div>}

      {unresolved.length > 0 && (
        <div className="banner warn">
          {unresolved.length} user(s) could not be resolved:{" "}
          {unresolved.map((f) => f.upn).join(", ")}
        </div>
      )}

      {plan && (
        <>
          <h3>Preview ({applicable} change(s), {plan.length - applicable} skipped)</h3>
          <table className="user-table">
            <thead>
              <tr>
                <th>User</th>
                <th>Attribute</th>
                <th>Before</th>
                <th>After</th>
                <th>Note</th>
              </tr>
            </thead>
            <tbody>
              {plan.map((p) => (
                <tr key={p.user.id} className={p.skip ? "skipped" : ""}>
                  <td className="mono">{p.user.userPrincipalName}</td>
                  <td className="mono">{p.attribute}</td>
                  <td className="mono">{p.before ?? <em>empty</em>}</td>
                  <td className="mono">{p.skip ? "—" : p.after ?? <em>empty</em>}</td>
                  <td>{p.note}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </>
      )}

      {results && (
        <>
          <h3>
            Results: {results.filter((r) => r.ok).length} succeeded,{" "}
            {results.filter((r) => !r.ok).length} failed
          </h3>
          {results.some((r) => !r.ok) && (
            <table className="user-table">
              <thead>
                <tr>
                  <th>User ID</th>
                  <th>Status</th>
                  <th>Error</th>
                </tr>
              </thead>
              <tbody>
                {results
                  .filter((r) => !r.ok)
                  .map((r) => (
                    <tr key={r.userId}>
                      <td className="mono">{r.userId}</td>
                      <td>{r.status}</td>
                      <td>{r.error}</td>
                    </tr>
                  ))}
              </tbody>
            </table>
          )}
        </>
      )}
    </div>
  );
}
