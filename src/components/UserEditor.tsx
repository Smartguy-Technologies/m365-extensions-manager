import { useMemo, useState } from "react";
import type { PublicClientApplication } from "@azure/msal-browser";
import type { AppSettings } from "../config";
import { getUser, updateExtensionAttributes, type GraphUser } from "../graph";
import {
  ATTRIBUTE_NAMES,
  addItems,
  dedupeItems,
  moveItem,
  parseItems,
  redistributeItems,
  serializeItems,
  shuffleItems,
  sortItemsGlobally,
  type ExtensionAttributes,
} from "../attributes";

interface Props {
  msal: PublicClientApplication;
  settings: AppSettings;
  user: GraphUser;
  onUserUpdated: (u: GraphUser) => void;
}

export default function UserEditor({ msal, settings, user, onUserUpdated }: Props) {
  const original = user.onPremisesExtensionAttributes;
  const [attrs, setAttrs] = useState<ExtensionAttributes>(() => ({ ...original }));
  const [newItem, setNewItem] = useState<Record<string, string>>({});
  const [quickAddValue, setQuickAddValue] = useState("");
  const [restockChecked, setRestockChecked] = useState<Set<string>>(new Set());
  const [notice, setNotice] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [saving, setSaving] = useState(false);

  const { delimiter } = settings;
  const synced = user.onPremisesSyncEnabled === true;

  const changedNames = useMemo(
    () => ATTRIBUTE_NAMES.filter((n) => (attrs[n] ?? null) !== (original[n] ?? null)),
    [attrs, original],
  );
  const dirty = changedNames.length > 0;

  function setItems(name: string, items: string[]) {
    setAttrs((prev) => ({ ...prev, [name]: serializeItems(items, delimiter) }));
  }

  function itemsOf(name: string): string[] {
    return parseItems(attrs[name], delimiter);
  }

  function handleAdd(name: string) {
    const raw = (newItem[name] ?? "").trim();
    if (!raw) return;
    const { items, skipped } = addItems(itemsOf(name), [raw]);
    if (skipped.length > 0) {
      setNotice(`"${skipped[0]}" is already in ${name} — duplicate skipped.`);
    } else {
      setItems(name, items);
      setNotice(null);
    }
    setNewItem((prev) => ({ ...prev, [name]: "" }));
  }

  function handleMoveAll(from: string, to: string) {
    if (from === to) return;
    const moving = itemsOf(from);
    if (moving.length === 0) return;
    const { items, skipped } = addItems(itemsOf(to), moving);
    setAttrs((prev) => ({
      ...prev,
      [from]: null,
      [to]: serializeItems(items, delimiter),
    }));
    setNotice(
      skipped.length > 0
        ? `Moved ${moving.length - skipped.length} item(s) to ${to}; ${skipped.length} duplicate(s) dropped.`
        : `Moved ${moving.length} item(s) from ${from} to ${to}.`,
    );
  }

  /** Assign a value to the user's next empty extensionAttribute (1→15). */
  function handleQuickAdd() {
    const value = quickAddValue.trim();
    if (!value) return;
    const target = ATTRIBUTE_NAMES.find((n) => itemsOf(n).length === 0);
    if (!target) {
      setError("All 15 extension attributes already have a value — none available to add to.");
      return;
    }
    setItems(target, [value]);
    setNotice(`Added "${value}" to ${target}.`);
    setError(null);
    setQuickAddValue("");
  }

  function handleRestock() {
    const names = ATTRIBUTE_NAMES.filter((n) => restockChecked.has(n));
    if (names.length < 2) {
      setNotice("Check at least two attributes to restock across.");
      return;
    }
    setAttrs((prev) => redistributeItems(prev, names, delimiter));
    setNotice(`Redistributed items evenly across ${names.length} attributes.`);
  }

  function handleSortAll() {
    setAttrs((prev) => sortItemsGlobally(prev, ATTRIBUTE_NAMES, delimiter));
    setNotice("Sorted all items A–Z across extensionAttribute1–15.");
  }

  function handleDedupeAll() {
    setAttrs((prev) => redistributeItems(prev, ATTRIBUTE_NAMES, delimiter));
    setNotice("Removed duplicate items across extensionAttribute1–15.");
  }

  async function handleSave() {
    setSaving(true);
    setError(null);
    try {
      const patch: ExtensionAttributes = {};
      for (const n of changedNames) patch[n] = attrs[n] ?? null;
      await updateExtensionAttributes(msal, user.id, patch);
      const fresh = await getUser(msal, user.id);
      onUserUpdated(fresh);
      setAttrs({ ...fresh.onPremisesExtensionAttributes });
      setNotice("Saved.");
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setSaving(false);
    }
  }

  return (
    <div>
      <div className="editor-head">
        <div>
          <h2>{user.displayName}</h2>
          <div className="mono hint">{user.userPrincipalName}</div>
        </div>
        <div className="row">
          <button
            onClick={() => {
              setAttrs({ ...original });
              setNotice(null);
              setError(null);
            }}
            disabled={!dirty || saving}
          >
            Discard changes
          </button>
          <button className="primary" onClick={handleSave} disabled={!dirty || saving || synced}>
            {saving ? "Saving…" : `Save${dirty ? ` (${changedNames.length})` : ""}`}
          </button>
        </div>
      </div>

      {synced && (
        <div className="banner warn">
          This user is synced from on-premises Active Directory. extensionAttribute1–15 are
          read-only in Entra for synced users — edit them in local AD.
        </div>
      )}
      {error && <div className="banner error">{error}</div>}
      {notice && !error && <div className="banner info">{notice}</div>}

      <div className="row restock-bar">
        <button onClick={handleSortAll} disabled={synced}>
          A–Z (all)
        </button>
        <button onClick={handleDedupeAll} disabled={synced}>
          De-dupe (all)
        </button>
        <span className="hint">
          Gathers items from all 15 attributes, sorts or de-duplicates them, and redistributes
          them evenly.
        </span>
      </div>

      <form
        className="row restock-bar"
        onSubmit={(e) => {
          e.preventDefault();
          handleQuickAdd();
        }}
      >
        <input
          value={quickAddValue}
          onChange={(e) => setQuickAddValue(e.target.value)}
          placeholder="New value…"
          spellCheck={false}
          disabled={synced}
        />
        <button className="primary" type="submit" disabled={!quickAddValue.trim() || synced}>
          + Add attribute
        </button>
        <span className="hint">Assigns the value to the next empty extensionAttribute1–15.</span>
      </form>

      <div className="row restock-bar">
        <button onClick={handleRestock} disabled={synced}>
          ♻ Restock checked attributes
        </button>
        <span className="hint">
          Gathers all items from the checked attributes, removes duplicates, and spreads them
          evenly across them.
        </span>
      </div>

      <div className="attr-list">
        {ATTRIBUTE_NAMES.map((name, i) => {
          const items = itemsOf(name);
          const changed = changedNames.includes(name);
          const used = items.length > 0;
          return (
            <div
              key={name}
              className={`attr-row ${used ? "used" : "available"} ${changed ? "changed" : ""}`}
            >
              <div className="attr-row-head">
                <label className="attr-name">
                  <input
                    type="checkbox"
                    checked={restockChecked.has(name)}
                    onChange={(e) => {
                      setRestockChecked((prev) => {
                        const next = new Set(prev);
                        if (e.target.checked) next.add(name);
                        else next.delete(name);
                        return next;
                      });
                    }}
                    title="Include in restock"
                  />
                  <strong>{i + 1}</strong> {name}
                  {changed && <span className="tag info">modified</span>}
                  <span className={`tag ${used ? "muted" : "ok"}`}>
                    {used ? "in use" : "available"}
                  </span>
                </label>
                <div className="attr-actions">
                  <button
                    className="mini"
                    title="Shuffle item order"
                    disabled={items.length < 2}
                    onClick={() => setItems(name, shuffleItems(items))}
                  >
                    🔀
                  </button>
                  <button
                    className="mini"
                    title="Sort A–Z"
                    disabled={items.length < 2}
                    onClick={() =>
                      setItems(name, [...items].sort((a, b) => a.localeCompare(b)))
                    }
                  >
                    A–Z
                  </button>
                  <button
                    className="mini"
                    title="Remove duplicate items"
                    disabled={items.length < 2}
                    onClick={() => setItems(name, dedupeItems(items))}
                  >
                    De-dupe
                  </button>
                  <select
                    className="mini"
                    title="Move all items to another attribute"
                    value=""
                    disabled={items.length === 0}
                    onChange={(e) => {
                      if (e.target.value) handleMoveAll(name, e.target.value);
                      e.target.value = "";
                    }}
                  >
                    <option value="">Move to…</option>
                    {ATTRIBUTE_NAMES.filter((n) => n !== name).map((n) => (
                      <option key={n} value={n}>
                        {n}
                      </option>
                    ))}
                  </select>
                  <button
                    className="mini danger"
                    title="Clear this attribute"
                    disabled={items.length === 0}
                    onClick={() => setItems(name, [])}
                  >
                    Clear
                  </button>
                </div>
              </div>
              <div className="chips">
                {items.length === 0 && <span className="hint">empty</span>}
                {items.map((item, idx) => (
                  <span key={`${item}-${idx}`} className="chip">
                    <button
                      className="chip-btn"
                      title="Move left"
                      disabled={idx === 0}
                      onClick={() => setItems(name, moveItem(items, idx, idx - 1))}
                    >
                      ‹
                    </button>
                    {item}
                    <button
                      className="chip-btn"
                      title="Move right"
                      disabled={idx === items.length - 1}
                      onClick={() => setItems(name, moveItem(items, idx, idx + 1))}
                    >
                      ›
                    </button>
                    <button
                      className="chip-btn remove"
                      title="Remove item"
                      onClick={() =>
                        setItems(
                          name,
                          items.filter((_, j) => j !== idx),
                        )
                      }
                    >
                      ✕
                    </button>
                  </span>
                ))}
              </div>
              {settings.allowMultiValueAttributes || items.length === 0 ? (
                <form
                  className="row add-item"
                  onSubmit={(e) => {
                    e.preventDefault();
                    handleAdd(name);
                  }}
                >
                  <input
                    value={newItem[name] ?? ""}
                    onChange={(e) => setNewItem((prev) => ({ ...prev, [name]: e.target.value }))}
                    placeholder="Add item…"
                    spellCheck={false}
                  />
                  <button className="mini" type="submit" disabled={!(newItem[name] ?? "").trim()}>
                    Add
                  </button>
                </form>
              ) : (
                <div className="hint add-item">
                  Multiple values per attribute are disabled — clear this attribute first.
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}
