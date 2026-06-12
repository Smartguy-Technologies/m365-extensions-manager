/**
 * Helpers for treating a single extension attribute string as a list of items
 * separated by a configurable delimiter, with duplicate prevention.
 */

export const ATTRIBUTE_NAMES = Array.from(
  { length: 15 },
  (_, i) => `extensionAttribute${i + 1}`,
);

export type ExtensionAttributes = Record<string, string | null>;

export function parseItems(value: string | null | undefined, delimiter: string): string[] {
  if (!value) return [];
  return value
    .split(delimiter)
    .map((s) => s.trim())
    .filter((s) => s.length > 0);
}

export function serializeItems(items: string[], delimiter: string): string | null {
  const cleaned = items.map((s) => s.trim()).filter((s) => s.length > 0);
  if (cleaned.length === 0) return null; // null clears the attribute in Graph
  return cleaned.join(delimiter);
}

/** Case-insensitive de-duplication preserving first occurrence. */
export function dedupeItems(items: string[]): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const item of items) {
    const key = item.toLowerCase();
    if (!seen.has(key)) {
      seen.add(key);
      out.push(item);
    }
  }
  return out;
}

/** Add items, skipping any that already exist (case-insensitive). Returns the new list and the items that were skipped as duplicates. */
export function addItems(
  existing: string[],
  toAdd: string[],
): { items: string[]; skipped: string[] } {
  const seen = new Set(existing.map((s) => s.toLowerCase()));
  const items = [...existing];
  const skipped: string[] = [];
  for (const raw of toAdd) {
    const item = raw.trim();
    if (!item) continue;
    const key = item.toLowerCase();
    if (seen.has(key)) {
      skipped.push(item);
    } else {
      seen.add(key);
      items.push(item);
    }
  }
  return { items, skipped };
}

/** Remove items (case-insensitive). Returns the new list and which requested items were actually found. */
export function removeItems(
  existing: string[],
  toRemove: string[],
): { items: string[]; removed: string[] } {
  const removeSet = new Set(toRemove.map((s) => s.trim().toLowerCase()).filter(Boolean));
  const items: string[] = [];
  const removed: string[] = [];
  for (const item of existing) {
    if (removeSet.has(item.toLowerCase())) {
      removed.push(item);
    } else {
      items.push(item);
    }
  }
  return { items, removed };
}

/** Fisher-Yates shuffle (returns a new array). */
export function shuffleItems<T>(items: T[]): T[] {
  const out = [...items];
  for (let i = out.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [out[i], out[j]] = [out[j], out[i]];
  }
  return out;
}

export function moveItem<T>(items: T[], from: number, to: number): T[] {
  if (to < 0 || to >= items.length) return items;
  const out = [...items];
  const [moved] = out.splice(from, 1);
  out.splice(to, 0, moved);
  return out;
}

/**
 * "Restock": gather every item from the given attributes and redistribute them
 * evenly (round-robin) across those same attributes, de-duplicated.
 */
export function redistributeItems(
  attrs: ExtensionAttributes,
  attributeNames: string[],
  delimiter: string,
): ExtensionAttributes {
  const all = dedupeItems(
    attributeNames.flatMap((name) => parseItems(attrs[name], delimiter)),
  );
  const buckets: string[][] = attributeNames.map(() => []);
  all.forEach((item, i) => buckets[i % attributeNames.length].push(item));
  const result: ExtensionAttributes = { ...attrs };
  attributeNames.forEach((name, i) => {
    result[name] = serializeItems(buckets[i], delimiter);
  });
  return result;
}

/** Parse a textarea of values: one per line, also splitting on the delimiter and commas. */
export function parseInputValues(raw: string, delimiter: string): string[] {
  return dedupeItems(
    raw
      .split(/\r?\n/)
      .flatMap((line) => line.split(delimiter))
      .flatMap((part) => (delimiter === "," ? [part] : part.split(",")))
      .map((s) => s.trim())
      .filter((s) => s.length > 0),
  );
}
