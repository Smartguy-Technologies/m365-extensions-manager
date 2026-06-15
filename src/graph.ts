import type { PublicClientApplication } from "@azure/msal-browser";
import { getToken } from "./auth";
import type { ExtensionAttributes } from "./attributes";

const GRAPH = "https://graph.microsoft.com/v1.0";

export interface GraphUser {
  id: string;
  displayName: string;
  userPrincipalName: string;
  mail: string | null;
  accountEnabled?: boolean;
  onPremisesSyncEnabled: boolean | null;
  onPremisesExtensionAttributes: ExtensionAttributes;
}

export const USER_SELECT =
  "id,displayName,userPrincipalName,mail,accountEnabled,onPremisesSyncEnabled,onPremisesExtensionAttributes";

export class GraphError extends Error {
  status: number;
  constructor(status: number, message: string) {
    super(message);
    this.status = status;
  }
}

async function graphFetch(
  app: PublicClientApplication,
  path: string,
  init: RequestInit = {},
  extraHeaders: Record<string, string> = {},
): Promise<Response> {
  const token = await getToken(app);
  const url = path.startsWith("http") ? path : `${GRAPH}${path}`;
  const res = await fetch(url, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...extraHeaders,
      ...(init.headers as Record<string, string> | undefined),
    },
  });
  if (!res.ok) {
    let message = `${res.status} ${res.statusText}`;
    try {
      const body = await res.json();
      if (body?.error?.message) message = body.error.message;
    } catch {
      /* keep default message */
    }
    throw new GraphError(res.status, message);
  }
  return res;
}

function normalizeUser(u: GraphUser): GraphUser {
  return { ...u, onPremisesExtensionAttributes: u.onPremisesExtensionAttributes ?? {} };
}

/** Search users by first name, last name, display name, nickname, UPN or mail prefix. */
export async function searchUsers(
  app: PublicClientApplication,
  query: string,
  top = 25,
): Promise<GraphUser[]> {
  const q = query.trim().replace(/'/g, "''");
  let path: string;
  if (q.length === 0) {
    path = `/users?$select=${USER_SELECT}&$top=${top}&$orderby=displayName&$count=true`;
  } else {
    const filter = encodeURIComponent(
      [
        `startswith(displayName,'${q}')`,
        `startswith(givenName,'${q}')`,
        `startswith(surname,'${q}')`,
        `startswith(mailNickname,'${q}')`,
        `startswith(userPrincipalName,'${q}')`,
        `startswith(mail,'${q}')`,
      ].join(" or "),
    );
    path = `/users?$select=${USER_SELECT}&$filter=${filter}&$top=${top}&$count=true`;
  }
  const res = await graphFetch(app, path, {}, { ConsistencyLevel: "eventual" });
  const body = await res.json();
  return (body.value as GraphUser[]).map(normalizeUser);
}

/** Look up a single user by id or userPrincipalName. */
export async function getUser(
  app: PublicClientApplication,
  idOrUpn: string,
): Promise<GraphUser> {
  const res = await graphFetch(
    app,
    `/users/${encodeURIComponent(idOrUpn.trim())}?$select=${USER_SELECT}`,
  );
  return normalizeUser(await res.json());
}

/** Find users whose given extensionAttribute exactly equals a value. */
export async function queryByAttribute(
  app: PublicClientApplication,
  attributeName: string,
  value: string,
  top = 100,
): Promise<GraphUser[]> {
  const v = value.trim().replace(/'/g, "''");
  const filter = encodeURIComponent(
    `onPremisesExtensionAttributes/${attributeName} eq '${v}'`,
  );
  const path = `/users?$select=${USER_SELECT}&$filter=${filter}&$top=${top}&$count=true`;
  const res = await graphFetch(app, path, {}, { ConsistencyLevel: "eventual" });
  const body = await res.json();
  return (body.value as GraphUser[]).map(normalizeUser);
}

/**
 * List all users that have ANY value set for the given attribute.
 * Graph cannot filter "ne null" on this property without $count + eventual
 * consistency; we use `not(... eq null)` which requires the same headers.
 */
export async function listUsersWithAttribute(
  app: PublicClientApplication,
  attributeName: string,
  top = 100,
): Promise<GraphUser[]> {
  const filter = encodeURIComponent(
    `onPremisesExtensionAttributes/${attributeName} ne null`,
  );
  const path = `/users?$select=${USER_SELECT}&$filter=${filter}&$top=${top}&$count=true`;
  const res = await graphFetch(app, path, {}, { ConsistencyLevel: "eventual" });
  const body = await res.json();
  return (body.value as GraphUser[]).map(normalizeUser);
}

/**
 * Run an arbitrary OData $filter against /users and return the matches.
 * Used by the Expression Builder to evaluate compound filters. Filtering on
 * onPremisesExtensionAttributes requires advanced query support
 * (ConsistencyLevel: eventual + $count=true).
 */
export async function queryUsersByFilter(
  app: PublicClientApplication,
  filter: string,
  top = 100,
): Promise<GraphUser[]> {
  const path = `/users?$select=${USER_SELECT}&$filter=${encodeURIComponent(
    filter,
  )}&$top=${top}&$count=true`;
  const res = await graphFetch(app, path, {}, { ConsistencyLevel: "eventual" });
  const body = await res.json();
  return (body.value as GraphUser[]).map(normalizeUser);
}

/**
 * Page through every user in the tenant, following @odata.nextLink.
 * Reports progress as pages arrive so the UI can show a running count.
 */
export async function listAllUsers(
  app: PublicClientApplication,
  onProgress?: (count: number) => void,
): Promise<GraphUser[]> {
  let url: string | null = `/users?$select=${USER_SELECT}&$top=999&$count=true`;
  const all: GraphUser[] = [];
  while (url) {
    const res = await graphFetch(app, url, {}, { ConsistencyLevel: "eventual" });
    const body: { value: GraphUser[]; "@odata.nextLink"?: string } = await res.json();
    all.push(...body.value.map(normalizeUser));
    onProgress?.(all.length);
    url = body["@odata.nextLink"] ?? null;
  }
  return all;
}

/**
 * Patch one or more extension attributes on a user.
 * Pass null as a value to clear that attribute.
 * NOTE: for users synced from on-premises AD this property is read-only in
 * Graph and the call will fail — manage those in local AD instead.
 */
export async function updateExtensionAttributes(
  app: PublicClientApplication,
  userId: string,
  attrs: ExtensionAttributes,
): Promise<void> {
  await graphFetch(app, `/users/${encodeURIComponent(userId)}`, {
    method: "PATCH",
    body: JSON.stringify({ onPremisesExtensionAttributes: attrs }),
  });
}

export interface BatchResult {
  userId: string;
  ok: boolean;
  status: number;
  error?: string;
}

/**
 * Update extension attributes for many users via the $batch endpoint
 * (20 requests per batch). Returns one result per user.
 */
export async function batchUpdateExtensionAttributes(
  app: PublicClientApplication,
  updates: { userId: string; attrs: ExtensionAttributes }[],
  onProgress?: (done: number, total: number) => void,
): Promise<BatchResult[]> {
  const results: BatchResult[] = [];
  const CHUNK = 20;
  for (let i = 0; i < updates.length; i += CHUNK) {
    const chunk = updates.slice(i, i + CHUNK);
    const res = await graphFetch(app, "/$batch", {
      method: "POST",
      body: JSON.stringify({
        requests: chunk.map((u, idx) => ({
          id: String(idx),
          method: "PATCH",
          url: `/users/${encodeURIComponent(u.userId)}`,
          headers: { "Content-Type": "application/json" },
          body: { onPremisesExtensionAttributes: u.attrs },
        })),
      }),
    });
    const body = await res.json();
    for (const r of body.responses as { id: string; status: number; body?: any }[]) {
      const u = chunk[Number(r.id)];
      results.push({
        userId: u.userId,
        ok: r.status >= 200 && r.status < 300,
        status: r.status,
        error:
          r.status >= 300 ? r.body?.error?.message ?? `HTTP ${r.status}` : undefined,
      });
    }
    onProgress?.(Math.min(i + CHUNK, updates.length), updates.length);
  }
  return results;
}
