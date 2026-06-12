# Entra Extension Attribute Manager

A local React app for managing **extensionAttribute1–15** (`onPremisesExtensionAttributes`)
on Microsoft Entra ID users via Microsoft Graph. Everything runs in your browser —
there is no backend and nothing leaves your machine except calls to Microsoft Graph.

## Features

- **Sign in to Entra** with MSAL (popup flow, delegated permissions)
- **Per-user editor** for all 15 extension attributes
  - Treats each attribute value as a delimited list of items (delimiter configurable, default `;`)
  - Add items with **duplicate prevention** (case-insensitive)
  - Remove individual items, clear an attribute
  - **Reshuffle**: reorder items (move left/right), random shuffle, sort A–Z
  - **Restock**: gather items from selected attributes, de-dupe, and redistribute evenly across them
  - Move all items from one attribute to another
  - De-dupe button for existing values
- **Query by user**: search by name, UPN, or email
- **Query by attribute**: partial text match, exact item membership, exact value match,
  or "has any value" — against a single attribute or **all 15 attributes at once**
- **All users report**: page through every user in the tenant, filter the results,
  review all extension attributes, and **export to CSV**
- **Bulk operations** with preview-before-apply:
  - Bulk add items (duplicates skipped per user)
  - Bulk remove items
  - Overwrite or clear an attribute
  - Paste a list of UPNs/object IDs (e.g. a CSV column); updates are sent via Graph `$batch`
- Flags **AD-synced users**, whose extension attributes are read-only in Entra and must be
  managed in on-premises Active Directory

## Prerequisites

1. **Node.js 18+**
2. **An Entra app registration** in your tenant:
   - Microsoft Entra admin center → Identity → Applications → App registrations → **New registration**
   - Supported account types: *Accounts in this organizational directory only*
   - Platform: **Single-page application (SPA)** with redirect URI `http://localhost:5173`
   - API permissions: Microsoft Graph → **Delegated** → `User.ReadWrite.All`, then **Grant admin consent**
   - Note the **Application (client) ID** and **Directory (tenant) ID**

## Run locally

```bash
npm install
npm run dev
```

Open <http://localhost:5173>, go to **Settings**, paste your tenant ID and client ID,
save, then click **Sign in to Entra**.

### Pre-configuring via environment variables

Instead of entering IDs in the Settings tab, you can provide defaults through Vite
environment variables — copy `.env.example` to `.env.local` and fill in:

```bash
VITE_TENANT_ID=<your tenant id>
VITE_CLIENT_ID=<your app registration client id>
# optional, defaults to ;
VITE_DELIMITER=;
```

These are read at dev-server start (or build time). Values saved in the Settings tab
take precedence over the environment defaults; empty saved fields fall back to them.

## Notes & limitations

- Each extension attribute stores a single string (max 1024 characters). This app packs
  multiple "items" into one value using a delimiter — keep total length under the limit.
- Graph supports server-side filtering on these attributes only as exact `eq` matches
  (with `ConsistencyLevel: eventual`). The "contains item" and "has any value" query modes
  fetch users with a value set and filter client-side.
- Writes fail for users synced from on-premises AD (`onPremisesSyncEnabled = true`);
  the app detects and skips them.
- The signed-in account needs a role allowed to update users (e.g. User Administrator),
  in addition to the app's `User.ReadWrite.All` consent.
