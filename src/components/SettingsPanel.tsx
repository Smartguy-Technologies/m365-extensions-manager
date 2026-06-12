import { useState } from "react";
import { ENV_DEFAULTS, type AppSettings } from "../config";

interface Props {
  settings: AppSettings;
  onSave: (s: AppSettings) => void;
}

export default function SettingsPanel({ settings, onSave }: Props) {
  const [tenantId, setTenantId] = useState(settings.tenantId);
  const [clientId, setClientId] = useState(settings.clientId);
  const [delimiter, setDelimiter] = useState(settings.delimiter);
  const [saved, setSaved] = useState(false);

  return (
    <div className="panel narrow">
      <h2>Connection settings</h2>
      <p className="hint">
        Create an app registration in Entra (single-page application, redirect URI{" "}
        <code>{window.location.origin}</code>) with delegated permission{" "}
        <code>User.ReadWrite.All</code> (admin consent required), then paste its IDs here.
        Settings are stored only in this browser. Defaults can also be supplied via a{" "}
        <code>.env.local</code> file (<code>M365_TENANT_ID</code>, <code>EAM_APP_CLIENT_ID</code>,{" "}
        <code>VITE_DELIMITER</code>) — values saved here override them.
      </p>
      <label>
        Tenant ID
        <input
          value={tenantId}
          onChange={(e) => setTenantId(e.target.value)}
          placeholder="00000000-0000-0000-0000-000000000000"
          spellCheck={false}
        />
        {ENV_DEFAULTS.tenantId && tenantId === ENV_DEFAULTS.tenantId && (
          <span className="hint">Using default from M365_TENANT_ID</span>
        )}
      </label>
      <label>
        Application (client) ID
        <input
          value={clientId}
          onChange={(e) => setClientId(e.target.value)}
          placeholder="00000000-0000-0000-0000-000000000000"
          spellCheck={false}
        />
        {ENV_DEFAULTS.clientId && clientId === ENV_DEFAULTS.clientId && (
          <span className="hint">Using default from EAM_APP_CLIENT_ID</span>
        )}
      </label>
      <label>
        Item delimiter
        <input
          value={delimiter}
          onChange={(e) => setDelimiter(e.target.value)}
          maxLength={3}
          style={{ width: "4rem" }}
        />
        <span className="hint">
          Used to store multiple items inside one attribute value (default <code>;</code>).
        </span>
      </label>
      <div className="row">
        <button
          className="primary"
          disabled={!tenantId.trim() || !clientId.trim() || !delimiter}
          onClick={() => {
            onSave({ tenantId: tenantId.trim(), clientId: clientId.trim(), delimiter });
            setSaved(true);
            setTimeout(() => setSaved(false), 2000);
          }}
        >
          Save settings
        </button>
        {saved && <span className="ok">Saved ✓</span>}
      </div>

      <h3>Notes</h3>
      <ul className="hint">
        <li>
          extensionAttribute1–15 live on <code>onPremisesExtensionAttributes</code>. They are
          writable through this app only for <strong>cloud-only</strong> users; for users
          synced from on-premises AD they must be edited in local Active Directory. Synced
          users are flagged in the UI.
        </li>
        <li>Each attribute holds one string (max 1024 characters); this app can pack multiple delimited items into it.</li>
      </ul>
    </div>
  );
}
