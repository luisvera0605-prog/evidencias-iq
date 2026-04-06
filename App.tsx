import { useState, useEffect, useCallback, useRef } from "react";

const CONFIG = {
  clientId: "b271f29f-65f7-476e-a272-63669bdfd85e",
  tenantId: "746b050c-a1ff-45b9-9858-e142490982b7",
  siteUrl: "https://cisurft.sharepoint.com/sites/PlaneacionFinanciera",
  redirectUri: window.location.origin,
  scopes: ["Files.Read", "Files.Read.All", "offline_access", "User.Read"],
};

// ============================================================
// AUTH (PKCE)
// ============================================================
const AUTH = {
  async getToken() {
    const stored = sessionStorage.getItem("ms_token");
    if (stored) {
      const { token, expires } = JSON.parse(stored);
      if (Date.now() < expires) return token;
    }
    return null;
  },
  async login() {
    const verifier = Array.from(crypto.getRandomValues(new Uint8Array(32)))
      .map((b) => b.toString(16).padStart(2, "0")).join("");
    const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier));
    const challenge = btoa(String.fromCharCode(...new Uint8Array(digest)))
      .replace(/\+/g, "-").replace(/\//g, "_").replace(/=/g, "");
    sessionStorage.setItem("pkce_verifier", verifier);
    const state = Math.random().toString(36).slice(2);
    sessionStorage.setItem("oauth_state", state);
    const params = new URLSearchParams({
      client_id: CONFIG.clientId, response_type: "code",
      redirect_uri: CONFIG.redirectUri, scope: CONFIG.scopes.join(" "),
      state, code_challenge: challenge, code_challenge_method: "S256", response_mode: "query",
    });
    window.location.href = `https://login.microsoftonline.com/${CONFIG.tenantId}/oauth2/v2.0/authorize?${params}`;
  },
  async handleCallback(code) {
    const body = new URLSearchParams({
      client_id: CONFIG.clientId, grant_type: "authorization_code", code,
      redirect_uri: CONFIG.redirectUri, code_verifier: sessionStorage.getItem("pkce_verifier"),
      scope: CONFIG.scopes.join(" "),
    });
    const res = await fetch(`https://login.microsoftonline.com/${CONFIG.tenantId}/oauth2/v2.0/token`,
      { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body });
    const data = await res.json();
    if (data.access_token) {
      sessionStorage.setItem("ms_token", JSON.stringify({
        token: data.access_token,
        expires: Date.now() + (data.expires_in - 60) * 1000,
      }));
      return data.access_token;
    }
    throw new Error(data.error_description || "Auth failed");
  },
  logout() {
    ["ms_token","pkce_verifier","oauth_state"].forEach(k => sessionStorage.removeItem(k));
  },
};

// ============================================================
// GRAPH API
// ============================================================
async function graphGet(url, token, asBlob = false) {
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) throw new Error(`Graph ${res.status}: ${url}`);
  return asBlob ? res.blob() : res.json();
}

async function getSiteId(token) {
  const u = new URL(CONFIG.siteUrl);
  const data = await graphGet(
    `https://graph.microsoft.com/v1.0/sites/${u.hostname}:${u.pathname}`, token);
  return data.id;
}

// Get the drive ID for the "Evidencias" document library
async function getEvidenciasDriveId(token, siteId) {
  const data = await graphGet(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`, token);
  const drives = data.value || [];
  // Find the Evidencias library by name
  const ev = drives.find(d =>
    d.name === "Evidencias" ||
    d.webUrl?.toLowerCase().includes("/evidencias")
  );
  if (ev) return ev.id;
  // Fallback: return first non-default drive or default drive
  return drives[0]?.id || null;
}

// List ALL children with automatic pagination — handles 500, 1000, 5000+ items
async function listChildren(token, driveId, path) {
  const firstUrl = path
    ? `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${encodeURIComponent(path)}:/children?$top=999`
    : `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children?$top=999`;

  const allItems = [];
  let url = firstUrl;

  while (url) {
    const data = await graphGet(url, token);
    allItems.push(...(data.value || []));
    // Graph API returns @odata.nextLink when there are more pages
    url = data["@odata.nextLink"] || null;
  }

  return allItems;
}

const MEDIA_EXTS = /\.(jpg|jpeg|png|gif|webp|pdf|bmp|tiff|heic)$/i;

// Recursively find all evidence files under Ventas folder inside Evidencias library
async function scanEvidencias(token, driveId, basePath, onProgress) {
  const results = [];
  // List folio folders
  const folders = await listChildren(token, driveId, basePath);
  const folioFolders = folders.filter(f => f.folder);
  onProgress({ phase: "scanning", total: folioFolders.length, current: 0 });

  for (let i = 0; i < folioFolders.length; i++) {
    const folder = folioFolders[i];
    onProgress({ phase: "scanning", total: folioFolders.length, current: i + 1, current_folio: folder.name });
    try {
      const files = await listChildren(token, driveId, `${basePath}/${folder.name}`);
      const mediaFiles = files.filter(f => f.file && MEDIA_EXTS.test(f.name));
      for (const file of mediaFiles) {
        results.push({
          id: file.id,
          name: file.name,
          folio: folder.name,
          size: file.size,
          modified: file.lastModifiedDateTime,
          mimeType: file.file?.mimeType || "",
          folderName: folder.name,
          downloadUrl: file["@microsoft.graph.downloadUrl"] || null,
        });
      }
    } catch (e) {
      // skip inaccessible folder
    }
  }
  return results;
}

async function getFileBlob(token, driveId, fileId) {
  return graphGet(
    `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${fileId}/content`,
    token, true
  );
}

async function blobToBase64(blob) {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(reader.result.split(",")[1]);
    reader.readAsDataURL(blob);
  });
}

// ============================================================
// CLAUDE VISION
// ============================================================
async function analyzeEvidence(base64Data, mediaType, folio) {
  const isPdf = mediaType === "application/pdf" || mediaType?.includes("pdf");
  
  const prompt = `Eres auditor financiero. Analiza esta evidencia de pago/entrega del folio "${folio}".

Extrae:
1. ¿Es legible? (true/false)
2. tipo_documento: transferencia, ticket_caja, factura, remision, foto_entrega, comprobante_pago, otro
3. fecha: fecha que aparece en el documento (formato DD/MM/YYYY o null)
4. monto: importe total que aparece (número o null)
5. referencia: número de operación/referencia/folio del documento
6. cliente_documento: nombre del cliente o receptor que aparece
7. banco_emisor: banco o institución que emite (si aplica)
8. folio_presente: ¿aparece "${folio}" en el documento? (true/false/null)
9. observaciones: máximo 1 línea de observación relevante para conciliación
10. semaforo: "verde" si todo claro, "amarillo" si datos parciales o dudas, "rojo" si ilegible o inconsistencias

Responde SOLO con JSON válido, sin texto adicional, sin markdown:
{"legible":true,"tipo_documento":"...","fecha":"...","monto":null,"referencia":"...","cliente_documento":"...","banco_emisor":"...","folio_presente":true,"observaciones":"...","semaforo":"verde"}`;

  const content = isPdf
    ? [
        { type: "document", source: { type: "base64", media_type: "application/pdf", data: base64Data } },
        { type: "text", text: prompt },
      ]
    : [
        { type: "image", source: { type: "base64", media_type: mediaType, data: base64Data } },
        { type: "text", text: prompt },
      ];

  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens: 500,
      messages: [{ role: "user", content }],
    }),
  });
  const data = await res.json();
  const text = data.content?.[0]?.text || "{}";
  try {
    return JSON.parse(text.replace(/```json|```/g, "").trim());
  } catch {
    return { legible: false, tipo_documento: "error_parse", observaciones: text.slice(0, 100), semaforo: "rojo" };
  }
}

// ============================================================
// HELPERS
// ============================================================
const fmtMXN = (n) => n != null ? new Intl.NumberFormat("es-MX", { style: "currency", currency: "MXN" }).format(n) : "—";
const fmtKB = (b) => b > 1024 * 1024 ? `${(b / 1024 / 1024).toFixed(1)} MB` : `${(b / 1024).toFixed(0)} KB`;

function getMime(file) {
  if (file.mimeType) return file.mimeType;
  const ext = file.name.split(".").pop().toLowerCase();
  const map = { jpg: "image/jpeg", jpeg: "image/jpeg", png: "image/png", gif: "image/gif", webp: "image/webp", pdf: "application/pdf", bmp: "image/bmp" };
  return map[ext] || "image/jpeg";
}

const SEM_STYLE = {
  verde:    { bg: "#d1fae5", color: "#064e3b", label: "✓ OK" },
  amarillo: { bg: "#fef3c7", color: "#78350f", label: "⚠ Revisar" },
  rojo:     { bg: "#fee2e2", color: "#7f1d1d", label: "✗ Alerta" },
};

function Pill({ text, sem }) {
  const s = SEM_STYLE[sem] || { bg: "#f1f5f9", color: "#475569" };
  return <span style={{ background: s.bg, color: s.color, fontSize: 11, fontWeight: 700, padding: "2px 8px", borderRadius: 20 }}>{text}</span>;
}

function AnalysisResult({ r }) {
  if (!r) return null;
  const s = SEM_STYLE[r.semaforo] || SEM_STYLE.amarillo;
  return (
    <div style={{ marginTop: 8, padding: 10, background: "#f8faff", borderRadius: 8, border: "1px solid #e2e8f0", fontSize: 12 }}>
      <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 6 }}>
        <span style={{ background: s.bg, color: s.color, fontSize: 11, fontWeight: 700, padding: "2px 8px", borderRadius: 20 }}>
          {s.label}
        </span>
        {r.tipo_documento && <span style={{ background: "#dbeafe", color: "#1e3a8a", fontSize: 11, fontWeight: 600, padding: "2px 8px", borderRadius: 20 }}>{r.tipo_documento}</span>}
        {!r.legible && <span style={{ background: "#fee2e2", color: "#7f1d1d", fontSize: 11, fontWeight: 600, padding: "2px 8px", borderRadius: 20 }}>Ilegible</span>}
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "3px 12px" }}>
        {[["Fecha", r.fecha],["Monto", r.monto != null ? fmtMXN(r.monto) : null],
          ["Referencia", r.referencia],["Cliente", r.cliente_documento],["Banco", r.banco_emisor]
        ].filter(([,v]) => v && v !== "null" && v !== "..." && v !== "").map(([k, v]) => (
          <div key={k} style={{ color: "#475569" }}><span style={{ color: "#94a3b8" }}>{k}: </span><strong>{v}</strong></div>
        ))}
      </div>
      {r.observaciones && <p style={{ margin: "6px 0 0", color: "#64748b", fontStyle: "italic" }}>{r.observaciones}</p>}
    </div>
  );
}

// ============================================================
// EXPORT TO CSV
// ============================================================
function exportCSV(files, analyses) {
  const headers = ["Folio","Archivo","Tipo","Semaforo","Legible","Fecha","Monto","Referencia","Cliente_Documento","Banco","Folio_Presente","Observaciones","Tamaño","Modificado"];
  const rows = files.map(f => {
    const a = analyses[f.id] || {};
    return [
      f.folio, f.name, a.tipo_documento || "", a.semaforo || "sin_analizar",
      a.legible ?? "", a.fecha || "", a.monto ?? "", a.referencia || "",
      a.cliente_documento || "", a.banco_emisor || "", a.folio_presente ?? "",
      a.observaciones || "", fmtKB(f.size),
      new Date(f.modified).toLocaleDateString("es-MX"),
    ].map(v => `"${String(v).replace(/"/g, '""')}"`);
  });
  const csv = [headers, ...rows].map(r => r.join(",")).join("\n");
  const blob = new Blob(["\ufeff" + csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a"); a.href = url;
  a.download = `Evidencias_Ventas_${new Date().toISOString().slice(0,10)}.csv`;
  a.click(); URL.revokeObjectURL(url);
}

// ============================================================
// MAIN APP
// ============================================================
export default function App() {
  const [token, setToken] = useState(null);
  const [user, setUser] = useState(null);
  const [siteId, setSiteId] = useState(null);
  const [driveId, setDriveId] = useState(null);
  const [files, setFiles] = useState([]);
  const [analyses, setAnalyses] = useState({});
  const [scanning, setScanning] = useState(false);
  const [scanProgress, setScanProgress] = useState(null);
  const [batchProgress, setBatchProgress] = useState(null);
  const [analyzing, setAnalyzing] = useState({});
  const [error, setError] = useState(null);
  const [search, setSearch] = useState("");
  const [filterSem, setFilterSem] = useState("todos");
  const [preview, setPreview] = useState(null);
  const [previewUrl, setPreviewUrl] = useState(null);
  const [basePath, setBasePath] = useState("Evidencias/Ventas");
  const stopRef = useRef(false);

  // Handle OAuth callback
  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const code = params.get("code");
    if (code) {
      window.history.replaceState({}, "", window.location.pathname);
      AUTH.handleCallback(code).then(setToken).catch(e => setError(e.message));
    } else {
      AUTH.getToken().then(t => { if (t) setToken(t); });
    }
  }, []);

  useEffect(() => {
    if (!token) return;
    graphGet("https://graph.microsoft.com/v1.0/me", token).then(setUser).catch(() => {});
    getSiteId(token)
      .then(sid => {
        setSiteId(sid);
        return getEvidenciasDriveId(token, sid);
      })
      .then(did => setDriveId(did))
      .catch(e => setError("Error conectando SharePoint: " + e.message));
  }, [token]);

  const scan = useCallback(async () => {
    if (!token || !driveId) return;
    setScanning(true); setError(null); setFiles([]); setAnalyses({});
    try {
      const found = await scanEvidencias(token, driveId, basePath, setScanProgress);
      setFiles(found);
    } catch (e) {
      setError("Error escaneando: " + e.message);
    }
    setScanning(false); setScanProgress(null);
  }, [token, driveId, basePath]);

  const analyzeOne = useCallback(async (file) => {
    setAnalyzing(a => ({ ...a, [file.id]: true }));
    try {
      const blob = await getFileBlob(token, file.driveId || driveId, file.id);
      const b64 = await blobToBase64(blob);
      const mime = getMime(file);
      const result = await analyzeEvidence(b64, mime, file.folio);
      setAnalyses(a => ({ ...a, [file.id]: result }));
    } catch (e) {
      setAnalyses(a => ({ ...a, [file.id]: { legible: false, tipo_documento: "error", observaciones: e.message, semaforo: "rojo" } }));
    }
    setAnalyzing(a => ({ ...a, [file.id]: false }));
  }, [token, driveId]);

  const analyzeAll = useCallback(async () => {
    stopRef.current = false;
    const pending = files.filter(f => !analyses[f.id]);
    setBatchProgress({ current: 0, total: pending.length, errors: 0 });
    for (let i = 0; i < pending.length; i++) {
      if (stopRef.current) break;
      await analyzeOne(pending[i]);
      setBatchProgress(p => ({ ...p, current: i + 1 }));
      // Small delay to avoid rate limiting
      if (i % 10 === 9) await new Promise(r => setTimeout(r, 1000));
    }
    setBatchProgress(null);
  }, [files, analyses, analyzeOne]);

  const openPreview = useCallback(async (file) => {
    setPreview(file); setPreviewUrl(null);
    try {
      const blob = await getFileBlob(token, file.driveId || driveId, file.id);
      setPreviewUrl(URL.createObjectURL(blob));
    } catch {}
  }, [token, driveId]);

  // Stats
  const stats = {
    total: files.length,
    analizados: Object.keys(analyses).length,
    verde: Object.values(analyses).filter(a => a.semaforo === "verde").length,
    amarillo: Object.values(analyses).filter(a => a.semaforo === "amarillo").length,
    rojo: Object.values(analyses).filter(a => a.semaforo === "rojo").length,
    montoTotal: Object.values(analyses).reduce((s, a) => s + (typeof a.monto === "number" ? a.monto : 0), 0),
  };

  const filtered = files.filter(f => {
    const a = analyses[f.id];
    const matchSearch = !search || f.folio.toLowerCase().includes(search.toLowerCase()) || f.name.toLowerCase().includes(search.toLowerCase());
    if (filterSem === "verde") return matchSearch && a?.semaforo === "verde";
    if (filterSem === "amarillo") return matchSearch && a?.semaforo === "amarillo";
    if (filterSem === "rojo") return matchSearch && a?.semaforo === "rojo";
    if (filterSem === "sin_analizar") return matchSearch && !a;
    return matchSearch;
  });

  // LOGIN SCREEN
  if (!token) return (
    <div style={{ fontFamily: "'Segoe UI',sans-serif", minHeight: "100vh", background: "#f1f5f9", display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div style={{ background: "#fff", borderRadius: 16, padding: 48, textAlign: "center", maxWidth: 400, width: "90%" }}>
        <div style={{ fontSize: 52, marginBottom: 12 }}>📋</div>
        <h1 style={{ fontSize: 22, fontWeight: 800, color: "#0f172a", margin: "0 0 8px" }}>EvidenciasIQ</h1>
        <p style={{ color: "#64748b", fontSize: 13, margin: "0 0 8px" }}>Verificación de evidencias con IA</p>
        <p style={{ color: "#94a3b8", fontSize: 12, margin: "0 0 28px" }}>SharePoint · Claude Vision · Flor de Tabasco</p>
        {error && <div style={{ background: "#fee2e2", color: "#7f1d1d", borderRadius: 8, padding: 10, marginBottom: 16, fontSize: 12 }}>{error}</div>}
        <button onClick={AUTH.login} style={{ background: "#0078d4", color: "#fff", border: "none", borderRadius: 10, padding: "13px 0", fontSize: 14, fontWeight: 700, cursor: "pointer", width: "100%" }}>
          🔐 Iniciar sesión con Microsoft
        </button>
      </div>
    </div>
  );

  return (
    <div style={{ fontFamily: "'Segoe UI',sans-serif", background: "#f1f5f9", minHeight: "100vh" }}>
      {/* HEADER */}
      <div style={{ background: "linear-gradient(135deg,#0f172a,#1e3a5f)", padding: "14px 24px", display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 10 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ width: 34, height: 34, background: "#0078d4", borderRadius: 8, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16 }}>📋</div>
          <div>
            <p style={{ color: "#fff", fontSize: 15, fontWeight: 700, margin: 0 }}>EvidenciasIQ</p>
            <p style={{ color: "#94a3b8", fontSize: 11, margin: 0 }}>{user?.displayName || "..."} · PlaneaciónFinanciera</p>
          </div>
        </div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          {files.length > 0 && !batchProgress && (
            <button onClick={analyzeAll}
              style={{ background: "#7c3aed", color: "#fff", border: "none", borderRadius: 8, padding: "8px 14px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>
              🤖 Analizar todo ({files.filter(f => !analyses[f.id]).length} pendientes)
            </button>
          )}
          {batchProgress && (
            <button onClick={() => stopRef.current = true}
              style={{ background: "#ef4444", color: "#fff", border: "none", borderRadius: 8, padding: "8px 14px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>
              ⏹ Detener
            </button>
          )}
          {files.length > 0 && Object.keys(analyses).length > 0 && (
            <button onClick={() => exportCSV(files, analyses)}
              style={{ background: "#059669", color: "#fff", border: "none", borderRadius: 8, padding: "8px 14px", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>
              ⬇ Exportar CSV
            </button>
          )}
          <button onClick={() => { AUTH.logout(); setToken(null); }}
            style={{ background: "rgba(255,255,255,0.1)", color: "#fff", border: "1px solid rgba(255,255,255,0.2)", borderRadius: 8, padding: "8px 12px", fontSize: 12, cursor: "pointer" }}>
            Salir
          </button>
        </div>
      </div>

      <div style={{ padding: 20, maxWidth: 1300, margin: "0 auto" }}>
        {error && <div style={{ background: "#fee2e2", border: "1px solid #fecaca", borderRadius: 10, padding: 12, marginBottom: 14, color: "#7f1d1d", fontSize: 13 }}>⚠️ {error}</div>}

        {/* SCAN BAR */}
        <div style={{ background: "#fff", borderRadius: 12, padding: 18, border: "1px solid #e2e8f0", marginBottom: 16 }}>
          <div style={{ display: "flex", gap: 10, alignItems: "flex-end", flexWrap: "wrap" }}>
            <div style={{ flex: 1, minWidth: 220 }}>
              <label style={{ fontSize: 11, fontWeight: 700, color: "#64748b", textTransform: "uppercase", letterSpacing: "0.05em" }}>Ruta en SharePoint</label>
              <input value={basePath} onChange={e => setBasePath(e.target.value)}
                style={{ display: "block", width: "100%", marginTop: 4, border: "1px solid #e2e8f0", borderRadius: 8, padding: "8px 12px", fontSize: 13, boxSizing: "border-box" }} />
            </div>
            <button onClick={scan} disabled={scanning || !driveId}
              style={{ background: "#1e40af", color: "#fff", border: "none", borderRadius: 8, padding: "10px 20px", fontSize: 13, fontWeight: 600, cursor: scanning ? "not-allowed" : "pointer", opacity: scanning ? 0.7 : 1 }}>
              {scanning ? "Escaneando..." : "📂 Escanear carpetas"}
            </button>
          </div>
          {scanProgress && (
            <div style={{ marginTop: 12 }}>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: "#64748b", marginBottom: 4 }}>
                <span>Leyendo carpetas: {scanProgress.current_folio || "..."}</span>
                <span>{scanProgress.current}/{scanProgress.total}</span>
              </div>
              <div style={{ background: "#e2e8f0", borderRadius: 4, height: 6 }}>
                <div style={{ background: "#3b82f6", height: "100%", borderRadius: 4, width: `${scanProgress.total ? (scanProgress.current/scanProgress.total*100) : 0}%`, transition: "width 0.3s" }} />
              </div>
            </div>
          )}
          {batchProgress && (
            <div style={{ marginTop: 12 }}>
              <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12, color: "#64748b", marginBottom: 4 }}>
                <span>🤖 Analizando con IA...</span>
                <span>{batchProgress.current}/{batchProgress.total}</span>
              </div>
              <div style={{ background: "#e2e8f0", borderRadius: 4, height: 6 }}>
                <div style={{ background: "#7c3aed", height: "100%", borderRadius: 4, width: `${batchProgress.total ? (batchProgress.current/batchProgress.total*100) : 0}%`, transition: "width 0.3s" }} />
              </div>
            </div>
          )}
        </div>

        {/* KPIs */}
        {files.length > 0 && (
          <div style={{ display: "grid", gridTemplateColumns: "repeat(6,1fr)", gap: 10, marginBottom: 16 }}>
            {[
              ["Archivos", stats.total, "#1e40af", "#dbeafe"],
              ["Analizados", stats.analizados, "#0f766e", "#ccfbf1"],
              ["✓ OK", stats.verde, "#064e3b", "#d1fae5"],
              ["⚠ Revisar", stats.amarillo, "#78350f", "#fef3c7"],
              ["✗ Alertas", stats.rojo, "#7f1d1d", "#fee2e2"],
              ["Monto total", fmtMXN(stats.montoTotal), "#4c1d95", "#ede9fe"],
            ].map(([lbl, val, c, bg]) => (
              <div key={lbl} style={{ background: "#fff", borderRadius: 10, padding: 12, border: "1px solid #e2e8f0", textAlign: "center" }}>
                <p style={{ fontSize: 10, color: "#94a3b8", margin: 0, fontWeight: 700, textTransform: "uppercase" }}>{lbl}</p>
                <p style={{ fontSize: typeof val === "string" && val.length > 8 ? 13 : 20, fontWeight: 800, color: c, margin: "4px 0 0" }}>{val}</p>
              </div>
            ))}
          </div>
        )}

        {/* FILTERS */}
        {files.length > 0 && (
          <div style={{ display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap" }}>
            <input value={search} onChange={e => setSearch(e.target.value)} placeholder="🔍 Buscar folio o archivo..."
              style={{ flex: 1, minWidth: 180, border: "1px solid #e2e8f0", borderRadius: 8, padding: "7px 12px", fontSize: 13, background: "#fff" }} />
            {["todos","verde","amarillo","rojo","sin_analizar"].map(f => (
              <button key={f} onClick={() => setFilterSem(f)}
                style={{ padding: "7px 14px", borderRadius: 20, fontSize: 12, fontWeight: 600, border: "2px solid",
                  borderColor: filterSem === f ? "#3b82f6" : "#e2e8f0",
                  background: filterSem === f ? "#eff6ff" : "#fff",
                  color: filterSem === f ? "#1e40af" : "#64748b", cursor: "pointer" }}>
                {f === "todos" ? `Todos (${files.length})` : f === "verde" ? `✓ OK (${stats.verde})` : f === "amarillo" ? `⚠ Revisar (${stats.amarillo})` : f === "rojo" ? `✗ Alertas (${stats.rojo})` : `Sin analizar (${files.length - stats.analizados})`}
              </button>
            ))}
          </div>
        )}

        {/* MAIN GRID */}
        <div style={{ display: "grid", gridTemplateColumns: preview ? "1fr 400px" : "1fr", gap: 16 }}>
          {/* FILE LIST */}
          <div>
            {filtered.length === 0 && files.length === 0 && !scanning && (
              <div style={{ background: "#fff", borderRadius: 12, padding: 60, textAlign: "center", border: "1px solid #e2e8f0" }}>
                <div style={{ fontSize: 44, marginBottom: 12 }}>📁</div>
                <p style={{ color: "#0f172a", fontWeight: 700, fontSize: 15 }}>La ruta está configurada para:</p>
                <code style={{ background: "#f1f5f9", padding: "4px 12px", borderRadius: 6, fontSize: 14, color: "#1e40af" }}>Evidencias/Ventas</code>
                <p style={{ color: "#64748b", fontSize: 13, marginTop: 12 }}>Haz clic en "Escanear carpetas" para leer los {">"}300 folios de Marzo</p>
              </div>
            )}

            {filtered.length === 0 && files.length > 0 && (
              <div style={{ background: "#fff", borderRadius: 12, padding: 40, textAlign: "center", border: "1px solid #e2e8f0" }}>
                <p style={{ color: "#64748b" }}>No hay archivos que coincidan con el filtro actual</p>
              </div>
            )}

            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              {filtered.slice(0, 200).map(file => {
                const a = analyses[file.id];
                const isAnalyzing = analyzing[file.id];
                const isPdf = file.name.toLowerCase().endsWith(".pdf");
                const sem = a?.semaforo;
                const rowBg = sem === "verde" ? "#f0fdf4" : sem === "amarillo" ? "#fffbeb" : sem === "rojo" ? "#fef2f2" : "#fff";

                return (
                  <div key={file.id} onClick={() => openPreview(file)}
                    style={{ background: rowBg, borderRadius: 10, padding: 14, border: `1px solid ${preview?.id === file.id ? "#3b82f6" : "#e2e8f0"}`, cursor: "pointer", transition: "border-color 0.15s" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10 }}>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap", marginBottom: 3 }}>
                          <span style={{ fontSize: 16 }}>{isPdf ? "📄" : "🖼️"}</span>
                          <strong style={{ fontSize: 13, color: "#0f172a" }}>{file.folio}</strong>
                          <span style={{ fontSize: 11, color: "#94a3b8" }}>{file.name}</span>
                          {a && (
                            <span style={{ background: SEM_STYLE[sem]?.bg, color: SEM_STYLE[sem]?.color, fontSize: 11, fontWeight: 700, padding: "1px 7px", borderRadius: 20 }}>
                              {SEM_STYLE[sem]?.label}
                            </span>
                          )}
                        </div>
                        <p style={{ margin: 0, fontSize: 11, color: "#94a3b8" }}>{fmtKB(file.size)} · {new Date(file.modified).toLocaleDateString("es-MX")}</p>
                        {a && <AnalysisResult r={a} />}
                      </div>
                      <button onClick={e => { e.stopPropagation(); analyzeOne(file); }} disabled={isAnalyzing}
                        style={{ flexShrink: 0, background: isAnalyzing ? "#f1f5f9" : a ? "#f1f5f9" : "#7c3aed", color: isAnalyzing ? "#94a3b8" : a ? "#64748b" : "#fff", border: "none", borderRadius: 8, padding: "6px 11px", fontSize: 11, fontWeight: 600, cursor: isAnalyzing ? "not-allowed" : "pointer", whiteSpace: "nowrap" }}>
                        {isAnalyzing ? "⏳..." : a ? "↻" : "🤖"}
                      </button>
                    </div>
                  </div>
                );
              })}
              {filtered.length > 200 && (
                <p style={{ textAlign: "center", color: "#94a3b8", fontSize: 12, padding: 12 }}>
                  Mostrando 200 de {filtered.length} — usa el buscador para filtrar
                </p>
              )}
            </div>
          </div>

          {/* PREVIEW PANEL */}
          {preview && (
            <div style={{ background: "#fff", borderRadius: 12, border: "1px solid #e2e8f0", overflow: "hidden", position: "sticky", top: 16, maxHeight: "88vh", display: "flex", flexDirection: "column" }}>
              <div style={{ padding: "10px 14px", borderBottom: "1px solid #f1f5f9", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div>
                  <p style={{ margin: 0, fontSize: 13, fontWeight: 700, color: "#0f172a" }}>{preview.folio}</p>
                  <p style={{ margin: 0, fontSize: 11, color: "#94a3b8" }}>{preview.name}</p>
                </div>
                <button onClick={() => { setPreview(null); setPreviewUrl(null); }}
                  style={{ background: "none", border: "none", fontSize: 18, cursor: "pointer", color: "#64748b" }}>✕</button>
              </div>
              <div style={{ flex: 1, overflow: "auto", padding: 12, background: "#f8faff", display: "flex", alignItems: "center", justifyContent: "center" }}>
                {previewUrl
                  ? preview.name.toLowerCase().endsWith(".pdf")
                    ? <embed src={previewUrl} type="application/pdf" style={{ width: "100%", height: 480 }} />
                    : <img src={previewUrl} alt={preview.name} style={{ maxWidth: "100%", maxHeight: 480, borderRadius: 6, objectFit: "contain" }} />
                  : <p style={{ color: "#94a3b8", fontSize: 13 }}>Cargando...</p>}
              </div>
              {analyses[preview.id] && (
                <div style={{ padding: 12, borderTop: "1px solid #f1f5f9" }}>
                  <AnalysisResult r={analyses[preview.id]} />
                </div>
              )}
              {!analyses[preview.id] && (
                <div style={{ padding: 12, borderTop: "1px solid #f1f5f9" }}>
                  <button onClick={() => analyzeOne(preview)} disabled={analyzing[preview.id]}
                    style={{ width: "100%", background: "#7c3aed", color: "#fff", border: "none", borderRadius: 8, padding: "10px", fontSize: 13, fontWeight: 600, cursor: "pointer" }}>
                    {analyzing[preview.id] ? "⏳ Analizando con IA..." : "🤖 Analizar esta evidencia"}
                  </button>
                </div>
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
