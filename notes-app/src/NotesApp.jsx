import { useState, useEffect, useCallback } from "react";

// ─── CONFIGURATION ────────────────────────────────────────────────────────────
// 1. Go to https://console.cloud.google.com/
// 2. Create a project → Enable "Google Drive API" + "Google Picker API"
// 3. Create OAuth 2.0 credentials (Web application)
//    - Authorized JS origins: http://localhost:3000
// 4. Create an API Key (restrict to Drive + Picker APIs)
// 5. Create a file at the project root named .env and add:
//    VITE_GOOGLE_CLIENT_ID=your-client-id
//    VITE_GOOGLE_API_KEY=your-api-key
// 6. Restart the Vite dev server after updating .env.
const CLIENT_ID = import.meta.env.VITE_GOOGLE_CLIENT_ID;
const API_KEY = import.meta.env.VITE_GOOGLE_API_KEY;
const HAS_GOOGLE_CONFIG = Boolean(CLIENT_ID && API_KEY);
const SCOPES = "https://www.googleapis.com/auth/drive.file";
const FOLDER_NAME = "My Notes";

// ─── DOCX BUILDER (pure browser, no backend) ─────────────────────────────────
// Uses docx.js loaded via CDN in index.html (see README at bottom)
function buildDocx(title, rawText) {
  const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType } =
    window.docx;

  const lines = rawText.split("\n").filter((l) => l.trim() !== "");

  const children = [
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [new TextRun({ text: title, bold: true })],
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: `Created: ${new Date().toLocaleString()}`,
          color: "888888",
          size: 20,
          italics: true,
        }),
      ],
    }),
    new Paragraph({ children: [new TextRun("")] }), // spacer
    ...lines.map(
      (line) =>
        new Paragraph({
          children: [new TextRun({ text: line, size: 24 })],
          spacing: { after: 160 },
        })
    ),
  ];

  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            size: { width: 12240, height: 15840 },
            margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
          },
        },
        children,
      },
    ],
  });
return Packer.toBlob(doc);
}

// Append text to existing doc by downloading, adding a section, re-uploading
async function appendToExistingDoc(fileId, newTitle, rawText, accessToken) {
  const { Document, Packer, Paragraph, TextRun, HeadingLevel, BorderStyle } =
    window.docx;

  // Download existing file bytes
  const dlRes = await fetch(
    `https://www.googleapis.com/drive/v3/files/${fileId}?alt=media`,
    { headers: { Authorization: `Bearer ${accessToken}` } }
  );
  const existingBuffer = await dlRes.arrayBuffer();

  // We can't truly "merge" two docx files in pure browser JS,
  // so we append as a clearly-marked new section using mammoth to extract text
  // then rebuild the whole doc.
  // For simplicity: download existing text, append new content.
  const mammoth = window.mammoth;
  const { value: existingText } = await mammoth.extractRawText({
    arrayBuffer: existingBuffer,
  });

  const lines = (existingText + "\n\n---\n\n" + `## ${newTitle}\n\n` + rawText)
    .split("\n")
    .filter((l) => l.trim() !== "");

  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            size: { width: 12240, height: 15840 },
            margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
          },
        },
        children: lines.map(
          (line) =>
            new Paragraph({
              children: [new TextRun({ text: line, size: 24 })],
              spacing: { after: 160 },
            })
        ),
      },
    ],
  });

return Packer.toBlob(doc);
}

// ─── GOOGLE DRIVE HELPERS ─────────────────────────────────────────────────────
async function findOrCreateFolder(accessToken) {
  // Search for existing folder
  const searchRes = await fetch(
    `https://www.googleapis.com/drive/v3/files?q=name='${FOLDER_NAME}' and mimeType='application/vnd.google-apps.folder' and trashed=false&fields=files(id,name)`,
    { headers: { Authorization: `Bearer ${accessToken}` } }
  );
  const { files } = await searchRes.json();
  if (files && files.length > 0) return files[0].id;

  // Create new folder
  const createRes = await fetch("https://www.googleapis.com/drive/v3/files", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      name: FOLDER_NAME,
      mimeType: "application/vnd.google-apps.folder",
    }),
  });
  const folder = await createRes.json();
  return folder.id;
}

async function listDocsInFolder(accessToken, folderId) {
  const res = await fetch(
    `https://www.googleapis.com/drive/v3/files?q='${folderId}' in parents and mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document' and trashed=false&fields=files(id,name,modifiedTime)&orderBy=modifiedTime desc`,
    { headers: { Authorization: `Bearer ${accessToken}` } }
  );
  const { files } = await res.json();
  return files || [];
}

// async function uploadDocx(accessToken, folderId, fileName, arrayBuffer, existingFileId = null) {
//   const boundary = "-------notesapp_boundary";
//   const contentType =
//     "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

//   const metadata = JSON.stringify(
//     existingFileId
//       ? { name: fileName }
//       : { name: fileName, parents: [folderId] }
//   );

//   const body = [
//     `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${metadata}\r\n`,
//     `--${boundary}\r\nContent-Type: ${contentType}\r\n\r\n`,
//     new Uint8Array(arrayBuffer),
//     `\r\n--${boundary}--`,
//   ];

//   const blob = new Blob(body, { type: `multipart/related; boundary="${boundary}"` });

//   const url = existingFileId
//     ? `https://www.googleapis.com/upload/drive/v3/files/${existingFileId}?uploadType=multipart`
//     : `https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart`;

//   const method = existingFileId ? "PATCH" : "POST";

//   const res = await fetch(url, {
//     method,
//     headers: { Authorization: `Bearer ${accessToken}` },
//     body: blob,
//   });
//   return res.json();
// }
async function uploadDocx(accessToken, folderId, fileName, blob, existingFileId = null) {
  const metadata = JSON.stringify(
    existingFileId
      ? { name: fileName }
      : { name: fileName, parents: [folderId] }
  );

  const formData = new FormData();
  formData.append("metadata", new Blob([metadata], { type: "application/json" }));
  formData.append("file", blob, fileName);

  const url = existingFileId
    ? `https://www.googleapis.com/upload/drive/v3/files/${existingFileId}?uploadType=multipart`
    : `https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart`;

  const method = existingFileId ? "PATCH" : "POST";

  const res = await fetch(url, {
    method,
    headers: { Authorization: `Bearer ${accessToken}` },
    body: formData,
  });
  return res.json();
}

// ─── MAIN APP ─────────────────────────────────────────────────────────────────
export default function NotesApp() {
  const [gapiReady, setGapiReady] = useState(false);
  const [accessToken, setAccessToken] = useState(null);
  const [folderId, setFolderId] = useState(null);
  const [existingDocs, setExistingDocs] = useState([]);

  const [title, setTitle] = useState("");
  const [rawText, setRawText] = useState("");
  const [mode, setMode] = useState("new"); // "new" | "append"
  const [selectedDoc, setSelectedDoc] = useState("");
  const [status, setStatus] = useState({ type: "", msg: "" });
  const [loading, setLoading] = useState(false);

  // Load Google Identity Services
  useEffect(() => {
    const script = document.createElement("script");
    script.src = "https://accounts.google.com/gsi/client";
    script.onload = () => setGapiReady(true);
    document.head.appendChild(script);
  }, []);

  const signIn = useCallback(() => {
    if (!window.google) return;
    const client = window.google.accounts.oauth2.initTokenClient({
      client_id: CLIENT_ID,
      scope: SCOPES,
      callback: async (response) => {
        if (response.error) {
          setStatus({ type: "error", msg: "Sign-in failed: " + response.error });
          return;
        }
        setAccessToken(response.access_token);
        setStatus({ type: "info", msg: "Signed in! Setting up Notes folder…" });
        try {
          const id = await findOrCreateFolder(response.access_token);
          setFolderId(id);
          const docs = await listDocsInFolder(response.access_token, id);
          setExistingDocs(docs);
          setStatus({ type: "success", msg: `✓ Ready! Folder "${FOLDER_NAME}" found/created.` });
        } catch (e) {
          setStatus({ type: "error", msg: "Drive error: " + e.message });
        }
      },
    });
    client.requestAccessToken();
  }, [gapiReady]);

  const refreshDocs = useCallback(async () => {
    if (!accessToken || !folderId) return;
    const docs = await listDocsInFolder(accessToken, folderId);
    setExistingDocs(docs);
  }, [accessToken, folderId]);

  const handleSave = async () => {
    if (!title.trim()) return setStatus({ type: "error", msg: "Title is required." });
    if (!rawText.trim()) return setStatus({ type: "error", msg: "Please enter some text." });
    if (mode === "append" && !selectedDoc) return setStatus({ type: "error", msg: "Select a document to append to." });

    setLoading(true);
    setStatus({ type: "info", msg: "Generating Word document…" });

    try {
      let arrayBuffer;
      let existingFileId = null;
      let fileName;

      if (mode === "new") {
        arrayBuffer = await buildDocx(title, rawText);
        fileName = `${title}.docx`;
      } else {
        const doc = existingDocs.find((d) => d.id === selectedDoc);
        existingFileId = doc.id;
        fileName = doc.name;
        setStatus({ type: "info", msg: "Downloading existing doc to append…" });
        arrayBuffer = await appendToExistingDoc(existingFileId, title, rawText, accessToken);
      }

      setStatus({ type: "info", msg: "Uploading to Google Drive…" });
      const result = await uploadDocx(accessToken, folderId, fileName, arrayBuffer, existingFileId);

      if (result.id) {
        setStatus({
          type: "success",
          msg: `✓ Saved "${result.name}" to your "${FOLDER_NAME}" folder on Drive!`,
        });
        setTitle("");
        setRawText("");
        await refreshDocs();
      } else {
        throw new Error(JSON.stringify(result));
      }
    } catch (e) {
      setStatus({ type: "error", msg: "Error: " + e.message });
    } finally {
      setLoading(false);
    }
  };

  // ── UI ──────────────────────────────────────────────────────────────────────
  return (
    <div style={styles.page}>
      <div style={styles.card}>
        {/* Header */}
        <div style={styles.header}>
          <span style={styles.logo}>📝</span>
          <div>
            <h1 style={styles.h1}>Notes to Drive</h1>
            <p style={styles.subtitle}>Type raw text → formatted Word doc → saved to Google Drive</p>
          </div>
        </div>

        {/* Sign In */}
        {!accessToken ? (
          <div style={styles.center}>
            <button
              style={{ ...styles.btn, ...styles.btnPrimary }}
              onClick={signIn}
              disabled={!gapiReady || !HAS_GOOGLE_CONFIG}
            >
              {gapiReady ? "🔐 Sign in with Google" : "Loading…"}
            </button>
            <p style={styles.hint}>
              First time? Make sure you've set up your Google Cloud credentials.<br />
              {HAS_GOOGLE_CONFIG ? "See the setup guide in the code comments." : "Create a .env file with VITE_GOOGLE_CLIENT_ID and VITE_GOOGLE_API_KEY."}
            </p>
          </div>
        ) : (
          <>
            {/* Mode Toggle */}
            <div style={styles.toggleRow}>
              <button
                style={{ ...styles.toggle, ...(mode === "new" ? styles.toggleActive : {}) }}
                onClick={() => setMode("new")}
              >
                ✨ New Document
              </button>
              <button
                style={{ ...styles.toggle, ...(mode === "append" ? styles.toggleActive : {}) }}
                onClick={() => setMode("append")}
              >
                ➕ Append to Existing
              </button>
            </div>

            {/* Append: Select existing doc */}
            {mode === "append" && (
              <div style={styles.field}>
                <label style={styles.label}>Select Document to Append To</label>
                {existingDocs.length === 0 ? (
                  <p style={styles.hint}>No docs found in "{FOLDER_NAME}" folder yet.</p>
                ) : (
                  <select
                    style={styles.select}
                    value={selectedDoc}
                    onChange={(e) => setSelectedDoc(e.target.value)}
                  >
                    <option value="">-- Choose a document --</option>
                    {existingDocs.map((doc) => (
                      <option key={doc.id} value={doc.id}>
                        {doc.name} ({new Date(doc.modifiedTime).toLocaleDateString()})
                      </option>
                    ))}
                  </select>
                )}
              </div>
            )}

            {/* Section Title */}
            <div style={styles.field}>
              <label style={styles.label}>
                {mode === "new" ? "Document Title" : "Section / Note Title"}
              </label>
              <input
                style={styles.input}
                placeholder={mode === "new" ? "e.g. Meeting Notes - May 2026" : "e.g. Follow-up items"}
                value={title}
                onChange={(e) => setTitle(e.target.value)}
              />
            </div>

            {/* Raw Text */}
            <div style={styles.field}>
              <label style={styles.label}>Raw Text</label>
              <textarea
                style={styles.textarea}
                placeholder={"Paste or type your notes here.\nEach line becomes a paragraph in the Word doc.\n\nYou can write freely — no formatting needed!"}
                value={rawText}
                onChange={(e) => setRawText(e.target.value)}
                rows={10}
              />
              <p style={styles.hint}>{rawText.split("\n").filter((l) => l.trim()).length} lines • {rawText.length} characters</p>
            </div>

            {/* Save Button */}
            <button
              style={{
                ...styles.btn,
                ...styles.btnPrimary,
                ...(loading ? styles.btnDisabled : {}),
              }}
              onClick={handleSave}
              disabled={loading}
            >
              {loading ? "⏳ Saving…" : mode === "new" ? "💾 Save New Doc to Drive" : "📎 Append to Doc on Drive"}
            </button>

            {/* Existing docs list */}
            {existingDocs.length > 0 && (
              <div style={styles.docList}>
                <div style={styles.docListHeader}>
                  <span>📁 Notes Folder ({existingDocs.length} docs)</span>
                  <button style={styles.refreshBtn} onClick={refreshDocs}>↻ Refresh</button>
                </div>
                {existingDocs.slice(0, 8).map((doc) => (
                  <div key={doc.id} style={styles.docItem}>
                    <span>📄 {doc.name}</span>
                    <a
                      href={`https://drive.google.com/file/d/${doc.id}/view`}
                      target="_blank"
                      rel="noreferrer"
                      style={styles.openLink}
                    >
                      Open ↗
                    </a>
                  </div>
                ))}
              </div>
            )}
          </>
        )}

        {/* Status Message */}
        {status.msg && (
          <div
            style={{
              ...styles.statusBox,
              background:
                status.type === "success" ? "#d1fae5"
                : status.type === "error" ? "#fee2e2"
                : "#dbeafe",
              color:
                status.type === "success" ? "#065f46"
                : status.type === "error" ? "#991b1b"
                : "#1e40af",
            }}
          >
            {status.msg}
          </div>
        )}
      </div>
    </div>
  );
}

// ─── STYLES ───────────────────────────────────────────────────────────────────
const styles = {
  page: {
    minHeight: "100vh",
    background: "linear-gradient(135deg, #f0f4ff 0%, #faf5ff 100%)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    padding: "24px",
    fontFamily: "'Segoe UI', system-ui, sans-serif",
  },
  card: {
    background: "#fff",
    borderRadius: "20px",
    boxShadow: "0 8px 40px rgba(0,0,0,0.10)",
    padding: "36px",
    maxWidth: "580px",
    width: "100%",
  },
  header: {
    display: "flex",
    alignItems: "center",
    gap: "16px",
    marginBottom: "28px",
  },
  logo: { fontSize: "42px" },
  h1: { margin: 0, fontSize: "22px", fontWeight: 700, color: "#1e1e2e" },
  subtitle: { margin: "4px 0 0", fontSize: "13px", color: "#6b7280" },
  center: { textAlign: "center", padding: "20px 0" },
  field: { marginBottom: "18px" },
  label: { display: "block", fontWeight: 600, marginBottom: "6px", color: "#374151", fontSize: "14px" },
  input: {
    width: "100%", padding: "10px 14px", borderRadius: "10px",
    border: "1.5px solid #e5e7eb", fontSize: "15px", outline: "none",
    boxSizing: "border-box", transition: "border-color 0.2s",
  },
  textarea: {
    width: "100%", padding: "12px 14px", borderRadius: "10px",
    border: "1.5px solid #e5e7eb", fontSize: "14px", outline: "none",
    boxSizing: "border-box", resize: "vertical", lineHeight: 1.6,
    fontFamily: "'Segoe UI', system-ui, sans-serif",
  },
  select: {
    width: "100%", padding: "10px 14px", borderRadius: "10px",
    border: "1.5px solid #e5e7eb", fontSize: "14px", outline: "none",
    boxSizing: "border-box", background: "#fff",
  },
  hint: { fontSize: "12px", color: "#9ca3af", marginTop: "6px" },
  toggleRow: { display: "flex", gap: "10px", marginBottom: "22px" },
  toggle: {
    flex: 1, padding: "10px", borderRadius: "10px", border: "1.5px solid #e5e7eb",
    background: "#f9fafb", cursor: "pointer", fontWeight: 500, fontSize: "14px",
    color: "#374151", transition: "all 0.2s",
  },
  toggleActive: {
    background: "#4f46e5", color: "#fff", border: "1.5px solid #4f46e5",
  },
  btn: {
    width: "100%", padding: "13px", borderRadius: "12px", border: "none",
    fontSize: "15px", fontWeight: 600, cursor: "pointer", transition: "all 0.2s",
    marginTop: "4px",
  },
  btnPrimary: { background: "#4f46e5", color: "#fff" },
  btnDisabled: { background: "#9ca3af", cursor: "not-allowed" },
  statusBox: {
    marginTop: "20px", padding: "14px 16px", borderRadius: "12px",
    fontSize: "14px", fontWeight: 500,
  },
  docList: {
    marginTop: "24px", border: "1.5px solid #e5e7eb",
    borderRadius: "12px", overflow: "hidden",
  },
  docListHeader: {
    padding: "10px 14px", background: "#f9fafb", fontWeight: 600,
    fontSize: "13px", color: "#374151", display: "flex", justifyContent: "space-between",
    alignItems: "center",
  },
  refreshBtn: {
    background: "none", border: "none", cursor: "pointer",
    color: "#4f46e5", fontWeight: 600, fontSize: "13px",
  },
  docItem: {
    padding: "10px 14px", borderTop: "1px solid #f3f4f6",
    fontSize: "13px", color: "#374151", display: "flex",
    justifyContent: "space-between", alignItems: "center",
  },
  openLink: { color: "#4f46e5", textDecoration: "none", fontSize: "12px", fontWeight: 600 },
};
