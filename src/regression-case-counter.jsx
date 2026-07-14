import React, { useState, useCallback } from "react";
import * as XLSX from "xlsx";
import { Upload, ScanLine, RotateCcw, AlertCircle, FileSpreadsheet, Layers, ListOrdered, Plus, X, ArrowRight } from "lucide-react";

// ---- design tokens -------------------------------------------------------
const T = {
  paper: "#EEF1EC",
  paperRaised: "#F6F8F5",
  ink: "#12262B",
  inkFaint: "#5C6B67",
  scan: "#2F7C6E",
  scanDim: "#DCE9E3",
  flag: "#D98E2B",
  flagDim: "#F6E6C9",
  hair: "#D2D8D0",
  danger: "#B3452F",
};

function normalizeHeader(h) {
  return String(h ?? "").trim().toLowerCase();
}

function guessColumn(headers, candidates) {
  const norm = headers.map(normalizeHeader);
  for (const c of candidates) {
    const idx = norm.findIndex((h) => h === c);
    if (idx !== -1) return idx;
  }
  for (const c of candidates) {
    const idx = norm.findIndex((h) => h.includes(c));
    if (idx !== -1) return idx;
  }
  return -1;
}

function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array" });
        resolve(wb);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error("read failed"));
    reader.readAsArrayBuffer(file);
  });
}

function computeResult(rows, storyColIdx, titleColIdx, keyword) {
  const dataRows = rows
    .slice(1)
    .filter((r) => r.some((c) => String(c).trim() !== ""))
    .filter((r) => String(r[storyColIdx] ?? "").trim() !== "");
  const kw = keyword.trim().toLowerCase();
  const order = [];
  const map = new Map();

  dataRows.forEach((r) => {
    const story = String(r[storyColIdx] ?? "").trim();
    const title = String(r[titleColIdx] ?? "");
    const matched = kw && title.toLowerCase().includes(kw);
    if (!map.has(story)) {
      map.set(story, { story, count: 0, total: 0 });
      order.push(story);
    }
    const entry = map.get(story);
    entry.total += 1;
    if (matched) entry.count += 1;
  });

  const list = order.map((s) => map.get(s)).sort((a, b) => b.count - a.count);
  const totalMatches = list.reduce((sum, e) => sum + e.count, 0);
  return { list, totalMatches, totalRows: dataRows.length };
}

export default function RegressionCaseCounter() {
  // top-level flow: mode -> upload -> configure (per file) -> result-single (sequential) -> summary
  const [stage, setStage] = useState("mode");
  const [mode, setMode] = useState(null); // "batch" | "sequential"
  const [keyword, setKeyword] = useState("regression");
  const [error, setError] = useState("");
  const [dragOver, setDragOver] = useState(false);

  // files awaiting processing (batch mode picks several at once)
  const [pendingFiles, setPendingFiles] = useState([]); // File objects
  // parsed files needing manual column selection: {fileName, headers, rows}
  const [configQueue, setConfigQueue] = useState([]);
  const [configIndex, setConfigIndex] = useState(0);
  const [storyColIdx, setStoryColIdx] = useState(-1);
  const [titleColIdx, setTitleColIdx] = useState(-1);

  // finished files: {fileName, list, totalMatches, totalRows}
  const [completedFiles, setCompletedFiles] = useState([]);
  const [busy, setBusy] = useState(false);

  const resetAll = () => {
    setStage("mode");
    setMode(null);
    setPendingFiles([]);
    setConfigQueue([]);
    setConfigIndex(0);
    setStoryColIdx(-1);
    setTitleColIdx(-1);
    setCompletedFiles([]);
    setError("");
  };

  const chooseMode = (m) => {
    setMode(m);
    setStage("upload");
    setError("");
  };

  // ---- file selection ----
  const onFilesChosen = (fileList) => {
    const files = Array.from(fileList || []);
    if (!files.length) return;
    if (mode === "sequential") {
      processSingleFile(files[0]);
    } else {
      setPendingFiles((prev) => [...prev, ...files]);
    }
  };

  const removePending = (idx) => {
    setPendingFiles((prev) => prev.filter((_, i) => i !== idx));
  };

  // ---- sequential mode: parse + auto-guess, fall back to manual config ----
  const processSingleFile = async (file) => {
    setError("");
    setBusy(true);
    try {
      const wb = await readWorkbook(file);
      const sheet = wb.SheetNames[0];
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheet], { header: 1, defval: "" });
      if (!rows.length) throw new Error("empty");
      const headers = rows[0];
      const sIdx = guessColumn(headers, ["reference", "story", "story name", "story id", "key"]);
      const tIdx = guessColumn(headers, ["title", "summary", "test title", "test case", "name"]);
      setConfigQueue([{ fileName: file.name, headers, rows }]);
      setConfigIndex(0);
      setStoryColIdx(sIdx);
      setTitleColIdx(tIdx);
      setStage("configure");
    } catch (e) {
      setError("Couldn't read that file. Make sure it's a valid .xlsx or .xls file.");
    } finally {
      setBusy(false);
    }
  };

  // ---- batch mode: parse all pending files, auto-resolve where possible ----
  const processBatch = async () => {
    if (!pendingFiles.length) return;
    setBusy(true);
    setError("");
    const autoResolved = [];
    const needsManual = [];
    try {
      for (const file of pendingFiles) {
        const wb = await readWorkbook(file);
        const sheet = wb.SheetNames[0];
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheet], { header: 1, defval: "" });
        if (!rows.length) continue;
        const headers = rows[0];
        const sIdx = guessColumn(headers, ["reference", "story", "story name", "story id", "key"]);
        const tIdx = guessColumn(headers, ["title", "summary", "test title", "test case", "name"]);
        if (sIdx !== -1 && tIdx !== -1) {
          const result = computeResult(rows, sIdx, tIdx, keyword);
          autoResolved.push({ fileName: file.name, ...result });
        } else {
          needsManual.push({ fileName: file.name, headers, rows, sIdx, tIdx });
        }
      }
      setCompletedFiles((prev) => [...prev, ...autoResolved]);
      if (needsManual.length) {
        setConfigQueue(needsManual);
        setConfigIndex(0);
        setStoryColIdx(needsManual[0].sIdx);
        setTitleColIdx(needsManual[0].tIdx);
        setStage("configure");
      } else {
        setStage("summary");
      }
      setPendingFiles([]);
    } catch (e) {
      setError("Something went wrong reading one of those files.");
    } finally {
      setBusy(false);
    }
  };

  // ---- confirm column choice for the file currently in configQueue ----
  const confirmConfig = () => {
    const current = configQueue[configIndex];
    const result = computeResult(current.rows, storyColIdx, titleColIdx, keyword);
    const finished = { fileName: current.fileName, ...result };
    setCompletedFiles((prev) => [...prev, finished]);

    const nextIndex = configIndex + 1;
    if (nextIndex < configQueue.length) {
      const next = configQueue[nextIndex];
      setConfigIndex(nextIndex);
      setStoryColIdx(next.sIdx ?? -1);
      setTitleColIdx(next.tIdx ?? -1);
    } else {
      // done with this queue
      if (mode === "sequential") {
        setStage("result-single");
      } else {
        setStage("summary");
      }
      setConfigQueue([]);
      setConfigIndex(0);
    }
  };

  const grandTotal = completedFiles.reduce((sum, f) => sum + f.totalMatches, 0);
  const grandRows = completedFiles.reduce((sum, f) => sum + f.totalRows, 0);
  const lastFile = completedFiles[completedFiles.length - 1];

  const downloadCsv = () => {
    const lines = ["File,Story Name,Regression Count"];
    completedFiles.forEach((f) => {
      f.list.forEach((e) => lines.push(`"${f.fileName.replace(/"/g, '""')}","${e.story.replace(/"/g, '""')}",${e.count}`));
    });
    lines.push(`,GRAND TOTAL,${grandTotal}`);
    const blob = new Blob([lines.join("\n")], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "regression-counts.csv";
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div
      style={{
        background: T.paper,
        minHeight: "100%",
        fontFamily: "'IBM Plex Sans', sans-serif",
        color: T.ink,
        padding: "28px 20px",
        boxSizing: "border-box",
      }}
    >
      <link rel="preconnect" href="https://fonts.googleapis.com" />
      <link
        href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500;600&family=Space+Grotesk:wght@500;600;700&display=swap"
        rel="stylesheet"
      />

      <div style={{ maxWidth: 880, margin: "0 auto" }}>
        {/* header */}
        <div style={{ display: "flex", alignItems: "baseline", justifyContent: "space-between", marginBottom: 22, flexWrap: "wrap", gap: 10 }}>
          <div>
            <div
              style={{
                fontFamily: "'Space Grotesk', sans-serif",
                fontWeight: 700,
                fontSize: 24,
                letterSpacing: "-0.01em",
                display: "flex",
                alignItems: "center",
                gap: 10,
              }}
            >
              <ScanLine size={22} color={T.scan} strokeWidth={2.4} />
              Regression Case Counter
            </div>
            <div style={{ fontSize: 13, color: T.inkFaint, marginTop: 3, fontFamily: "'IBM Plex Mono', monospace" }}>
              scans a title column, tallies matches per story
            </div>
          </div>
          {stage !== "mode" && (
            <button onClick={resetAll} style={ghostBtn}>
              <RotateCcw size={14} />
              Start over
            </button>
          )}
        </div>

        {/* ---- mode select ---- */}
        {stage === "mode" && (
          <div>
            <div style={{ fontSize: 14.5, marginBottom: 16, color: T.inkFaint }}>Do you have multiple files, or would you like to go one at a time?</div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
              <ModeCard
                icon={<Layers size={20} color={T.scan} />}
                title="Multiple files"
                desc="Upload them all together, review any unclear columns, then get one combined result."
                onClick={() => chooseMode("batch")}
              />
              <ModeCard
                icon={<ListOrdered size={20} color={T.scan} />}
                title="One at a time"
                desc="Upload a file, see its result, then decide whether to add another."
                onClick={() => chooseMode("sequential")}
              />
            </div>
          </div>
        )}

        {/* ---- upload ---- */}
        {stage === "upload" && (
          <div>
            <Field label="Keyword to match (case-insensitive, matches partial words too)">
              <input value={keyword} onChange={(e) => setKeyword(e.target.value)} style={{ ...selectStyle, fontFamily: "'IBM Plex Mono', monospace" }} />
            </Field>

            <div
              onDragOver={(e) => {
                e.preventDefault();
                setDragOver(true);
              }}
              onDragLeave={() => setDragOver(false)}
              onDrop={(e) => {
                e.preventDefault();
                setDragOver(false);
                onFilesChosen(e.dataTransfer.files);
              }}
              style={{
                background: T.paperRaised,
                border: `2px dashed ${dragOver ? T.scan : T.hair}`,
                borderRadius: 14,
                padding: "48px 24px",
                textAlign: "center",
                transition: "border-color 0.15s ease",
              }}
            >
              <div
                style={{
                  width: 52,
                  height: 52,
                  margin: "0 auto 16px",
                  borderRadius: 12,
                  background: T.scanDim,
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                }}
              >
                <Upload size={24} color={T.scan} />
              </div>
              <div style={{ fontSize: 16, fontWeight: 600, marginBottom: 6 }}>
                {mode === "batch" ? "Drop all your spreadsheets here" : "Drop your spreadsheet here"}
              </div>
              <div style={{ fontSize: 13, color: T.inkFaint, marginBottom: 18 }}>.xlsx or .xls — nothing leaves your browser</div>
              <label style={primaryBtnLabel}>
                {mode === "batch" ? "Choose files" : "Choose file"}
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  multiple={mode === "batch"}
                  onChange={(e) => onFilesChosen(e.target.files)}
                  style={{ display: "none" }}
                />
              </label>
            </div>

            {mode === "batch" && pendingFiles.length > 0 && (
              <div style={{ marginTop: 18 }}>
                <div style={{ fontSize: 12.5, color: T.inkFaint, marginBottom: 8, fontWeight: 500 }}>{pendingFiles.length} file(s) ready</div>
                <div style={{ display: "flex", flexDirection: "column", gap: 6, marginBottom: 16 }}>
                  {pendingFiles.map((f, i) => (
                    <div
                      key={f.name + i}
                      style={{
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "space-between",
                        background: T.paperRaised,
                        border: `1px solid ${T.hair}`,
                        borderRadius: 8,
                        padding: "8px 12px",
                        fontSize: 13,
                        fontFamily: "'IBM Plex Mono', monospace",
                      }}
                    >
                      <span style={{ display: "flex", alignItems: "center", gap: 8 }}>
                        <FileSpreadsheet size={14} color={T.inkFaint} />
                        {f.name}
                      </span>
                      <button onClick={() => removePending(i)} style={iconBtn}>
                        <X size={14} />
                      </button>
                    </div>
                  ))}
                </div>
                <button onClick={processBatch} style={primaryBtn} disabled={busy}>
                  {busy ? "Reading files…" : `Process ${pendingFiles.length} file(s)`} <ArrowRight size={14} style={{ marginLeft: 6, verticalAlign: -2 }} />
                </button>
              </div>
            )}

            {error && (
              <div style={{ marginTop: 16, color: T.danger, fontSize: 13, display: "flex", alignItems: "center", gap: 6 }}>
                <AlertCircle size={14} /> {error}
              </div>
            )}
          </div>
        )}

        {/* ---- configure (per file needing manual column pick) ---- */}
        {stage === "configure" && configQueue[configIndex] && (
          <div>
            <div
              style={{
                display: "flex",
                alignItems: "center",
                gap: 8,
                fontSize: 13,
                color: T.inkFaint,
                marginBottom: 16,
                fontFamily: "'IBM Plex Mono', monospace",
              }}
            >
              <FileSpreadsheet size={15} />
              {configQueue[configIndex].fileName}
              {configQueue.length > 1 && <span> · file {configIndex + 1} of {configQueue.length}</span>}
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginBottom: 14 }}>
              <Field label="Story name column (reference)">
                <select value={storyColIdx} onChange={(e) => setStoryColIdx(Number(e.target.value))} style={selectStyle}>
                  <option value={-1}>Select a column…</option>
                  {configQueue[configIndex].headers.map((h, i) => (
                    <option key={i} value={i}>
                      {String(h) || `Column ${i + 1}`}
                    </option>
                  ))}
                </select>
              </Field>
              <Field label="Title column to scan">
                <select value={titleColIdx} onChange={(e) => setTitleColIdx(Number(e.target.value))} style={selectStyle}>
                  <option value={-1}>Select a column…</option>
                  {configQueue[configIndex].headers.map((h, i) => (
                    <option key={i} value={i}>
                      {String(h) || `Column ${i + 1}`}
                    </option>
                  ))}
                </select>
              </Field>
            </div>

            {storyColIdx === -1 || titleColIdx === -1 ? (
              <div style={{ fontSize: 13, color: T.inkFaint }}>Pick both columns to continue.</div>
            ) : (
              <button onClick={confirmConfig} style={primaryBtn}>
                <ScanLine size={15} style={{ marginRight: 7, verticalAlign: -2 }} />
                {configIndex + 1 < configQueue.length ? "Scan and continue" : "Scan"}
              </button>
            )}
          </div>
        )}

        {/* ---- single result (sequential mode) + ask for next ---- */}
        {stage === "result-single" && lastFile && (
          <div>
            <FileResultBlock file={lastFile} />
            <div
              style={{
                marginTop: 22,
                background: T.paperRaised,
                border: `1px solid ${T.hair}`,
                borderRadius: 12,
                padding: "16px 18px",
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
                flexWrap: "wrap",
                gap: 12,
              }}
            >
              <div style={{ fontSize: 14, fontWeight: 500 }}>Do you have another file to upload?</div>
              <div style={{ display: "flex", gap: 10 }}>
                <button
                  onClick={() => {
                    setStage("upload");
                    setError("");
                  }}
                  style={primaryBtn}
                >
                  <Plus size={14} style={{ marginRight: 6, verticalAlign: -2 }} />
                  Yes, add another
                </button>
                <button onClick={() => setStage("summary")} style={secondaryBtn}>
                  No, show final total
                </button>
              </div>
            </div>
          </div>
        )}

        {/* ---- final summary across all processed files ---- */}
        {stage === "summary" && (
          <div>
            <div style={{ display: "flex", gap: 14, marginBottom: 20, flexWrap: "wrap" }}>
              <StatCard label="Files processed" value={completedFiles.length} />
              <StatCard label="Rows scanned" value={grandRows} />
              <StatCard label="Grand total regressions" value={grandTotal} accent />
            </div>

            {completedFiles.map((f, i) => <FileResultBlock key={f.fileName + i} file={f} />)}

            <div
              style={{
                marginTop: 8,
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                background: T.ink,
                color: T.paperRaised,
                borderRadius: 10,
                padding: "14px 18px",
                fontFamily: "'IBM Plex Mono', monospace",
              }}
            >
              <span style={{ fontWeight: 600 }}>Grand total across {completedFiles.length} file(s)</span>
              <span style={{ fontSize: 20, fontWeight: 700 }}>{grandTotal}</span>
            </div>

            <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
              <button onClick={downloadCsv} style={primaryBtn}>
                Export CSV
              </button>
              <button onClick={resetAll} style={secondaryBtn}>
                Start over
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

function FileResultBlock({ file }) {
  const maxCount = file.list[0]?.count || 1;
  return (
    <div style={{ marginBottom: 20 }}>
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 8,
          fontSize: 13,
          color: T.inkFaint,
          marginBottom: 8,
          fontFamily: "'IBM Plex Mono', monospace",
        }}
      >
        <FileSpreadsheet size={14} />
        {file.fileName} · {file.totalRows} rows · {file.totalMatches} matches
      </div>
      <div style={{ background: T.paperRaised, borderRadius: 12, border: `1px solid ${T.hair}`, overflow: "hidden" }}>
        <div
          style={{
            display: "grid",
            gridTemplateColumns: "1fr 120px 100px",
            padding: "10px 16px",
            fontSize: 11.5,
            letterSpacing: "0.06em",
            textTransform: "uppercase",
            color: T.inkFaint,
            borderBottom: `1px solid ${T.hair}`,
            fontWeight: 600,
          }}
        >
          <div>Story name</div>
          <div style={{ textAlign: "right" }}>Regression count</div>
          <div style={{ textAlign: "right" }}>Share</div>
        </div>
        <div>
          {file.list.map((e, i) => {
            const pct = e.count === 0 ? 0 : Math.max(6, (e.count / maxCount) * 100);
            return (
              <div
                key={e.story + i}
                style={{
                  display: "grid",
                  gridTemplateColumns: "1fr 120px 100px",
                  padding: "10px 16px",
                  alignItems: "center",
                  borderBottom: i === file.list.length - 1 ? "none" : `1px solid ${T.hair}`,
                  fontSize: 13.5,
                }}
              >
                <div style={{ paddingRight: 12, wordBreak: "break-word" }}>{e.story}</div>
                <div style={{ textAlign: "right", fontFamily: "'IBM Plex Mono', monospace", fontWeight: 600 }}>
                  <span
                    style={{
                      background: e.count > 0 ? T.flagDim : "transparent",
                      color: e.count > 0 ? "#8A5A12" : T.inkFaint,
                      borderRadius: 6,
                      padding: "2px 8px",
                    }}
                  >
                    {e.count}
                  </span>
                </div>
                <div style={{ display: "flex", justifyContent: "flex-end" }}>
                  <div style={{ width: 70, height: 6, background: T.hair, borderRadius: 3, overflow: "hidden" }}>
                    <div style={{ width: `${pct}%`, height: "100%", background: e.count > 0 ? T.flag : "transparent" }} />
                  </div>
                </div>
              </div>
            );
          })}
        </div>
        <div
          style={{
            display: "grid",
            gridTemplateColumns: "1fr 120px 100px",
            padding: "10px 16px",
            fontSize: 13,
            fontWeight: 700,
            borderTop: `1px solid ${T.hair}`,
            background: "#E7EDE7",
          }}
        >
          <div>Subtotal</div>
          <div style={{ textAlign: "right", fontFamily: "'IBM Plex Mono', monospace" }}>{file.totalMatches}</div>
          <div />
        </div>
      </div>
    </div>
  );
}

function ModeCard({ icon, title, desc, onClick }) {
  return (
    <button
      onClick={onClick}
      style={{
        textAlign: "left",
        background: T.paperRaised,
        border: `1px solid ${T.hair}`,
        borderRadius: 12,
        padding: "18px 16px",
        cursor: "pointer",
        display: "flex",
        flexDirection: "column",
        gap: 10,
      }}
    >
      <div style={{ width: 38, height: 38, borderRadius: 9, background: T.scanDim, display: "flex", alignItems: "center", justifyContent: "center" }}>
        {icon}
      </div>
      <div style={{ fontWeight: 600, fontSize: 15 }}>{title}</div>
      <div style={{ fontSize: 12.5, color: T.inkFaint, lineHeight: 1.45 }}>{desc}</div>
    </button>
  );
}

function Field({ label, children }) {
  return (
    <div style={{ marginBottom: 14 }}>
      <div style={{ fontSize: 12.5, color: "#5C6B67", marginBottom: 6, fontWeight: 500 }}>{label}</div>
      {children}
    </div>
  );
}

function StatCard({ label, value, accent }) {
  return (
    <div
      style={{
        background: accent ? "#DCE9E3" : "#F6F8F5",
        border: `1px solid ${accent ? "#2F7C6E33" : "#D2D8D0"}`,
        borderRadius: 10,
        padding: "12px 18px",
        minWidth: 130,
      }}
    >
      <div style={{ fontSize: 11.5, color: "#5C6B67", textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 4 }}>{label}</div>
      <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 26, fontWeight: 700, color: accent ? "#215E52" : "#12262B" }}>{value}</div>
    </div>
  );
}

const selectStyle = {
  width: "100%",
  padding: "9px 12px",
  borderRadius: 8,
  border: "1px solid #D2D8D0",
  background: "#F6F8F5",
  fontSize: 13.5,
  color: "#12262B",
  boxSizing: "border-box",
};

const primaryBtn = {
  background: "#12262B",
  color: "#F6F8F5",
  border: "none",
  borderRadius: 8,
  padding: "10px 18px",
  fontSize: 13.5,
  fontWeight: 600,
  cursor: "pointer",
};

const primaryBtnLabel = {
  display: "inline-block",
  background: "#12262B",
  color: "#F6F8F5",
  borderRadius: 8,
  padding: "9px 18px",
  fontSize: 13.5,
  fontWeight: 600,
  cursor: "pointer",
};

const secondaryBtn = {
  background: "transparent",
  color: "#12262B",
  border: "1px solid #D2D8D0",
  borderRadius: 8,
  padding: "10px 18px",
  fontSize: 13.5,
  fontWeight: 600,
  cursor: "pointer",
};

const ghostBtn = {
  display: "flex",
  alignItems: "center",
  gap: 6,
  background: "transparent",
  border: `1px solid #D2D8D0`,
  borderRadius: 8,
  padding: "7px 12px",
  fontSize: 13,
  color: "#5C6B67",
  cursor: "pointer",
  fontFamily: "'IBM Plex Sans', sans-serif",
};

const iconBtn = {
  background: "transparent",
  border: "none",
  color: "#5C6B67",
  cursor: "pointer",
  display: "flex",
  alignItems: "center",
};