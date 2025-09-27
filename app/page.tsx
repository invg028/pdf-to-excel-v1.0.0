"use client";

import { useRef, useState } from "react";
import * as XLSX from "xlsx";
import * as pdfjsLib from "pdfjs-dist/legacy/build/pdf";
import { createWorker } from "tesseract.js";

// Pakai CDN worker untuk PDF.js agar ringan
// @ts-ignore
pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.8.69/pdf.worker.min.js";

type TextCell = { x: number; y: number; w?: number; h?: number; text: string };

export default function Page() {
  const inputRef = useRef<HTMLInputElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const [busy, setBusy] = useState(false);
  const [log, setLog] = useState<string[]>([]);

  const [tolY, setTolY] = useState(3);   // toleransi baris (px)
  const [gapX, setGapX] = useState(22);  // gap kolom (px)
  const [scale, setScale] = useState(2); // skala render OCR

  const pushLog = (s: string) => setLog((p) => [...p, s]);

  async function extractDigital(pdf: any) {
    let rows: string[][] = [];
    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
      const page = await pdf.getPage(pageNum);
      const content = await page.getTextContent();
      const items: any[] = content.items || [];
      const cells: TextCell[] = [];
      for (const it of items) {
        const s = it?.str?.trim?.();
        if (!s) continue;
        const tr = it?.transform;
        let x = 0, y = 0;
        if (Array.isArray(tr) && tr.length >= 6) { x = tr[4]; y = tr[5]; }
        cells.push({ x, y, text: s });
      }
      if (!cells.length) continue;
      cells.sort((a, b) => b.y - a.y || a.x - b.x);
      const lineBuckets: TextCell[][] = [];
      for (const c of cells) {
        const row = lineBuckets.find(r => Math.abs(r[0].y - c.y) <= tolY);
        if (row) row.push(c); else lineBuckets.push([c]);
      }
      for (const line of lineBuckets) {
        line.sort((a, b) => a.x - b.x);
        const cols: string[] = [];
        let cur = line[0].text;
        for (let i = 1; i < line.length; i++) {
          const gap = line[i].x - line[i - 1].x;
          if (gap > gapX) { cols.push(cur); cur = line[i].text; }
          else { cur += (cur ? " " : "") + line[i].text; }
        }
        cols.push(cur);
        const joined = cols.join("").trim();
        if (joined) rows.push(cols);
      }
      pushLog(`Digital: selesai ekstraksi halaman ${pageNum}`);
    }
    return rows;
  }

  async function extractOCR(pdf: any) {
    let rows: string[][] = [];
    const worker: any = await createWorker({ logger: () => {} });
    await worker.loadLanguage("ind+eng");
    await worker.initialize("ind+eng");

    const canvas = canvasRef.current!;
    const ctx = canvas.getContext("2d")!;

    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
      const page = await pdf.getPage(pageNum);
      const viewport = page.getViewport({ scale });
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      ctx.clearRect(0, 0, canvas.width, canvas.height);
      await page.render({ canvasContext: ctx, viewport }).promise;

    const { data } = await worker.recognize(canvas);
    const words = (data?.words || []) as Array<any>;
    const cells: TextCell[] = words
      .filter(w => (w?.text || "").trim().length)
      .map(w => ({
        x: w.bbox.x0,
        y: canvas.height - w.bbox.y1,
        w: w.bbox.x1 - w.bbox.x0,
        h: w.bbox.y1 - w.bbox.y0,
        text: (w.text || "").trim()
      }));

      cells.sort((a, b) => b.y - a.y || a.x - b.x);
      const lineBuckets: TextCell[][] = [];
      for (const c of cells) {
        const row = lineBuckets.find(r => Math.abs(r[0].y - c.y) <= Math.max(tolY * scale, 6));
        if (row) row.push(c); else lineBuckets.push([c]);
      }
      for (const line of lineBuckets) {
        line.sort((a, b) => a.x - b.x);
        const cols: string[] = [];
        let cur = line[0].text;
        for (let i = 1; i < line.length; i++) {
          const gap = line[i].x - line[i - 1].x;
          if (gap > gapX * scale) { cols.push(cur); cur = line[i].text; }
          else { cur += (cur ? " " : "") + line[i].text; }
        }
        cols.push(cur);
        const joined = cols.join("").trim();
        if (joined) rows.push(cols);
      }
      pushLog(`OCR: selesai ekstraksi halaman ${pageNum}`);
    }
    await worker.terminate();
    return rows;
  }

  async function handleFile(files: FileList | null) {
    if (!files || !files[0]) return;
    setBusy(true); setLog([]);
    try {
      const file = files[0];
      pushLog(`Memuat: ${file.name}`);
      const buf = await file.arrayBuffer();
      const pdf = await pdfjsLib.getDocument({ data: buf }).promise;

      pushLog("Mencoba mode digital…");
      const rowsDigital = await extractDigital(pdf);

      let finalRows: string[][] = rowsDigital;
      if (rowsDigital.length === 0) {
        pushLog("Hasil kosong. Beralih ke OCR…");
        finalRows = await extractOCR(pdf);
      } else {
        const tokens = rowsDigital.reduce((n, r) => n + r.length, 0);
        if (tokens < 5) {
          pushLog("Ekstraksi digital minim. Menjalankan OCR untuk perbandingan…");
          const ocrRows = await extractOCR(pdf);
          if (ocrRows?.length > rowsDigital.length) finalRows = ocrRows;
        }
      }

      if (!finalRows.length) { pushLog("Tidak ada teks terdeteksi."); return; }

      const maxCols = finalRows.reduce((m, r) => Math.max(m, r.length), 0);
      const padded = finalRows.map(r => { const c = r.slice(); while (c.length < maxCols) c.push(""); return c; });

      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet(padded);
      XLSX.utils.book_append_sheet(wb, ws, "Hasil");
      const outName = file.name.replace(/\.pdf$/i, "") + ".xlsx";
      XLSX.writeFile(wb, outName);
      pushLog(`Berhasil ekspor: ${outName}`);
    } catch (e: any) {
      console.error(e);
      pushLog("Gagal: " + (e?.message || String(e)));
    } finally {
      setBusy(false);
    }
  }

  return (
    <div>
      <h1 style={{ fontSize: 28, fontWeight: 800, marginBottom: 8 }}>PDF → Excel Pro (Dengan OCR)</h1>
      <p style={{ opacity: 0.85, marginBottom: 18 }}>
        Konversi PDF <b>digital</b> atau <b>scan</b> (OCR Tesseract.js) ke Excel (XLSX). Semua proses <i>client-side</i>.
      </p>

      <div style={{ display: "grid", gap: 16, padding: 16, background: "#111827", border: "1px solid #223", borderRadius: 16 }}>
        <label style={{ display: "block" }}>
          <span style={{ display: "block", marginBottom: 6, fontWeight: 600 }}>Pilih PDF</span>
          <input
            ref={inputRef}
            type="file"
            accept="application/pdf"
            onChange={(e) => handleFile(e.target.files)}
            disabled={busy}
            style={{ padding: 8, borderRadius: 8, border: "1px solid #334", background: "#0b1220" }}
          />
        </label>

        <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
          <div>
            <label>Toleransi Baris (Y)</label>
            <input type="number" value={tolY} min={1} onChange={(e)=>setTolY(Number(e.target.value)||1)}
              disabled={busy} style={{ marginLeft: 8, width: 90, padding: 6, borderRadius: 8, border: "1px solid #334", background: "#0b1220" }} />
          </div>
          <div>
            <label>Gap Kolom (X)</label>
            <input type="number" value={gapX} min={5} onChange={(e)=>setGapX(Number(e.target.value)||5)}
              disabled={busy} style={{ marginLeft: 8, width: 120, padding: 6, borderRadius: 8, border: "1px solid #334", background: "#0b1220" }} />
          </div>
          <div>
            <label>Scale OCR</label>
            <input type="number" value={scale} min={1} step={0.5} onChange={(e)=>setScale(Number(e.target.value)||1)}
              disabled={busy} style={{ marginLeft: 8, width: 120, padding: 6, borderRadius: 8, border: "1px solid #334", background: "#0b1220" }} />
          </div>
        </div>
      </div>

      <div style={{ marginTop: 24 }}>
        <h3 style={{ fontWeight: 700, marginBottom: 8 }}>Log</h3>
        <pre style={{ background: "#0b1220", borderRadius: 12, padding: 16, border: "1px solid #223", whiteSpace: "pre-wrap", lineHeight: 1.4 }}>
{`${
log.length ? log.join('\n') : 'Belum ada proses.'
}`}
        </pre>
      </div>

      <canvas ref={canvasRef} style={{ display: "none" }} />
    </div>
  );
}
