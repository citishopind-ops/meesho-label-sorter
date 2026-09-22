"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";
import * as XLSX from "xlsx";
import pdfWorkerUrl from "pdfjs-dist/legacy/build/pdf.worker.mjs?url";
import { ArrowDownToLine, CalendarDays, Check, FileSpreadsheet, FileText, LoaderCircle, LockKeyhole, PackageCheck, RefreshCcw, ShieldCheck, Sparkles, UploadCloud } from "lucide-react";

type Placement = "bottom" | "right";
type LabelRow = { originalIndex: number; pageNumber: number; size: string; sku: string };
type ToolRegistration = { name: string; title?: string; description: string; inputSchema: Record<string, unknown>; annotations?: { readOnlyHint?: boolean; untrustedContentHint?: boolean }; execute: (input: unknown) => unknown | Promise<unknown> };
declare global { interface Document { modelContext?: { registerTool: (tool: ToolRegistration, options?: { signal?: AbortSignal }) => void | Promise<void> } } }

const SIZE_ORDER = ["XXS", "XS", "S", "M", "L", "XL", "XXL", "3XL", "4XL", "5XL", "6XL", "7XL", "FREE SIZE", "UNKNOWN"];

function cleanSize(value?: string) {
  if (!value) return "UNKNOWN";
  const size = value.toUpperCase().replace(/[()\[\],]/g, " ").replace(/\s+/g, " ").trim();
  if (/FREE\s*SIZE|^FREE$/.test(size)) return "FREE SIZE";
  if (/^(?:XXXL|3XL)$/.test(size)) return "3XL";
  if (/^(?:XXXXL|4XL)$/.test(size)) return "4XL";
  if (/^(?:XXXXXL|5XL)$/.test(size)) return "5XL";
  if (/^(?:XXXXXXL|6XL)$/.test(size)) return "6XL";
  if (/^(?:XXL|2XL)$/.test(size)) return "XXL";
  return SIZE_ORDER.includes(size) ? size : "UNKNOWN";
}

function extractSize(text: string) {
  const patterns = [
    /(?:product\s*)?size\s*[:\-]?\s*(FREE\s*SIZE|FREE|XXS|XS|S|M|L|XL|XXL|XXXL|[2-7]XL)\b/i,
    /\b(FREE\s*SIZE|XXS|XS|XXXL|XXL|XL|[2-7]XL)\b/i,
  ];
  for (const pattern of patterns) { const match = text.match(pattern); if (match) return cleanSize(match[1]); }
  return "UNKNOWN";
}

function extractSku(text: string) {
  const patterns = [
    /(?:seller\s*)?sku(?:\s*id)?\s*[:#\-]?\s*([A-Z0-9][A-Z0-9._\-/]{1,50})/i,
    /style\s*(?:id|code)?\s*[:#\-]?\s*([A-Z0-9][A-Z0-9._\-/]{1,50})/i,
  ];
  for (const pattern of patterns) { const match = text.match(pattern); if (match) return match[1].trim().toUpperCase(); }
  return "UNKNOWN";
}

function extractProductDetails(text: string) {
  const productRow = text.match(
    /Product\s+Details\s+SKU\s+Size\s+Qty\s+Color\s+Order\s+No\.?\s+(.+?)\s+(FREE\s*SIZE|FREE|XXS|XS|S|M|L|XL|XXL|XXXL|XXXXL|XXXXXL|XXXXXXL|[2-7]XL)\s+\d+\s+/i,
  );
  if (productRow) {
    return {
      sku: productRow[1].replace(/\s+/g, " ").trim().toUpperCase(),
      size: cleanSize(productRow[2]),
    };
  }
  return { sku: extractSku(text), size: extractSize(text) };
}

const sizeRank = (size: string) => { const index = SIZE_ORDER.indexOf(size); return index < 0 ? SIZE_ORDER.length : index };
function indiaDate() {
  const parts = new Intl.DateTimeFormat("en-GB", { timeZone: "Asia/Kolkata", day: "2-digit", month: "2-digit", year: "numeric" }).formatToParts(new Date());
  const get = (type: string) => parts.find((part) => part.type === type)?.value;
  return `${get("day")}-${get("month")}-${get("year")}`;
}
function saveBlob(data: Uint8Array, filename: string, type: string) {
  const buffer = new ArrayBuffer(data.byteLength); new Uint8Array(buffer).set(data);
  const url = URL.createObjectURL(new Blob([buffer], { type }));
  const link = document.createElement("a"); link.href = url; link.download = filename; document.body.appendChild(link); link.click(); link.remove(); URL.revokeObjectURL(url);
}

export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [rows, setRows] = useState<LabelRow[]>([]);
  const [status, setStatus] = useState<"idle" | "reading" | "ready" | "exporting" | "error">("idle");
  const [message, setMessage] = useState("");
  const [placement, setPlacement] = useState<Placement>("bottom");
  const [dateText, setDateText] = useState(indiaDate);
  const inputRef = useRef<HTMLInputElement>(null);

  const sortedRows = useMemo(() => [...rows].sort((a, b) => sizeRank(a.size) - sizeRank(b.size) || a.sku.localeCompare(b.sku, undefined, { numeric: true }) || a.originalIndex - b.originalIndex), [rows]);
  const sizeTotals = useMemo(() => {
    const totals = new Map<string, number>(); rows.forEach((row) => totals.set(row.size, (totals.get(row.size) ?? 0) + 1));
    return [...totals.entries()].sort(([a], [b]) => sizeRank(a) - sizeRank(b)).map(([size, count]) => ({ size, count }));
  }, [rows]);
  const skuGroups = useMemo(() => {
    const groups = new Map<string, { size: string; sku: string; count: number }>();
    sortedRows.forEach((row) => { const key = `${row.size}|||${row.sku}`; const current = groups.get(key); groups.set(key, { size: row.size, sku: row.sku, count: (current?.count ?? 0) + 1 }) });
    return [...groups.values()];
  }, [sortedRows]);
  const unknownCount = rows.filter((row) => row.size === "UNKNOWN" || row.sku === "UNKNOWN").length;

  const analyseFile = useCallback(async (selected: File) => {
    if (selected.type !== "application/pdf" && !selected.name.toLowerCase().endsWith(".pdf")) { setStatus("error"); setMessage("Please upload a PDF file."); return; }
    setFile(selected); setRows([]); setStatus("reading"); setMessage("Reading Size and SKU from each label…");
    try {
      const bytes = new Uint8Array(await selected.arrayBuffer());
      const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
      pdfjs.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;
      const source = await pdfjs.getDocument({ data: bytes }).promise;
      const parsed = new Array<LabelRow>(source.numPages);
      const batchSize = 12;
      for (let start = 1; start <= source.numPages; start += batchSize) {
        const end = Math.min(source.numPages, start + batchSize - 1);
        setMessage(`Reading labels ${start}–${end} of ${source.numPages}…`);
        await Promise.all(Array.from({ length: end - start + 1 }, async (_, offset) => {
          const pageNumber = start + offset;
          const page = await source.getPage(pageNumber);
          const content = await page.getTextContent();
          const text = content.items.map((item) => ("str" in item ? item.str : "")).join(" ").replace(/\s+/g, " ").trim();
          const product = extractProductDetails(text);
          parsed[pageNumber - 1] = { originalIndex: pageNumber - 1, pageNumber, size: product.size, sku: product.sku };
          page.cleanup();
        }));
      }
      await source.cleanup();
      setRows(parsed); setStatus("ready"); setMessage(`Ready — ${parsed.length} labels found.`);
    } catch (error) { console.error(error); setStatus("error"); setMessage("The PDF reader could not start. Please try again after refreshing the page."); }
  }, []);

  const updateRow = (index: number, field: "size" | "sku", value: string) => setRows((current) => current.map((row) => row.originalIndex === index ? { ...row, [field]: field === "size" ? cleanSize(value) : value.toUpperCase().trim() || "UNKNOWN" } : row));

  async function exportPdf() {
    if (!file || !rows.length) return;
    setStatus("exporting"); setMessage("Creating sorted PDF without changing label details…");
    try {
      const source = await PDFDocument.load(new Uint8Array(await file.arrayBuffer()));
      const output = await PDFDocument.create(); const font = await output.embedFont(StandardFonts.HelveticaBold);
      const copied = await output.copyPages(source, sortedRows.map((row) => row.originalIndex));
      copied.forEach((page) => {
        const { width, height } = page.getSize(); const fontSize = Math.max(14, Math.min(24, width * 0.037)); const textWidth = font.widthOfTextAtSize(dateText, fontSize);
        const x = placement === "right" ? width * 0.75 - textWidth / 2 : (width - textWidth) / 2; const y = placement === "right" ? height * 0.38 : Math.max(18, height * 0.055);
        page.drawText(dateText, { x: Math.max(10, Math.min(width - textWidth - 10, x)), y, size: fontSize, font, color: rgb(0.08, 0.08, 0.1) }); output.addPage(page);
      });
      saveBlob(await output.save(), `Meesho_Labels_Size_SKU_Sorted_${dateText}.pdf`, "application/pdf"); setStatus("ready"); setMessage("Sorted PDF downloaded.");
    } catch (error) { console.error(error); setStatus("error"); setMessage("Could not create the PDF. Please try the original downloaded label file."); }
  }

  function exportExcel() {
    if (!rows.length) return;
    const workbook = XLSX.utils.book_new();
    const details = skuGroups.map((group, index) => ({ "No.": index + 1, Size: group.size, SKU: group.sku, "Order Count": group.count }));
    const totals: Array<Record<string, string | number>> = sizeTotals.map((item, index) => ({ "No.": index + 1, Size: item.size, "Total Orders": item.count })); totals.push({ "No.": "", Size: "GRAND TOTAL", "Total Orders": rows.length });
    const detailSheet = XLSX.utils.json_to_sheet(details); const totalSheet = XLSX.utils.json_to_sheet(totals);
    detailSheet["!cols"] = [{ wch: 8 }, { wch: 14 }, { wch: 28 }, { wch: 16 }]; totalSheet["!cols"] = [{ wch: 8 }, { wch: 18 }, { wch: 18 }]; detailSheet["!autofilter"] = { ref: detailSheet["!ref"] ?? "A1:D1" };
    XLSX.utils.book_append_sheet(workbook, detailSheet, "Size-SKU Picklist"); XLSX.utils.book_append_sheet(workbook, totalSheet, "Size Totals"); XLSX.writeFile(workbook, `Meesho_Size_Wise_Picklist_${dateText}.xlsx`); setMessage("Excel picklist downloaded.");
  }

  useEffect(() => {
    const context = document.modelContext; if (!context?.registerTool) return; const lifecycle = new AbortController();
    void Promise.resolve(context.registerTool({ name: "configure_meesho_label_export", title: "Configure label export", description: "Set the date and blank-space position for the visible Meesho label sorter.", inputSchema: { type: "object", properties: { dateText: { type: "string" }, placement: { type: "string", enum: ["bottom", "right"] } }, required: ["dateText", "placement"], additionalProperties: false }, annotations: { readOnlyHint: false, untrustedContentHint: false }, execute(input) { const value = input as { dateText?: unknown; placement?: unknown }; if (typeof value.dateText !== "string" || !/^\d{2}-\d{2}-\d{4}$/.test(value.dateText)) throw new Error("Use DD-MM-YYYY"); if (value.placement !== "bottom" && value.placement !== "right") throw new Error("Use bottom or right"); setDateText(value.dateText); setPlacement(value.placement); return { dateText: value.dateText, placement: value.placement } } }, { signal: lifecycle.signal })).catch(() => undefined);
    return () => lifecycle.abort();
  }, []);

  return <main className="min-h-screen bg-background text-foreground">
    <header className="border-b border-border/80 bg-white/90 backdrop-blur"><div className="mx-auto flex max-w-[1440px] items-center justify-between gap-4 px-5 py-4 lg:px-8">
      <div className="flex items-center gap-3"><div className="grid h-10 w-10 place-items-center rounded-xl bg-primary text-white"><PackageCheck className="h-5 w-5" /></div><div><p className="font-extrabold tracking-tight">Meesho Label Sorter</p><p className="text-xs font-medium text-muted-foreground">Size → SKU → Picklist</p></div></div>
      <div className="flex items-center gap-2 rounded-full border border-emerald-200 bg-emerald-50 px-3 py-2 text-xs font-bold text-emerald-800"><LockKeyhole className="h-3.5 w-3.5" />Files stay in your browser</div>
    </div></header>
    <section className="mx-auto max-w-[1440px] px-5 py-7 lg:px-8 lg:py-10">
      <div className="mb-7 flex flex-col justify-between gap-5 lg:flex-row lg:items-end"><div><div className="mb-3 inline-flex items-center gap-2 rounded-full bg-fuchsia-50 px-3 py-1.5 text-xs font-extrabold uppercase tracking-[.13em] text-fuchsia-700"><Sparkles className="h-3.5 w-3.5" />Daily packing workflow</div><h1 className="max-w-3xl text-3xl font-black tracking-[-.04em] sm:text-4xl">Sort Meesho labels. Keep every detail unchanged.</h1><p className="mt-3 max-w-2xl text-base leading-7 text-muted-foreground">Upload one label PDF. Get a Size–SKU sorted PDF with today’s date in the blank space, plus an Excel picklist with order counts.</p></div><div className="flex items-center gap-3 rounded-2xl border border-border bg-card px-4 py-3 shadow-sm"><CalendarDays className="h-5 w-5 text-primary" /><div><p className="text-xs font-bold uppercase tracking-wider text-muted-foreground">Current date</p><p className="text-lg font-black tabular-nums">{indiaDate()}</p></div></div></div>
      <div className="grid gap-6 xl:grid-cols-[minmax(0,1.5fr)_minmax(340px,.7fr)]">
        <section className="overflow-hidden rounded-[24px] border border-border bg-card shadow-[0_18px_55px_rgba(43,28,48,.08)]">
          <div className="flex items-center justify-between gap-4 bg-[#241826] px-5 py-4 text-white sm:px-6"><div><p className="text-xs font-bold uppercase tracking-[.14em] text-fuchsia-200">Step 1</p><h2 className="mt-1 text-xl font-extrabold">Upload Meesho label PDF</h2></div>{file && <button onClick={() => inputRef.current?.click()} className="inline-flex items-center gap-2 rounded-xl border border-white/20 px-3 py-2 text-sm font-bold hover:bg-white/10"><RefreshCcw className="h-4 w-4" />Replace</button>}</div>
          <div className="p-5 sm:p-6"><input ref={inputRef} type="file" accept="application/pdf,.pdf" className="sr-only" onChange={(e) => { const selected = e.target.files?.[0]; if (selected) void analyseFile(selected) }} />
            {!file ? <button onClick={() => inputRef.current?.click()} onDragOver={(e) => e.preventDefault()} onDrop={(e) => { e.preventDefault(); const selected = e.dataTransfer.files?.[0]; if (selected) void analyseFile(selected) }} className="group grid min-h-[310px] w-full place-items-center rounded-[20px] border-2 border-dashed border-fuchsia-200 bg-[linear-gradient(135deg,#fff8fd_0%,#fff_55%,#fff8ef_100%)] p-8 text-center hover:border-fuchsia-400 focus:outline-none focus:ring-4 focus:ring-fuchsia-100"><span><span className="mx-auto grid h-16 w-16 place-items-center rounded-2xl bg-primary text-white shadow-lg shadow-fuchsia-200 transition group-hover:-translate-y-1"><UploadCloud className="h-7 w-7" /></span><span className="mt-5 block text-xl font-black">Drop PDF here or click to upload</span><span className="mt-2 block text-sm font-medium text-muted-foreground">Original Meesho shipping-label PDF</span></span></button> : <div>
              <div className="flex flex-col gap-4 rounded-2xl border border-border bg-muted/40 p-4 sm:flex-row sm:items-center sm:justify-between"><div className="flex min-w-0 items-center gap-3"><div className="grid h-11 w-11 shrink-0 place-items-center rounded-xl bg-white text-primary shadow-sm"><FileText className="h-5 w-5" /></div><div className="min-w-0"><p className="truncate font-extrabold">{file.name}</p><p className="text-sm text-muted-foreground">{(file.size / 1048576).toFixed(2)} MB</p></div></div><div className="flex items-center gap-2 text-sm font-bold text-fuchsia-700">{(status === "reading" || status === "exporting") ? <LoaderCircle className="h-4 w-4 animate-spin" /> : status === "ready" ? <Check className="h-4 w-4" /> : null}{message}</div></div>
              {!!rows.length && <><div className="mt-5 grid grid-cols-2 gap-3 sm:grid-cols-4">{[["Labels", rows.length], ["Sizes", sizeTotals.length], ["Size–SKU groups", skuGroups.length], ["Check needed", unknownCount]].map(([label, value]) => <div key={label} className="rounded-2xl border border-border p-4"><p className="text-2xl font-black tabular-nums">{value}</p><p className="mt-1 text-xs font-bold uppercase tracking-wide text-muted-foreground">{label}</p></div>)}</div>
                <div className="mt-5 overflow-hidden rounded-2xl border border-border"><div className="flex items-center justify-between border-b border-border bg-muted/55 px-4 py-3"><div><h3 className="font-extrabold">Detected labels</h3><p className="text-xs text-muted-foreground">Correct any UNKNOWN value before export.</p></div><span className="rounded-lg bg-white px-2.5 py-1 text-xs font-bold shadow-sm">Sorted preview</span></div><div className="max-h-[330px] overflow-auto"><table className="w-full min-w-[560px] text-left text-sm"><thead className="sticky top-0 bg-white text-xs uppercase tracking-wider text-muted-foreground shadow-sm"><tr><th className="px-4 py-3">Order</th><th className="px-4 py-3">Original page</th><th className="px-4 py-3">Size</th><th className="px-4 py-3">SKU</th></tr></thead><tbody className="divide-y divide-border">{sortedRows.map((row, index) => <tr key={row.originalIndex} className={row.size === "UNKNOWN" || row.sku === "UNKNOWN" ? "bg-amber-50" : "bg-white"}><td className="px-4 py-2.5 font-bold">{index + 1}</td><td className="px-4 py-2.5 text-muted-foreground">{row.pageNumber}</td><td className="px-4 py-2.5"><select aria-label={`Size for page ${row.pageNumber}`} value={row.size} onChange={(e) => updateRow(row.originalIndex, "size", e.target.value)} className="h-9 rounded-lg border border-border bg-white px-2.5 font-bold">{SIZE_ORDER.map((size) => <option key={size}>{size}</option>)}</select></td><td className="px-4 py-2.5"><input aria-label={`SKU for page ${row.pageNumber}`} value={row.sku} onChange={(e) => updateRow(row.originalIndex, "sku", e.target.value)} className="h-9 w-full rounded-lg border border-border bg-white px-2.5 font-semibold" /></td></tr>)}</tbody></table></div></div>
              </>}</div>}
          </div>
        </section>
        <aside className="space-y-5">
          <section className="rounded-[24px] border border-border bg-card p-5 shadow-[0_18px_55px_rgba(43,28,48,.07)] sm:p-6"><div className="flex items-center justify-between"><div><p className="text-xs font-bold uppercase tracking-[.14em] text-fuchsia-700">Step 2</p><h2 className="mt-1 text-xl font-extrabold">Date position</h2></div><CalendarDays className="h-6 w-6 text-primary" /></div><label className="mt-5 block text-sm font-bold" htmlFor="dateText">Date shown on every label</label><input id="dateText" value={dateText} onChange={(e) => setDateText(e.target.value)} className="mt-2 h-12 w-full rounded-xl border border-border bg-white px-4 text-lg font-black tabular-nums focus:ring-4 focus:ring-fuchsia-100" />
            <div className="mt-4 grid gap-2"><PositionButton active={placement === "bottom"} title="Bottom blank space" note="Centered below label details" onClick={() => setPlacement("bottom")} /><PositionButton active={placement === "right"} title="Right blank space" note="Centered next to product details" onClick={() => setPlacement("right")} /></div>
            <div className="mt-5 rounded-2xl bg-[#f8f5f8] p-4"><div className="relative mx-auto aspect-[3/2] max-w-[280px] overflow-hidden rounded-xl border border-zinc-300 bg-white shadow-sm"><div className="absolute left-4 top-4 h-4 w-2/5 rounded bg-zinc-900" /><div className="absolute left-4 top-11 space-y-2"><div className="h-2 w-24 rounded bg-zinc-300" /><div className="h-2 w-20 rounded bg-zinc-200" /><div className="h-2 w-28 rounded bg-zinc-200" /></div><div className="absolute right-4 top-4 grid h-14 w-14 grid-cols-4 gap-0.5 bg-zinc-900 p-1 opacity-85">{Array.from({ length: 16 }).map((_, i) => <span key={i} className={i % 3 === 0 ? "bg-white" : "bg-zinc-900"} />)}</div><div className={`absolute font-black tabular-nums ${placement === "bottom" ? "bottom-3 left-1/2 -translate-x-1/2" : "right-4 top-[58%] -translate-y-1/2"}`}>{dateText}</div></div><p className="mt-3 flex items-center justify-center gap-2 text-xs font-bold text-emerald-700"><ShieldCheck className="h-3.5 w-3.5" />Original label stays unchanged</p></div>
          </section>
          <section className="rounded-[24px] bg-[#241826] p-5 text-white shadow-[0_18px_55px_rgba(43,28,48,.18)] sm:p-6"><p className="text-xs font-bold uppercase tracking-[.14em] text-fuchsia-200">Step 3</p><h2 className="mt-1 text-xl font-extrabold">Download files</h2><p className="mt-2 text-sm leading-6 text-zinc-300">Sort order: XXS → XS → S → M → L → XL → XXL → 3XL → 4XL → 5XL, then SKU.</p><div className="mt-5 grid gap-3"><button disabled={!rows.length || status === "exporting" || !/^\d{2}-\d{2}-\d{4}$/.test(dateText)} onClick={() => void exportPdf()} className="inline-flex h-12 items-center justify-center gap-2 rounded-xl bg-fuchsia-500 px-4 font-extrabold hover:bg-fuchsia-400 disabled:opacity-40">{status === "exporting" ? <LoaderCircle className="h-4 w-4 animate-spin" /> : <ArrowDownToLine className="h-4 w-4" />}Download sorted PDF</button><button disabled={!rows.length} onClick={exportExcel} className="inline-flex h-12 items-center justify-center gap-2 rounded-xl border border-white/20 bg-white/10 px-4 font-extrabold hover:bg-white/15 disabled:opacity-40"><FileSpreadsheet className="h-4 w-4" />Download Excel picklist</button></div><div className="mt-5 grid grid-cols-3 divide-x divide-white/10 rounded-xl border border-white/10 bg-black/10 py-3 text-center text-xs font-bold text-zinc-300"><span>PDF preserved</span><span>Size totals</span><span>SKU count</span></div></section>
        </aside>
      </div>
    </section>
  </main>;
}

function PositionButton({ active, title, note, onClick }: { active: boolean; title: string; note: string; onClick: () => void }) {
  return <button type="button" onClick={onClick} className={`flex items-center justify-between rounded-xl border p-3 text-left transition ${active ? "border-fuchsia-400 bg-fuchsia-50 ring-2 ring-fuchsia-100" : "border-border hover:bg-muted/50"}`}><span><span className="block text-sm font-extrabold">{title}</span><span className="text-xs text-muted-foreground">{note}</span></span>{active && <Check className="h-4 w-4 text-fuchsia-700" />}</button>;
}
