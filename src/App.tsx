"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { PDFDocument, StandardFonts, rgb, type PDFPage } from "pdf-lib";
import * as XLSX from "xlsx";
import pdfWorkerUrl from "pdfjs-dist/legacy/build/pdf.worker.mjs?url";
import { ArrowDownToLine, CalendarDays, Check, FileSpreadsheet, FileText, LoaderCircle, LockKeyhole, PackageCheck, RefreshCcw, ShieldCheck, Sparkles, UploadCloud } from "lucide-react";

type Placement = "bottom" | "right";
type Marketplace = "meesho" | "flipkart";
type LabelRow = { originalIndex: number; fileIndex: number; pageNumber: number; size: string; sku: string };
type ToolRegistration = { name: string; title?: string; description: string; inputSchema: Record<string, unknown>; annotations?: { readOnlyHint?: boolean; untrustedContentHint?: boolean }; execute: (input: unknown) => unknown | Promise<unknown> };
declare global { interface Document { modelContext?: { registerTool: (tool: ToolRegistration, options?: { signal?: AbortSignal }) => void | Promise<void> } } }

const SIZE_ORDER = ["XXS", "XS", "S", "M", "L", "XL", "XXL", "3XL", "4XL", "5XL", "6XL", "7XL", "FREE SIZE", "UNKNOWN"];
const FLIPKART_PAGE_WIDTH = 216;
const FLIPKART_PAGE_HEIGHT = 360;
const FLIPKART_CROP = { left: 190 / 595, bottom: 460 / 842, right: 406 / 595, top: 820 / 842 };

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

function extractFlipkartProduct(text: string) {
  const labelRow = text.match(
    /SKU\s*ID\s*\|\s*Description\s+QTY\s+\d+\s+(FREE\s*SIZE|FREE|XXS|XS|S|M|L|XL|XXL|XXXL|XXXXL|XXXXXL|XXXXXXL|[2-7]XL)\s+(.+?)\s+\|\s+/i,
  );
  if (labelRow) {
    return {
      sku: `${labelRow[1]} ${labelRow[2]}`.replace(/\s+/g, " ").trim().toUpperCase(),
      size: cleanSize(labelRow[1]),
    };
  }
  return { sku: "UNKNOWN", size: "UNKNOWN" };
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
  const [marketplace, setMarketplace] = useState<Marketplace>("meesho");
  const [files, setFiles] = useState<File[]>([]);
  const [rows, setRows] = useState<LabelRow[]>([]);
  const [status, setStatus] = useState<"idle" | "reading" | "ready" | "exporting" | "error">("idle");
  const [message, setMessage] = useState("");
  const [placement, setPlacement] = useState<Placement>("bottom");
  const [dateText, setDateText] = useState(indiaDate);
  const inputRef = useRef<HTMLInputElement>(null);
  const processingRef = useRef(false);
  const isBusy = status === "reading" || status === "exporting";

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

  const analyseFiles = useCallback(async (selected: File[]) => {
    if (!selected.length || processingRef.current) return;
    const invalid = selected.find((file) => file.type !== "application/pdf" && !file.name.toLowerCase().endsWith(".pdf"));
    if (invalid) { setFiles([]); setRows([]); setStatus("error"); setMessage(`Please select only PDF files. “${invalid.name}” is not a PDF.`); return; }
    processingRef.current = true;
    setFiles(selected); setRows([]); setStatus("reading"); setMessage("Reading Size and SKU from every selected PDF…");
    let currentFile = selected[0].name;
    try {
      const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
      pdfjs.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;
      const parsed: LabelRow[] = [];
      const batchSize = 12;
      for (let fileIndex = 0; fileIndex < selected.length; fileIndex++) {
        currentFile = selected[fileIndex].name;
        setMessage(`Opening PDF ${fileIndex + 1} of ${selected.length}: ${currentFile}…`);
        const bytes = new Uint8Array(await selected[fileIndex].arrayBuffer());
        const loadingTask = pdfjs.getDocument({ data: bytes });
        try {
          const source = await loadingTask.promise;
          const offset = parsed.length;
          for (let start = 1; start <= source.numPages; start += batchSize) {
            const end = Math.min(source.numPages, start + batchSize - 1);
            setMessage(`PDF ${fileIndex + 1}/${selected.length} · Reading labels ${start}–${end} of ${source.numPages}…`);
            await Promise.all(Array.from({ length: end - start + 1 }, async (_, batchOffset) => {
              const pageNumber = start + batchOffset;
              const page = await source.getPage(pageNumber);
              const content = await page.getTextContent();
              const text = content.items.map((item) => ("str" in item ? item.str : "")).join(" ").replace(/\s+/g, " ").trim();
              const product = marketplace === "flipkart" ? extractFlipkartProduct(text) : extractProductDetails(text);
              const originalIndex = offset + pageNumber - 1;
              parsed[originalIndex] = { originalIndex, fileIndex, pageNumber, size: product.size, sku: product.sku };
              page.cleanup();
            }));
          }
        } finally {
          await loadingTask.destroy();
        }
      }
      setRows(parsed); setStatus("ready"); setMessage(`Ready — ${parsed.length} labels from ${selected.length} PDF${selected.length === 1 ? "" : "s"}.`);
    } catch (error) {
      console.error(error); setRows([]); setStatus("error");
      setMessage(`Could not read “${currentFile}”. No labels have been exported. Check that it is a valid, unlocked PDF, then replace the selection.`);
    } finally {
      processingRef.current = false;
    }
  }, [marketplace]);

  function chooseMarketplace(next: Marketplace) {
    if (next === marketplace || processingRef.current) return;
    setMarketplace(next); setFiles([]); setRows([]); setStatus("idle"); setMessage("");
    if (inputRef.current) inputRef.current.value = "";
  }

  const updateRow = (index: number, field: "size" | "sku", value: string) => setRows((current) => current.map((row) => row.originalIndex === index ? { ...row, [field]: field === "size" ? cleanSize(value) : value.toUpperCase().trim() || "UNKNOWN" } : row));

  async function exportPdf() {
    if (!files.length || !rows.length || processingRef.current) return;
    processingRef.current = true;
    setStatus("exporting"); setMessage("Creating one combined sorted PDF without changing label details…");
    try {
      const sources: PDFDocument[] = [];
      for (const file of files) sources.push(await PDFDocument.load(new Uint8Array(await file.arrayBuffer())));
      const output = await PDFDocument.create();
      if (marketplace === "flipkart") {
        for (const row of sortedRows) {
          const sourcePage = sources[row.fileIndex].getPage(row.pageNumber - 1);
          const { width, height } = sourcePage.getSize();
          const embedded = await output.embedPage(sourcePage, {
            left: width * FLIPKART_CROP.left,
            bottom: height * FLIPKART_CROP.bottom,
            right: width * FLIPKART_CROP.right,
            top: height * FLIPKART_CROP.top,
          });
          const page = output.addPage([FLIPKART_PAGE_WIDTH, FLIPKART_PAGE_HEIGHT]);
          page.drawPage(embedded, { x: 0, y: 0, width: FLIPKART_PAGE_WIDTH, height: FLIPKART_PAGE_HEIGHT });
        }
        saveBlob(await output.save(), `Flipkart_Labels_3x5_Size_SKU_Sorted_${dateText}.pdf`, "application/pdf");
      } else {
        const font = await output.embedFont(StandardFonts.HelveticaBold);
        const copied = new Array<PDFPage>(rows.length);
        for (let fileIndex = 0; fileIndex < sources.length; fileIndex++) {
          const fileRows = sortedRows.filter((row) => row.fileIndex === fileIndex);
          const pages = await output.copyPages(sources[fileIndex], fileRows.map((row) => row.pageNumber - 1));
          fileRows.forEach((row, index) => { copied[row.originalIndex] = pages[index] });
        }
        sortedRows.forEach((row) => {
          const page = copied[row.originalIndex];
          const { width, height } = page.getSize(); const fontSize = Math.max(14, Math.min(24, width * 0.037)); const textWidth = font.widthOfTextAtSize(dateText, fontSize);
          const x = placement === "right" ? width * 0.75 - textWidth / 2 : (width - textWidth) / 2; const y = placement === "right" ? height * 0.38 : Math.max(18, height * 0.055);
          page.drawText(dateText, { x: Math.max(10, Math.min(width - textWidth - 10, x)), y, size: fontSize, font, color: rgb(0.08, 0.08, 0.1) }); output.addPage(page);
        });
        saveBlob(await output.save(), `Meesho_Labels_Size_SKU_Sorted_${dateText}.pdf`, "application/pdf");
      }
      setStatus("ready"); setMessage(marketplace === "flipkart" ? "Combined sorted 3×5 PDF downloaded." : "Combined sorted PDF downloaded.");
    } catch (error) { console.error(error); setStatus("error"); setMessage("Could not create the combined PDF. Please check that every selected file is an original, unlocked label PDF."); }
    finally { processingRef.current = false; }
  }

  function exportExcel() {
    if (!rows.length || processingRef.current) return;
    const workbook = XLSX.utils.book_new();
    const details = skuGroups.map((group, index) => ({ "No.": index + 1, Size: group.size, SKU: group.sku, "Order Count": group.count }));
    const totals: Array<Record<string, string | number>> = sizeTotals.map((item, index) => ({ "No.": index + 1, Size: item.size, "Total Orders": item.count })); totals.push({ "No.": "", Size: "GRAND TOTAL", "Total Orders": rows.length });
    const detailSheet = XLSX.utils.json_to_sheet(details); const totalSheet = XLSX.utils.json_to_sheet(totals);
    detailSheet["!cols"] = [{ wch: 8 }, { wch: 14 }, { wch: 28 }, { wch: 16 }]; totalSheet["!cols"] = [{ wch: 8 }, { wch: 18 }, { wch: 18 }]; detailSheet["!autofilter"] = { ref: detailSheet["!ref"] ?? "A1:D1" };
    XLSX.utils.book_append_sheet(workbook, detailSheet, "Size-SKU Picklist"); XLSX.utils.book_append_sheet(workbook, totalSheet, "Size Totals"); XLSX.writeFile(workbook, `${marketplace === "flipkart" ? "Flipkart" : "Meesho"}_Size_Wise_Picklist_${dateText}.xlsx`); setMessage("Excel picklist downloaded.");
  }

  useEffect(() => {
    const context = document.modelContext; if (!context?.registerTool) return; const lifecycle = new AbortController();
    void Promise.resolve(context.registerTool({ name: "configure_meesho_label_export", title: "Configure label export", description: "Set the date and blank-space position for the visible Meesho label sorter.", inputSchema: { type: "object", properties: { dateText: { type: "string" }, placement: { type: "string", enum: ["bottom", "right"] } }, required: ["dateText", "placement"], additionalProperties: false }, annotations: { readOnlyHint: false, untrustedContentHint: false }, execute(input) { const value = input as { dateText?: unknown; placement?: unknown }; if (typeof value.dateText !== "string" || !/^\d{2}-\d{2}-\d{4}$/.test(value.dateText)) throw new Error("Use DD-MM-YYYY"); if (value.placement !== "bottom" && value.placement !== "right") throw new Error("Use bottom or right"); setDateText(value.dateText); setPlacement(value.placement); return { dateText: value.dateText, placement: value.placement } } }, { signal: lifecycle.signal })).catch(() => undefined);
    return () => lifecycle.abort();
  }, []);

  return <main className="min-h-screen bg-background text-foreground">
    <header className="border-b border-border/80 bg-white/90 backdrop-blur"><div className="mx-auto flex max-w-[1440px] items-center justify-between gap-4 px-5 py-4 lg:px-8">
      <div className="flex items-center gap-3"><div className="grid h-10 w-10 place-items-center rounded-xl bg-primary text-white"><PackageCheck className="h-5 w-5" /></div><div><p className="font-extrabold tracking-tight">Marketplace Label Sorter</p><p className="text-xs font-medium text-muted-foreground">Meesho + Flipkart · Size → SKU</p></div></div>
      <div className="flex items-center gap-2 rounded-full border border-emerald-200 bg-emerald-50 px-3 py-2 text-xs font-bold text-emerald-800"><LockKeyhole className="h-3.5 w-3.5" />Files stay in your browser</div>
    </div></header>
    <section className="mx-auto max-w-[1440px] px-5 py-7 lg:px-8 lg:py-10">
      <div className="mb-7 flex flex-col justify-between gap-5 lg:flex-row lg:items-end"><div><div className="mb-3 inline-flex items-center gap-2 rounded-full bg-fuchsia-50 px-3 py-1.5 text-xs font-extrabold uppercase tracking-[.13em] text-fuchsia-700"><Sparkles className="h-3.5 w-3.5" />Daily packing workflow</div><h1 className="max-w-3xl text-3xl font-black tracking-[-.04em] sm:text-4xl">Sort shipping labels by Size and SKU.</h1><p className="mt-3 max-w-2xl text-base leading-7 text-muted-foreground">Choose Meesho or Flipkart, upload one or more original PDFs, then download one combined sorted PDF and a size-wise Excel picklist.</p></div><div className="flex items-center gap-3 rounded-2xl border border-border bg-card px-4 py-3 shadow-sm"><CalendarDays className="h-5 w-5 text-primary" /><div><p className="text-xs font-bold uppercase tracking-wider text-muted-foreground">{marketplace === "flipkart" ? "Flipkart output" : "Current date"}</p><p className="text-lg font-black tabular-nums">{marketplace === "flipkart" ? "3 × 5 inches" : indiaDate()}</p></div></div></div>
      <div className="grid gap-6 xl:grid-cols-[minmax(0,1.5fr)_minmax(340px,.7fr)]">
        <section className="overflow-hidden rounded-[24px] border border-border bg-card shadow-[0_18px_55px_rgba(43,28,48,.08)]">
          <div className="flex items-center justify-between gap-4 bg-[#241826] px-5 py-4 text-white sm:px-6"><div><p className="text-xs font-bold uppercase tracking-[.14em] text-fuchsia-200">Step 1</p><h2 className="mt-1 text-xl font-extrabold">Upload {marketplace === "flipkart" ? "Flipkart" : "Meesho"} label PDFs</h2></div>{!!files.length && <button disabled={isBusy} onClick={() => inputRef.current?.click()} className="inline-flex shrink-0 items-center gap-2 rounded-xl border border-white/20 px-3 py-2 text-sm font-bold hover:bg-white/10 disabled:opacity-40"><RefreshCcw className="h-4 w-4" />Replace PDFs</button>}</div>
          <div className="p-5 sm:p-6" onDragOver={(e) => e.preventDefault()} onDrop={(e) => { e.preventDefault(); void analyseFiles(Array.from(e.dataTransfer.files)) }}>
            <div className="mb-5 grid grid-cols-2 gap-2 rounded-2xl bg-muted p-1.5"><MarketplaceButton disabled={isBusy} active={marketplace === "meesho"} label="Meesho" note="Original labels + date" onClick={() => chooseMarketplace("meesho")} /><MarketplaceButton disabled={isBusy} active={marketplace === "flipkart"} label="Flipkart" note="Crop to 3 × 5 inches" onClick={() => chooseMarketplace("flipkart")} /></div>
            <input ref={inputRef} type="file" accept="application/pdf,.pdf" multiple disabled={isBusy} aria-label="Select label PDFs" className="sr-only" onChange={(e) => { const selected = Array.from(e.currentTarget.files ?? []); e.currentTarget.value = ""; void analyseFiles(selected) }} />
            {!files.length ? <button disabled={isBusy} onClick={() => inputRef.current?.click()} className="group grid min-h-[310px] w-full place-items-center rounded-[20px] border-2 border-dashed border-fuchsia-200 bg-[linear-gradient(135deg,#fff8fd_0%,#fff_55%,#fff8ef_100%)] p-8 text-center hover:border-fuchsia-400 focus:outline-none focus:ring-4 focus:ring-fuchsia-100"><span><span className="mx-auto grid h-16 w-16 place-items-center rounded-2xl bg-primary text-white shadow-lg shadow-fuchsia-200 transition group-hover:-translate-y-1"><UploadCloud className="h-7 w-7" /></span><span className="mt-5 block text-xl font-black">Drop PDFs here or click to upload</span><span className="mt-2 block text-sm font-medium text-muted-foreground">Select multiple original {marketplace === "flipkart" ? "Flipkart A4 invoice + label" : "Meesho shipping-label"} PDFs together</span><span className="mt-3 block text-xs font-medium text-muted-foreground">Hold Ctrl / Cmd or Shift to select several files. Use one marketplace per batch.</span></span></button> : <div>
              <div className="rounded-2xl border border-border bg-muted/40 p-4">
                <div className="flex min-w-0 items-center gap-3"><div className="grid h-11 w-11 shrink-0 place-items-center rounded-xl bg-white text-primary shadow-sm"><FileText className="h-5 w-5" /></div><div className="min-w-0"><p className="font-extrabold">{files.length} PDF{files.length === 1 ? "" : "s"} selected</p><p className="text-sm text-muted-foreground">{(files.reduce((total, file) => total + file.size, 0) / 1048576).toFixed(2)} MB total · All labels sorted together</p></div></div>
                <ul className="mt-3 max-h-40 space-y-2 overflow-auto border-t border-border pt-3">{files.map((file, index) => <li key={index} className="flex min-w-0 items-center justify-between gap-3 text-sm"><span className="truncate font-semibold" title={file.name}>{index + 1}. {file.name}</span><span className="shrink-0 text-muted-foreground">{(file.size / 1048576).toFixed(2)} MB</span></li>)}</ul>
              </div>
              {!!message && <p role={status === "error" ? "alert" : "status"} className={`mt-3 flex items-start gap-2 break-words text-sm font-bold ${status === "error" ? "text-red-700" : "text-fuchsia-700"}`}>{isBusy ? <LoaderCircle className="mt-0.5 h-4 w-4 shrink-0 animate-spin" /> : status === "ready" ? <Check className="mt-0.5 h-4 w-4 shrink-0" /> : null}{message}</p>}
              {!!rows.length && <><div className="mt-5 grid grid-cols-2 gap-3 sm:grid-cols-4">{[["Labels", rows.length], ["Sizes", sizeTotals.length], ["Size–SKU groups", skuGroups.length], ["Check needed", unknownCount]].map(([label, value]) => <div key={label} className="rounded-2xl border border-border p-4"><p className="text-2xl font-black tabular-nums">{value}</p><p className="mt-1 text-xs font-bold uppercase tracking-wide text-muted-foreground">{label}</p></div>)}</div>
                <div className="mt-5 overflow-hidden rounded-2xl border border-border"><div className="flex items-center justify-between border-b border-border bg-muted/55 px-4 py-3"><div><h3 className="font-extrabold">Detected labels</h3><p className="text-xs text-muted-foreground">Correct any UNKNOWN value before export.</p></div><span className="rounded-lg bg-white px-2.5 py-1 text-xs font-bold shadow-sm">Sorted preview</span></div><div className="max-h-[330px] overflow-auto"><table className="w-full min-w-[720px] text-left text-sm"><thead className="sticky top-0 bg-white text-xs uppercase tracking-wider text-muted-foreground shadow-sm"><tr><th className="px-4 py-3">Order</th><th className="px-4 py-3">Source PDF</th><th className="px-4 py-3">Original page</th><th className="px-4 py-3">Size</th><th className="px-4 py-3">SKU</th></tr></thead><tbody className="divide-y divide-border">{sortedRows.map((row, index) => <tr key={row.originalIndex} className={row.size === "UNKNOWN" || row.sku === "UNKNOWN" ? "bg-amber-50" : "bg-white"}><td className="px-4 py-2.5 font-bold">{index + 1}</td><td className="px-4 py-2.5 text-muted-foreground"><span className="block max-w-40 truncate" title={files[row.fileIndex].name}>{row.fileIndex + 1}. {files[row.fileIndex].name}</span></td><td className="px-4 py-2.5 text-muted-foreground">{row.pageNumber}</td><td className="px-4 py-2.5"><select disabled={isBusy} aria-label={`Size for PDF ${row.fileIndex + 1} page ${row.pageNumber}`} value={row.size} onChange={(e) => updateRow(row.originalIndex, "size", e.target.value)} className="h-9 rounded-lg border border-border bg-white px-2.5 font-bold">{SIZE_ORDER.map((size) => <option key={size}>{size}</option>)}</select></td><td className="px-4 py-2.5"><input disabled={isBusy} aria-label={`SKU for PDF ${row.fileIndex + 1} page ${row.pageNumber}`} value={row.sku} onChange={(e) => updateRow(row.originalIndex, "sku", e.target.value)} className="h-9 w-full rounded-lg border border-border bg-white px-2.5 font-semibold" /></td></tr>)}</tbody></table></div></div>
              </>}</div>}
            {!files.length && !!message && <p role="alert" className="mt-3 break-words text-sm font-bold text-red-700">{message}</p>}
          </div>
        </section>
        <aside className="space-y-5">
          {marketplace === "meesho" ? <section className="rounded-[24px] border border-border bg-card p-5 shadow-[0_18px_55px_rgba(43,28,48,.07)] sm:p-6"><div className="flex items-center justify-between"><div><p className="text-xs font-bold uppercase tracking-[.14em] text-fuchsia-700">Step 2</p><h2 className="mt-1 text-xl font-extrabold">Date position</h2></div><CalendarDays className="h-6 w-6 text-primary" /></div><label className="mt-5 block text-sm font-bold" htmlFor="dateText">Date shown on every label</label><input id="dateText" value={dateText} onChange={(e) => setDateText(e.target.value)} className="mt-2 h-12 w-full rounded-xl border border-border bg-white px-4 text-lg font-black tabular-nums focus:ring-4 focus:ring-fuchsia-100" />
            <div className="mt-4 grid gap-2"><PositionButton active={placement === "bottom"} title="Bottom blank space" note="Centered below label details" onClick={() => setPlacement("bottom")} /><PositionButton active={placement === "right"} title="Right blank space" note="Centered next to product details" onClick={() => setPlacement("right")} /></div>
            <div className="mt-5 rounded-2xl bg-[#f8f5f8] p-4"><div className="relative mx-auto aspect-[3/2] max-w-[280px] overflow-hidden rounded-xl border border-zinc-300 bg-white shadow-sm"><div className="absolute left-4 top-4 h-4 w-2/5 rounded bg-zinc-900" /><div className="absolute left-4 top-11 space-y-2"><div className="h-2 w-24 rounded bg-zinc-300" /><div className="h-2 w-20 rounded bg-zinc-200" /><div className="h-2 w-28 rounded bg-zinc-200" /></div><div className="absolute right-4 top-4 grid h-14 w-14 grid-cols-4 gap-0.5 bg-zinc-900 p-1 opacity-85">{Array.from({ length: 16 }).map((_, i) => <span key={i} className={i % 3 === 0 ? "bg-white" : "bg-zinc-900"} />)}</div><div className={`absolute font-black tabular-nums ${placement === "bottom" ? "bottom-3 left-1/2 -translate-x-1/2" : "right-4 top-[58%] -translate-y-1/2"}`}>{dateText}</div></div><p className="mt-3 flex items-center justify-center gap-2 text-xs font-bold text-emerald-700"><ShieldCheck className="h-3.5 w-3.5" />Original label stays unchanged</p></div>
          </section> : <section className="rounded-[24px] border border-border bg-card p-5 shadow-[0_18px_55px_rgba(43,28,48,.07)] sm:p-6"><div className="flex items-center justify-between"><div><p className="text-xs font-bold uppercase tracking-[.14em] text-fuchsia-700">Step 2</p><h2 className="mt-1 text-xl font-extrabold">3 × 5 label crop</h2></div><FileText className="h-6 w-6 text-primary" /></div><div className="mt-5 rounded-2xl border border-emerald-200 bg-emerald-50 p-4"><p className="flex items-center gap-2 font-extrabold text-emerald-900"><Check className="h-4 w-4" />Flipkart label area detected</p><p className="mt-2 text-sm leading-6 text-emerald-800">The shipping label is cropped from the top of each A4 invoice page and exported at exactly 3 × 5 inches. The invoice section is removed.</p></div><div className="mt-4 grid grid-cols-2 gap-3"><div className="rounded-xl border border-border p-3 text-center"><p className="text-xl font-black">3 in</p><p className="text-xs font-bold uppercase text-muted-foreground">Width</p></div><div className="rounded-xl border border-border p-3 text-center"><p className="text-xl font-black">5 in</p><p className="text-xs font-bold uppercase text-muted-foreground">Height</p></div></div></section>}
          <section className="rounded-[24px] bg-[#241826] p-5 text-white shadow-[0_18px_55px_rgba(43,28,48,.18)] sm:p-6"><p className="text-xs font-bold uppercase tracking-[.14em] text-fuchsia-200">Step 3</p><h2 className="mt-1 text-xl font-extrabold">Download files</h2><p className="mt-2 text-sm leading-6 text-zinc-300">Sort order: XXS → XS → S → M → L → XL → XXL → 3XL → 4XL → 5XL, then SKU.</p><div className="mt-5 grid gap-3"><button disabled={!rows.length || isBusy || !/^\d{2}-\d{2}-\d{4}$/.test(dateText)} onClick={() => void exportPdf()} className="inline-flex h-12 items-center justify-center gap-2 rounded-xl bg-fuchsia-500 px-4 font-extrabold hover:bg-fuchsia-400 disabled:opacity-40">{status === "exporting" ? <LoaderCircle className="h-4 w-4 animate-spin" /> : <ArrowDownToLine className="h-4 w-4" />}Download {marketplace === "flipkart" ? "3×5 " : ""}combined sorted PDF</button><button disabled={!rows.length || isBusy} onClick={exportExcel} className="inline-flex h-12 items-center justify-center gap-2 rounded-xl border border-white/20 bg-white/10 px-4 font-extrabold hover:bg-white/15 disabled:opacity-40"><FileSpreadsheet className="h-4 w-4" />Download Excel picklist</button></div><div className="mt-5 grid grid-cols-3 divide-x divide-white/10 rounded-xl border border-white/10 bg-black/10 py-3 text-center text-xs font-bold text-zinc-300"><span>{marketplace === "flipkart" ? "3×5 crop" : "PDF preserved"}</span><span>Size totals</span><span>SKU count</span></div></section>
        </aside>
      </div>
    </section>
  </main>;
}

function PositionButton({ active, title, note, onClick }: { active: boolean; title: string; note: string; onClick: () => void }) {
  return <button type="button" onClick={onClick} className={`flex items-center justify-between rounded-xl border p-3 text-left transition ${active ? "border-fuchsia-400 bg-fuchsia-50 ring-2 ring-fuchsia-100" : "border-border hover:bg-muted/50"}`}><span><span className="block text-sm font-extrabold">{title}</span><span className="text-xs text-muted-foreground">{note}</span></span>{active && <Check className="h-4 w-4 text-fuchsia-700" />}</button>;
}

function MarketplaceButton({ active, disabled, label, note, onClick }: { active: boolean; disabled: boolean; label: string; note: string; onClick: () => void }) {
  return <button type="button" disabled={disabled} onClick={onClick} className={`rounded-xl px-3 py-3 text-left transition disabled:opacity-50 ${active ? "bg-white shadow-sm ring-1 ring-border" : "text-muted-foreground hover:bg-white/60"}`}><span className="block text-sm font-extrabold">{label}</span><span className="text-xs font-medium">{note}</span></button>;
}
