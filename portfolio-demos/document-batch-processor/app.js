const fileInput = document.getElementById("files");
const sampleBtn = document.getElementById("sampleBtn");
const exportBtn = document.getElementById("exportBtn");
const resultsBody = document.getElementById("results");
const summary = document.getElementById("summary");

const sampleFiles = [
  "INV_ORD-1001_2026-10-01.pdf",
  "PO_ORD-1001_2026-09-30.pdf",
  "DN_ORD-1001_2026-10-02.pdf",
  "INV_ORD-1002_2026-10-01.pdf",
  "INV_ORD-1002_2026-10-01-copy.pdf",
  "PO_ORD-1003_2026-09-29.pdf",
  "delivery_1004.pdf",
  "INV_ORD-1005_2026-10-03.pdf",
  "PO_ORD-1005_2026-10-02.pdf"
];

let currentRows = [];

fileInput.addEventListener("change", () => {
  const names = [...fileInput.files].map(file => file.name);
  processNames(names);
});

sampleBtn.addEventListener("click", () => processNames(sampleFiles));

exportBtn.addEventListener("click", () => {
  if (!currentRows.length) return;

  const headers = ["Original","Type","Reference","Date","ProposedName","Status","Notes"];
  const lines = [headers.join(",")];

  for (const row of currentRows) {
    lines.push([
      csv(row.original),
      csv(row.type),
      csv(row.reference),
      csv(row.date),
      csv(row.proposedName),
      csv(row.status),
      csv(row.notes)
    ].join(","));
  }

  const blob = new Blob([lines.join("\n")], {type:"text/csv;charset=utf-8"});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "document_batch_report.csv";
  a.click();
  URL.revokeObjectURL(url);
});

function processNames(names) {
  const parsed = names.map(parseName);
  applyDuplicateChecks(parsed);
  applyPairChecks(parsed);
  currentRows = parsed;
  render(parsed);
}

function parseName(name) {
  const match = name.match(/^(INV|PO|DN)_([A-Z]+-\d+)_(\d{4}-\d{2}-\d{2})(?:-copy)?\.([a-z0-9]+)$/i);

  if (!match) {
    return {
      original:name,
      type:"",
      reference:"",
      date:"",
      proposedName:"",
      status:"Invalid format",
      notes:"Expected TYPE_REFERENCE_YYYY-MM-DD.ext"
    };
  }

  const [, rawType, rawReference, date, ext] = match;
  const type = rawType.toUpperCase();
  const reference = rawReference.toUpperCase();

  return {
    original:name,
    type,
    reference,
    date,
    proposedName:`${type}_${reference}_${date}.${ext.toLowerCase()}`,
    status:"OK",
    notes:""
  };
}

function applyDuplicateChecks(rows) {
  const seen = new Map();

  for (const row of rows) {
    if (!row.type || !row.reference) continue;

    const key = `${row.type}|${row.reference}`;

    if (seen.has(key)) {
      row.status = "Duplicate";
      row.notes = `Duplicate of ${seen.get(key)}`;
    } else {
      seen.set(key, row.original);
    }
  }
}

function applyPairChecks(rows) {
  const byReference = new Map();

  for (const row of rows) {
    if (!row.reference || row.status === "Invalid format") continue;
    if (!byReference.has(row.reference)) byReference.set(row.reference, new Set());
    byReference.get(row.reference).add(row.type);
  }

  for (const row of rows) {
    if (!row.reference || row.status !== "OK") continue;

    const types = byReference.get(row.reference);
    const missing = ["INV","PO"].filter(required => !types.has(required));

    if (missing.length) {
      row.status = "Missing pair";
      row.notes = `Missing required type: ${missing.join(" + ")}`;
    }
  }
}

function render(rows) {
  resultsBody.innerHTML = "";

  for (const row of rows) {
    const tr = document.createElement("tr");

    [
      row.original,
      row.type,
      row.reference,
      row.date,
      row.proposedName,
      row.status,
      row.notes
    ].forEach((value, index) => {
      const td = document.createElement("td");
      td.textContent = value;
      if (index === 5) td.className = row.status === "OK" ? "ok" : "issue";
      tr.appendChild(td);
    });

    resultsBody.appendChild(tr);
  }

  const issues = rows.filter(row => row.status !== "OK").length;
  summary.textContent = `${rows.length} documents checked, ${issues} exception(s) found.`;
  exportBtn.disabled = false;
}

function csv(value) {
  const text = String(value ?? "");
  return `"${text.replaceAll('"','""')}"`;
}
