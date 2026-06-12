import { unlockWorkbook, type UnlockError, type UnlockResult } from "./unlock";
import { unlockedFilename } from "./filename";
import "./style.css";

const dropZone = document.getElementById("drop-zone") as HTMLDivElement;
const fileInput = document.getElementById("file-input") as HTMLInputElement;
const resultsList = document.getElementById("results") as HTMLUListElement;
const downloadAllButton = document.getElementById("download-all") as HTMLButtonElement;

const EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

const ERROR_MESSAGES: Record<UnlockError, string> = {
  "encrypted-or-legacy":
    "Can't unlock: this file needs a password to open (real encryption) or is a legacy .xls file",
  "not-excel": "Not an Excel file",
  "invalid-zip": "File appears to be corrupted",
};

interface PendingDownload {
  url: string;
  filename: string;
}

// Running log across drops: "Download all" re-downloads every file unlocked
// this session, matching the cumulative results list.
const downloads: PendingDownload[] = [];

dropZone.addEventListener("click", () => fileInput.click());
dropZone.addEventListener("keydown", (event) => {
  if (event.key === "Enter" || event.key === " ") {
    event.preventDefault();
    fileInput.click();
  }
});
dropZone.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropZone.classList.add("dragging");
});
dropZone.addEventListener("dragleave", () => dropZone.classList.remove("dragging"));
dropZone.addEventListener("drop", (event) => {
  event.preventDefault();
  dropZone.classList.remove("dragging");
  if (event.dataTransfer) void processFiles(event.dataTransfer.files);
});
fileInput.addEventListener("change", () => {
  if (fileInput.files) void processFiles(fileInput.files);
  fileInput.value = "";
});
downloadAllButton.addEventListener("click", () => {
  for (const download of downloads) triggerDownload(download);
});

async function processFiles(files: FileList): Promise<void> {
  for (const file of files) {
    try {
      const data = new Uint8Array(await file.arrayBuffer());
      addResultRow(file.name, unlockWorkbook(data));
    } catch {
      appendRow(file.name, "error", "Couldn't read this file");
    }
  }
  downloadAllButton.hidden = downloads.length < 2;
}

function appendRow(name: string, kind: string, statusText: string): HTMLLIElement {
  const row = document.createElement("li");
  row.className = `result ${kind}`;

  const title = document.createElement("span");
  title.className = "filename";
  title.textContent = name;

  const status = document.createElement("span");
  status.className = "status";
  status.textContent = statusText;

  row.append(title, status);
  resultsList.append(row);
  return row;
}

function addResultRow(name: string, result: UnlockResult): void {
  const row = appendRow(name, result.kind, statusMessage(result));

  if (result.kind === "unlocked") {
    const blob = new Blob([result.data.slice().buffer], { type: EXCEL_MIME });
    const download: PendingDownload = {
      url: URL.createObjectURL(blob),
      filename: unlockedFilename(name),
    };
    downloads.push(download);

    const button = document.createElement("button");
    button.textContent = "Download";
    button.addEventListener("click", () => triggerDownload(download));
    row.append(button);
  }
}

function statusMessage(result: UnlockResult): string {
  switch (result.kind) {
    case "unlocked":
      return `Unlocked — removed ${removedSummary(result)}`;
    case "already-unlocked":
      return "No protection found — this file is already unlocked";
    case "error":
      return ERROR_MESSAGES[result.reason];
  }
}

function removedSummary(result: {
  sheetProtectionsRemoved: number;
  workbookProtectionsRemoved: number;
}): string {
  const parts: string[] = [];
  if (result.sheetProtectionsRemoved > 0) {
    parts.push(plural(result.sheetProtectionsRemoved, "sheet protection"));
  }
  if (result.workbookProtectionsRemoved > 0) {
    parts.push(plural(result.workbookProtectionsRemoved, "workbook protection"));
  }
  return parts.join(" and ");
}

function plural(count: number, noun: string): string {
  return `${count} ${noun}${count === 1 ? "" : "s"}`;
}

function triggerDownload(download: PendingDownload): void {
  const anchor = document.createElement("a");
  anchor.href = download.url;
  anchor.download = download.filename;
  anchor.click();
}
