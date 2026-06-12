import { unzipSync, zipSync } from "fflate";

export type UnlockError = "encrypted-or-legacy" | "not-excel" | "invalid-zip";

export type UnlockResult =
  | {
      kind: "unlocked";
      data: Uint8Array;
      sheetProtectionsRemoved: number;
      workbookProtectionsRemoved: number;
    }
  | { kind: "already-unlocked" }
  | { kind: "error"; reason: UnlockError };

const CFB_MAGIC = [0xd0, 0xcf, 0x11, 0xe0];
const ZIP_MAGIC = [0x50, 0x4b];

// Both elements are childless per the OOXML schema (CT_SheetProtection /
// CT_WorkbookProtection), so removing the element via regex cannot orphan
// children. Handles self-closing and paired forms, including namespace-prefixed
// forms (e.g. <x:sheetProtection>) produced by third-party generators such as
// Apache POI.
const SHEET_PROTECTION_RE =
  /<(?:[\w.-]+:)?sheetProtection\b[^>]*(?:\/>|>\s*<\/(?:[\w.-]+:)?sheetProtection\s*>)/g;
const WORKBOOK_PROTECTION_RE =
  /<(?:[\w.-]+:)?workbookProtection\b[^>]*(?:\/>|>\s*<\/(?:[\w.-]+:)?workbookProtection\s*>)/g;

export function unlockWorkbook(data: Uint8Array): UnlockResult {
  // OLE/CFB container: either an open-password-encrypted workbook or a legacy .xls.
  if (startsWith(data, CFB_MAGIC)) {
    return { kind: "error", reason: "encrypted-or-legacy" };
  }
  if (!startsWith(data, ZIP_MAGIC)) {
    return { kind: "error", reason: "not-excel" };
  }

  let entries: Record<string, Uint8Array>;
  try {
    entries = unzipSync(data);
  } catch {
    return { kind: "error", reason: "invalid-zip" };
  }

  if (!("xl/workbook.xml" in entries)) {
    return { kind: "error", reason: "not-excel" };
  }

  const decoder = new TextDecoder();
  const encoder = new TextEncoder();
  let sheetProtectionsRemoved = 0;
  let workbookProtectionsRemoved = 0;

  for (const name of Object.keys(entries)) {
    if (!/^xl\/.*\.xml$/.test(name)) continue;
    const xml = decoder.decode(entries[name]);
    const afterSheet = strip(xml, SHEET_PROTECTION_RE);
    const afterWorkbook = strip(afterSheet.xml, WORKBOOK_PROTECTION_RE);
    if (afterSheet.count + afterWorkbook.count === 0) continue;
    sheetProtectionsRemoved += afterSheet.count;
    workbookProtectionsRemoved += afterWorkbook.count;
    entries[name] = encoder.encode(afterWorkbook.xml);
  }

  if (sheetProtectionsRemoved + workbookProtectionsRemoved === 0) {
    return { kind: "already-unlocked" };
  }

  return {
    kind: "unlocked",
    data: zipSync(entries),
    sheetProtectionsRemoved,
    workbookProtectionsRemoved,
  };
}

function startsWith(data: Uint8Array, magic: number[]): boolean {
  if (magic.length === 0 || data.length < magic.length) return false;
  return magic.every((byte, i) => data[i] === byte);
}

function strip(xml: string, re: RegExp): { xml: string; count: number } {
  let count = 0;
  const result = xml.replace(re, () => {
    count++;
    return "";
  });
  return { xml: result, count };
}
