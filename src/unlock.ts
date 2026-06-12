import { unzipSync } from "fflate";

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

  return { kind: "already-unlocked" };
}

function startsWith(data: Uint8Array, magic: number[]): boolean {
  return magic.every((byte, i) => data[i] === byte);
}
