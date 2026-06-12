import { describe, expect, it } from "vitest";
import { strToU8, unzipSync, zipSync } from "fflate";
import { unlockWorkbook } from "../src/unlock";
import { buildXlsx } from "./helpers/build-xlsx";

describe("unlockWorkbook input detection", () => {
  it("reports encrypted-or-legacy for OLE/CFB containers", () => {
    const cfb = new Uint8Array(512);
    cfb.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
    expect(unlockWorkbook(cfb)).toEqual({ kind: "error", reason: "encrypted-or-legacy" });
  });

  it("reports not-excel for unknown bytes", () => {
    expect(unlockWorkbook(strToU8("hello, world"))).toEqual({ kind: "error", reason: "not-excel" });
  });

  it("reports not-excel for empty input", () => {
    expect(unlockWorkbook(new Uint8Array(0))).toEqual({ kind: "error", reason: "not-excel" });
  });

  it("reports invalid-zip for corrupt zip data", () => {
    const corrupt = new Uint8Array(64).fill(0xff);
    corrupt.set([0x50, 0x4b, 0x03, 0x04]);
    expect(unlockWorkbook(corrupt)).toEqual({ kind: "error", reason: "invalid-zip" });
  });

  it("reports not-excel for a zip that is not a workbook", () => {
    const docx = zipSync({ "word/document.xml": strToU8("<document/>") });
    expect(unlockWorkbook(docx)).toEqual({ kind: "error", reason: "not-excel" });
  });

  it("reports already-unlocked for a workbook without protection", () => {
    expect(unlockWorkbook(buildXlsx())).toEqual({ kind: "already-unlocked" });
  });

  it("reports not-excel for an empty zip archive", () => {
    const empty = zipSync({});
    expect(unlockWorkbook(empty)).toEqual({ kind: "error", reason: "not-excel" });
  });
});

describe("unlockWorkbook stripping", () => {
  const decoder = new TextDecoder();

  it("removes sheetProtection from every worksheet", () => {
    const result = unlockWorkbook(buildXlsx({ sheets: 3, sheetProtection: true }));
    expect(result.kind).toBe("unlocked");
    if (result.kind !== "unlocked") return;
    expect(result.sheetProtectionsRemoved).toBe(3);
    expect(result.workbookProtectionsRemoved).toBe(0);
    const entries = unzipSync(result.data);
    for (const n of [1, 2, 3]) {
      const xml = decoder.decode(entries[`xl/worksheets/sheet${n}.xml`]);
      expect(xml).not.toContain("sheetProtection");
      expect(xml).toContain("<sheetData>"); // regex must not eat neighbors
      expect(xml).toContain("</worksheet>");
    }
  });

  it("removes workbookProtection from workbook.xml", () => {
    const result = unlockWorkbook(buildXlsx({ workbookProtection: true }));
    expect(result.kind).toBe("unlocked");
    if (result.kind !== "unlocked") return;
    expect(result.sheetProtectionsRemoved).toBe(0);
    expect(result.workbookProtectionsRemoved).toBe(1);
    const xml = decoder.decode(unzipSync(result.data)["xl/workbook.xml"]);
    expect(xml).not.toContain("workbookProtection");
    expect(xml).toContain("<sheets>");
  });

  it("removes paired-form protection tags", () => {
    const result = unlockWorkbook(
      buildXlsx({ sheetProtection: true, workbookProtection: true, pairedTags: true }),
    );
    expect(result.kind).toBe("unlocked");
    if (result.kind !== "unlocked") return;
    expect(result.sheetProtectionsRemoved).toBe(1);
    expect(result.workbookProtectionsRemoved).toBe(1);
    const entries = unzipSync(result.data);
    expect(decoder.decode(entries["xl/worksheets/sheet1.xml"])).not.toContain("Protection");
    expect(decoder.decode(entries["xl/workbook.xml"])).not.toContain("Protection");
  });

  it("leaves untouched entries byte-identical and preserves entry order", () => {
    const input = buildXlsx({ sheets: 2, sheetProtection: true });
    const before = unzipSync(input);
    const result = unlockWorkbook(input);
    expect(result.kind).toBe("unlocked");
    if (result.kind !== "unlocked") return;
    const after = unzipSync(result.data);
    expect(Object.keys(after)).toEqual(Object.keys(before));
    for (const name of Object.keys(before)) {
      if (name.startsWith("xl/worksheets/")) continue;
      expect(after[name]).toEqual(before[name]);
    }
  });

  it("reports already-unlocked when run again on its own output", () => {
    const result = unlockWorkbook(buildXlsx({ sheetProtection: true, workbookProtection: true }));
    expect(result.kind).toBe("unlocked");
    if (result.kind !== "unlocked") return;
    expect(unlockWorkbook(result.data)).toEqual({ kind: "already-unlocked" });
  });

  it("removes namespace-prefixed protection tags (paired form)", () => {
    const result = unlockWorkbook(
      buildXlsx({ sheetProtection: true, workbookProtection: true, nsPrefix: "x", pairedTags: true }),
    );
    expect(result.kind).toBe("unlocked");
    if (result.kind !== "unlocked") return;
    expect(result.sheetProtectionsRemoved).toBe(1);
    expect(result.workbookProtectionsRemoved).toBe(1);
    const entries = unzipSync(result.data);
    expect(decoder.decode(entries["xl/worksheets/sheet1.xml"])).not.toContain("Protection");
    expect(decoder.decode(entries["xl/workbook.xml"])).not.toContain("Protection");
  });

  it("ignores escaped protection-like text in shared strings", () => {
    const entries = unzipSync(buildXlsx());
    entries["xl/sharedStrings.xml"] = strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">' +
        '<si><t>&lt;sheetProtection sheet="1"/&gt;</t></si></sst>',
    );
    expect(unlockWorkbook(zipSync(entries))).toEqual({ kind: "already-unlocked" });
  });
});
