import { describe, expect, it } from "vitest";
import { strToU8, zipSync } from "fflate";
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
});
