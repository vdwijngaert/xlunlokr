import { describe, expect, it } from "vitest";
import { unzipSync } from "fflate";
import { buildXlsx } from "./helpers/build-xlsx";

const decoder = new TextDecoder();

describe("buildXlsx", () => {
  it("produces a zip with workbook and one sheet by default", () => {
    const entries = unzipSync(buildXlsx());
    expect(Object.keys(entries)).toEqual(
      expect.arrayContaining([
        "[Content_Types].xml",
        "_rels/.rels",
        "xl/workbook.xml",
        "xl/_rels/workbook.xml.rels",
        "xl/worksheets/sheet1.xml",
      ]),
    );
    expect(decoder.decode(entries["xl/worksheets/sheet1.xml"])).not.toContain("<sheetProtection");
    expect(decoder.decode(entries["xl/workbook.xml"])).not.toContain("<workbookProtection");
  });

  it("adds sheetProtection to every sheet when requested", () => {
    const entries = unzipSync(buildXlsx({ sheets: 2, sheetProtection: true }));
    expect(decoder.decode(entries["xl/worksheets/sheet1.xml"])).toContain("<sheetProtection");
    expect(decoder.decode(entries["xl/worksheets/sheet2.xml"])).toContain("<sheetProtection");
  });

  it("adds workbookProtection when requested", () => {
    const entries = unzipSync(buildXlsx({ workbookProtection: true }));
    expect(decoder.decode(entries["xl/workbook.xml"])).toContain("<workbookProtection");
  });

  it("writes paired-form tags when requested", () => {
    const entries = unzipSync(buildXlsx({ sheetProtection: true, pairedTags: true }));
    expect(decoder.decode(entries["xl/worksheets/sheet1.xml"])).toContain("</sheetProtection>");
  });
});
