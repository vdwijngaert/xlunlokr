import { strToU8, zipSync, type Zippable } from "fflate";

export interface BuildOptions {
  /** Number of worksheets (default 1). */
  sheets?: number;
  /** Add a <sheetProtection> element to every worksheet. */
  sheetProtection?: boolean;
  /** Add a <workbookProtection> element to the workbook. */
  workbookProtection?: boolean;
  /** Write protection tags as <x ...></x> instead of self-closing <x .../>. */
  pairedTags?: boolean;
  /** Qualify protection tags with this namespace prefix (e.g. "x"). */
  nsPrefix?: string;
}

const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

export function buildXlsx(opts: BuildOptions = {}): Uint8Array {
  const sheetCount = opts.sheets ?? 1;
  const nums = Array.from({ length: sheetCount }, (_, i) => i + 1);
  const paired = opts.pairedTags ?? false;
  const prefix = opts.nsPrefix ? `${opts.nsPrefix}:` : "";

  // Attribute values deliberately contain base64 chars (+ / =) to exercise the
  // stripping regex against realistic hashValue/saltValue content.
  const sheetProtection = opts.sheetProtection
    ? tag(
        `${prefix}sheetProtection`,
        'algorithmName="SHA-512" hashValue="q+r/s7w9==" saltValue="c2FsdA==" spinCount="100000" sheet="1" objects="1" scenarios="1"',
        paired,
      )
    : "";
  const workbookProtection = opts.workbookProtection
    ? tag(`${prefix}workbookProtection`, 'lockStructure="1"', paired)
    : "";

  const files: Zippable = {
    "[Content_Types].xml": strToU8(
      XML_HEADER +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        nums
          .map(
            (n) =>
              `<Override PartName="/xl/worksheets/sheet${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
          )
          .join("") +
        "</Types>",
    ),
    "_rels/.rels": strToU8(
      XML_HEADER +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        "</Relationships>",
    ),
    "xl/workbook.xml": strToU8(
      XML_HEADER +
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        workbookProtection +
        "<sheets>" +
        nums.map((n) => `<sheet name="Sheet${n}" sheetId="${n}" r:id="rId${n}"/>`).join("") +
        "</sheets></workbook>",
    ),
    "xl/_rels/workbook.xml.rels": strToU8(
      XML_HEADER +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        nums
          .map(
            (n) =>
              `<Relationship Id="rId${n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${n}.xml"/>`,
          )
          .join("") +
        "</Relationships>",
    ),
  };

  for (const n of nums) {
    files[`xl/worksheets/sheet${n}.xml`] = strToU8(
      XML_HEADER +
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        `<sheetData><row r="1"><c r="A1"><v>${n}</v></c></row></sheetData>` +
        sheetProtection +
        "</worksheet>",
    );
  }

  return zipSync(files);
}

function tag(name: string, attrs: string, paired: boolean): string {
  return paired ? `<${name} ${attrs}></${name}>` : `<${name} ${attrs}/>`;
}
