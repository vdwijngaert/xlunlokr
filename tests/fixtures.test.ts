import { existsSync, readFileSync, readdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { unlockWorkbook } from "../src/unlock";

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "..", "fixtures");
const fixtures = existsSync(fixturesDir)
  ? readdirSync(fixturesDir).filter((name) => /\.(xlsx|xlsm|xltx|xltm)$/i.test(name))
  : [];

describe.skipIf(fixtures.length === 0)("real locked fixtures (local only)", () => {
  it.each(fixtures)("unlocks %s", (name) => {
    const data = new Uint8Array(readFileSync(join(fixturesDir, name)));
    const result = unlockWorkbook(data);
    expect(result.kind).toBe("unlocked");
    if (result.kind !== "unlocked") return;
    expect(result.sheetProtectionsRemoved + result.workbookProtectionsRemoved).toBeGreaterThan(0);
    expect(unlockWorkbook(result.data)).toEqual({ kind: "already-unlocked" });
  });
});
