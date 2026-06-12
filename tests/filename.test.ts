import { describe, expect, it } from "vitest";
import { unlockedFilename } from "../src/filename";

describe("unlockedFilename", () => {
  it("inserts -unlocked before the extension", () => {
    expect(unlockedFilename("report.xlsx")).toBe("report-unlocked.xlsx");
  });

  it("keeps the original extension for macro workbooks", () => {
    expect(unlockedFilename("budget.xlsm")).toBe("budget-unlocked.xlsm");
  });

  it("uses the last dot when the name contains several", () => {
    expect(unlockedFilename("plan.v2.xlsx")).toBe("plan.v2-unlocked.xlsx");
  });

  it("appends when there is no extension", () => {
    expect(unlockedFilename("workbook")).toBe("workbook-unlocked");
  });
});
