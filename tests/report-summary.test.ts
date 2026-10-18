import { expect, test } from "vitest";
import type { Report } from "../src/core/model";
import { findingCount, hiddenTotal } from "../src/web/report-summary";

const base: Report = {
  fileKind: "docx", visible: [], metadata: [], c2pa: [],
  hidden: { removed: {}, replaced: {}, kept: 0 }, phrases: [], images: [],
  suspicious: { homoglyphs: [], apostrophes: 0, specialSpaces: {} },
};

test("乾淨文件 = 0 項", () => {
  expect(findingCount(base)).toBe(0);
});

test("隱形字元無論幾多個都當一項", () => {
  const r = {
    ...base,
    visible: [{ code: "headerImage" as const, part: "word/header1.xml" }],
    metadata: [{ part: "docProps/core.xml", field: "dc:creator", value: "x" }],
    hidden: { removed: { "U+200B": 23 }, replaced: { "U+202F": 2 }, kept: 1 },
  };
  expect(hiddenTotal(r)).toBe(25);
  expect(findingCount(r)).toBe(3);
});
