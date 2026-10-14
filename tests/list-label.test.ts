import { expect, test } from "vitest";
import type { Block, ListLevel } from "../src/core/model";
import { formatNumber, listLabels } from "../src/web/list-label";

test("formatNumber", () => {
  expect(formatNumber(4, "upperRoman")).toBe("IV");
  expect(formatNumber(28, "lowerLetter")).toBe("ab");
  expect(formatNumber(3, "decimal")).toBe("3");
});

test("listLabels 計多層編號，上層前進時重設下層", () => {
  const lvl = (format: string, text: string): ListLevel => ({ format, text, start: 1, indentLeft: 0, hanging: 0 });
  const lists = { "1": [lvl("decimal", "%1."), lvl("lowerLetter", "%1.%2)"), lvl("bullet", "▪")] };
  const item = (id: string, level: number): Block => ({
    type: "paragraph", id, heading: 0, align: "left", indentLeft: 0, indentFirstLine: 0,
    spacingBefore: 0, spacingAfter: 0, lineSpacing: { line: 240, rule: "auto" },
    list: { numId: "1", level }, inlines: [],
  });
  const labels = listLabels([item("a", 0), item("b", 1), item("c", 1), item("d", 2), item("e", 0), item("f", 1)], lists);
  expect([...labels.values()]).toEqual(["1.", "1.a)", "1.b)", "▪", "2.", "2.a)"]);
});
