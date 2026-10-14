import { expect, test } from "vitest";
import { mapMathText, mathPlainText, mathToXml, type MathNode } from "../src/core/math";
import { mathPreviewText } from "../src/web/math-text";

const r = (t: string): MathNode => ({ tag: "m:r", attrs: {}, children: [{ tag: "m:t", attrs: {}, children: [t] }] });
const el = (tag: string, ...children: MathNode[]): MathNode => ({ tag, attrs: {}, children });

test("mathToXml 會跳脫特殊字元", () => {
  expect(mathToXml(el("m:oMath", r("a<b&\"c\"")))).toBe(
    `<m:oMath><m:r><m:t xml:space="preserve">a&lt;b&amp;&quot;c&quot;</m:t></m:r></m:oMath>`,
  );
});

test("mapMathText 同 mathPlainText", () => {
  const root = el("m:oMath", r("x"), r("+1"));
  expect(mathPlainText(mapMathText(root, (s) => s.toUpperCase()))).toBe("X+1");
  expect(mathPlainText(root)).toBe("x+1"); // 唔改原本
});

test("mathPreviewText 將分數、上下標、根號、括號、求和轉成一行文字", () => {
  const frac = el("m:f", el("m:num", r("a")), el("m:den", r("b+c")));
  const sup = el("m:sSup", el("m:e", r("x")), el("m:sup", r("2")));
  const sub = el("m:sSub", el("m:e", r("q")), el("m:sub", r("i")));
  const rad = el("m:rad", el("m:radPr"), el("m:deg"), el("m:e", r("n")));
  const delim: MathNode = { tag: "m:d", attrs: {}, children: [el("m:e", r("1-q"))] };
  const nary: MathNode = {
    tag: "m:nary",
    attrs: {},
    children: [
      { tag: "m:naryPr", attrs: {}, children: [{ tag: "m:chr", attrs: { "m:val": "∑" }, children: [] }] },
      el("m:sub", r("i=1")),
      el("m:sup", r("n")),
      el("m:e", r("i")),
    ],
  };
  expect(mathPreviewText(el("m:oMath", frac))).toBe("a/(b+c)");
  expect(mathPreviewText(el("m:oMath", sup, r("+"), sub))).toBe("x^2+q_i");
  expect(mathPreviewText(el("m:oMath", rad))).toBe("√(n)");
  expect(mathPreviewText(el("m:oMath", delim))).toBe("(1-q)");
  expect(mathPreviewText(el("m:oMath", nary))).toBe("∑_(i=1)^n i");
});
