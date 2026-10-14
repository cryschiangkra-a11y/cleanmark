import { describe, expect, test } from "vitest";
import { cleanText, emptySummary, mergeSummary } from "../src/core/clean/hidden-chars";

describe("cleanText", () => {
  test("清走零寬度空格同 BOM", () => {
    const r = cleanText("A\u200BB\uFEFFC");
    expect(r.text).toBe("ABC");
    expect(r.summary.removed).toEqual({ "U+200B": 1, "U+FEFF": 1 });
  });

  test("清走 Tag 字元（可以收埋隱形句子）", () => {
    const hidden = String.fromCodePoint(0xe0041, 0xe0042);
    const r = cleanText(`hi${hidden}`);
    expect(r.text).toBe("hi");
    expect(r.summary.removed).toEqual({ "U+E0041": 1, "U+E0042": 1 });
  });

  test("只有 U+202F 窄不換行空格換成普通空格；全形、細空格等排版空格原樣保留", () => {
    const r = cleanText("a\u202Fb\u2003c\u3000d\u2009e\u205Ff");
    expect(r.text).toBe("a b\u2003c\u3000d\u2009e\u205Ff");
    expect(r.summary.replaced).toEqual({ "U+202F": 1 });
  });

  test("中英文之間嘅 ZWJ / ZWNJ 清走", () => {
    const r = cleanText("中\u200D文 a\u200Cb");
    expect(r.text).toBe("中文 ab");
    expect(r.summary.removed).toEqual({ "U+200D": 1, "U+200C": 1 });
  });

  test("emoji 家庭序列入面嘅 ZWJ 保留", () => {
    const family = "👨\u200D👩\u200D👧";
    const r = cleanText(family);
    expect(r.text).toBe(family);
    expect(r.summary.kept).toBe(2);
  });

  test("膚色 emoji 後面嘅 ZWJ 保留", () => {
    const coder = "👩🏽\u200D💻";
    expect(cleanText(coder).text).toBe(coder);
  });

  test("波斯文入面嘅 ZWNJ 保留", () => {
    const word = "می\u200Cخواهم";
    const r = cleanText(word);
    expect(r.text).toBe(word);
    expect(r.summary.kept).toBe(1);
  });

  test("方向標記：段落有希伯來文就保留，冇就清走", () => {
    expect(cleanText("שלום\u200F abc").text).toBe("שלום\u200F abc");
    expect(cleanText("abc\u200F def").text).toBe("abc def");
  });

  test("方向覆寫一定清走，就算有阿拉伯文", () => {
    expect(cleanText("مرحبا\u202Eabc").text).toBe("مرحباabc");
  });

  test("異體選擇符：emoji 同漢字後面保留，英文後面清走", () => {
    expect(cleanText("❤\uFE0F").text).toBe("❤\uFE0F");
    expect(cleanText("葛\uFE00").text).toBe("葛\uFE00");
    expect(cleanText("1\uFE0F⃣").text).toBe("1\uFE0F⃣");
    expect(cleanText("a\uFE0Fb").text).toBe("ab");
  });

  test("上下文可以由前後文字片段提供", () => {
    const r = cleanText("\u200D👩", { before: "👨", after: "", rtl: false });
    expect(r.text).toBe("\u200D👩");
  });

  test("冇隱形字元就原封不動", () => {
    const r = cleanText("Hello 你好");
    expect(r.text).toBe("Hello 你好");
    expect(r.summary).toEqual(emptySummary());
  });
});

test("mergeSummary 加埋兩份統計", () => {
  const a = { removed: { "U+200B": 1 }, replaced: {}, kept: 1 };
  const b = { removed: { "U+200B": 2, "U+FEFF": 1 }, replaced: { "U+202F": 1 }, kept: 0 };
  expect(mergeSummary(a, b)).toEqual({
    removed: { "U+200B": 3, "U+FEFF": 1 },
    replaced: { "U+202F": 1 },
    kept: 1,
  });
});
