import { describe, expect, test } from "vitest";
import { fixSuspicious, scanSuspicious, splitByLengths } from "../src/core/clean/suspicious";

const ALL = { homoglyphs: true, apostrophes: true, specialSpaces: true };
const NONE = { homoglyphs: false, apostrophes: false, specialSpaces: false };

describe("同形字", () => {
  test("英文字入面夾雜西里爾 / 希臘字母：認出並換返英文字母", () => {
    const text = "Clаude can сreаte ΟPENAI content."; // а с а 係西里爾，Ο 係希臘
    const hits = scanSuspicious(text);
    expect(hits.homoglyphs.map((h) => [h.word, h.fixed])).toEqual([
      ["Clаude", "Claude"],
      ["сreаte", "create"],
      ["ΟPENAI", "OPENAI"],
    ]);
    expect(fixSuspicious(text, ALL)).toBe("Claude can create OPENAI content.");
  });

  test("純俄文、純希臘文、方程式式嘅希臘字母唔郁", () => {
    for (const text of ["Привет мир", "Καλημέρα", "M + δ − 1 and α = 0.5", "10 μs delay"]) {
      expect(scanSuspicious(text).homoglyphs, text).toEqual([]);
      expect(fixSuspicious(text, ALL)).toBe(text);
    }
  });

  test("關閉就唔換", () => {
    expect(fixSuspicious("Clаude", NONE)).toBe("Clаude");
  });
});

describe("撇號變體", () => {
  test("英文字母之間嘅 U+02BC、U+02B9 換成正常撇號", () => {
    const text = "donʼt and itʹs";
    expect(scanSuspicious(text).apostrophes).toHaveLength(2);
    expect(fixSuspicious(text, ALL)).toBe("don’t and it’s");
  });

  test("文件用直撇號就跟直撇號", () => {
    expect(fixSuspicious("donʼt", { ...ALL, apostrophe: "'" })).toBe("don't");
  });

  test("烏克蘭文入面嘅撇號唔郁", () => {
    const text = "мʼясо";
    expect(scanSuspicious(text).apostrophes).toEqual([]);
    expect(fixSuspicious(text, ALL)).toBe(text);
  });
});

describe("特殊空格", () => {
  test("列出 U+00A0、U+2000–U+200A、U+205F；唔計 U+3000 全形空格", () => {
    const text = "Claude generated this　文字";
    expect(scanSuspicious(text).specialSpaces).toEqual({ "U+00A0": 1, "U+2009": 1 });
  });

  test("預設唔換；揀咗先換成普通空格", () => {
    const text = "a b c　d";
    expect(fixSuspicious(text, { ...ALL, specialSpaces: false })).toBe(text);
    expect(fixSuspicious(text, ALL)).toBe("a b c　d");
  });
});

test("一對一替換：長度唔變，可以按原本格式段切返開", () => {
  const runs = ["Cl", "аude ", "donʼ", "t"];
  const fixed = fixSuspicious(runs.join(""), ALL);
  expect(fixed.length).toBe(runs.join("").length);
  expect(splitByLengths(fixed, runs.map((r) => r.length))).toEqual(["Cl", "aude ", "don’", "t"]);
});
