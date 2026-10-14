import { expect, test } from "vitest";
import en from "../src/web/i18n/en";
import zhHant from "../src/web/i18n/zh-Hant";

test("兩種語言嘅 key 一樣", () => {
  expect(Object.keys(en).sort()).toEqual(Object.keys(zhHant).sort());
});

test("參數 {x} 兩種語言都一樣", () => {
  const params = (s: string) => (s.match(/\{\w+\}/g) ?? []).sort();
  for (const k of Object.keys(zhHant) as (keyof typeof zhHant)[]) {
    expect(params(en[k]), k).toEqual(params(zhHant[k]));
  }
});
