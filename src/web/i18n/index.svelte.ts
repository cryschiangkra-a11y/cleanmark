import en from "./en";
import zhHant, { type Messages } from "./zh-Hant";

export type Lang = "zh-Hant" | "en";
export type MessageKey = keyof Messages;

const DICTS: Record<Lang, Messages> = { "zh-Hant": zhHant, en };
const STORAGE_KEY = "cleanmark.lang";

function initialLang(): Lang {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved === "zh-Hant" || saved === "en") return saved;
  } catch {
    // 私密瀏覽等情況讀唔到，用瀏覽器語言
  }
  return navigator.language.toLowerCase().startsWith("zh") ? "zh-Hant" : "en";
}

export const i18n = $state({ lang: initialLang() });

export function setLang(lang: Lang): void {
  i18n.lang = lang;
  document.documentElement.lang = lang;
  try {
    localStorage.setItem(STORAGE_KEY, lang);
  } catch {
    // 儲存唔到就算，下次用返預設
  }
}

export function t(key: MessageKey, params: Record<string, string | number> = {}): string {
  let s = DICTS[i18n.lang][key];
  for (const [k, v] of Object.entries(params)) s = s.replaceAll(`{${k}}`, String(v));
  return s;
}
