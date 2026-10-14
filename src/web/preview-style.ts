import type { Paragraph, RunStyle } from "../core/model";

const HIGHLIGHT: Record<string, string> = {
  yellow: "#ffff00", green: "#00ff00", cyan: "#00ffff", magenta: "#ff00ff", blue: "#0000ff",
  red: "#ff0000", darkBlue: "#000080", darkCyan: "#008080", darkGreen: "#008000",
  darkMagenta: "#800080", darkRed: "#800000", darkYellow: "#808000", darkGray: "#808080",
  lightGray: "#c0c0c0", black: "#000000", white: "#ffffff",
};

const quote = (f: string) => `"${f.replace(/"/g, "")}"`;
const TWIP_TO_PX = 96 / 1440;

export function runCss(s: RunStyle): string {
  const fonts = [s.font, s.fontEastAsia].filter(Boolean).map((f) => quote(f!));
  const deco = [s.underline && "underline", s.strike && "line-through"].filter(Boolean).join(" ");
  return [
    fonts.length ? `font-family:${fonts.join(",")},inherit` : "",
    `font-size:${s.size}pt`,
    `font-weight:${s.bold ? 700 : 400}`,
    `font-style:${s.italic ? "italic" : "normal"}`,
    deco ? `text-decoration:${deco}` : "",
    s.color ? `color:#${s.color}` : "",
    s.highlight && HIGHLIGHT[s.highlight] ? `background:${HIGHLIGHT[s.highlight]}` : "",
    s.vertAlign === "superscript" ? "vertical-align:super;font-size:0.7em" : "",
    s.vertAlign === "subscript" ? "vertical-align:sub;font-size:0.7em" : "",
  ]
    .filter(Boolean)
    .join(";");
}

export function paragraphCss(p: Paragraph): string {
  const line = p.lineSpacing.rule === "auto" ? String(p.lineSpacing.line / 240 * 1.15) : `${(p.lineSpacing.line / 20).toFixed(1)}pt`;
  return [
    `text-align:${p.align}`,
    `margin:${Math.round(p.spacingBefore * TWIP_TO_PX)}px 0 ${Math.round(p.spacingAfter * TWIP_TO_PX)}px ${Math.round(p.indentLeft * TWIP_TO_PX)}px`,
    `text-indent:${Math.round(p.indentFirstLine * TWIP_TO_PX)}px`,
    `line-height:${line}`,
  ].join(";");
}

export const EMU_TO_PX = 96 / 914400;
