// 「內容清單」：讀取器嘅輸出、輸出器嘅輸入。
// 只有呢度定義咗嘅欄位會被搬去新檔案（白名單），水印同 metadata 冇欄位可以裝。

export type Align = "left" | "center" | "right" | "justify";

export interface RunStyle {
  font?: string; // 西文字型
  fontEastAsia?: string; // 中日韓字型
  size: number; // pt
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  color?: string; // RRGGBB，undefined = 自動（黑色）
  highlight?: string; // OOXML highlight 名稱，例如 "yellow"
  vertAlign?: "superscript" | "subscript";
}

export interface TextInline {
  type: "text";
  text: string; // 可以包含 "\t"
  style: RunStyle;
  link?: string; // 外部超連結網址
}

export interface BreakInline {
  type: "break"; // 行內換行
}

export interface FootnoteRefInline {
  type: "footnoteRef";
  id: number;
}

/** 方程式 (OMML) 嘅元素樹：只保留白名單內嘅元素同屬性，字串 = 文字 */
export interface MathNode {
  tag: string; // 例如 "m:oMath"、"m:f"、"w:rPr"
  attrs: Record<string, string>; // 例如 { "m:val": "∑" }
  children: (MathNode | string)[];
}

export interface MathInline {
  type: "math";
  display: boolean; // true = 獨立一行（m:oMathPara）
  root: MathNode;
}

export type Inline = TextInline | BreakInline | FootnoteRefInline | MathInline;

export interface ListRef {
  numId: string;
  level: number; // 0 起
}

export interface LineSpacing {
  line: number; // twip；rule 係 auto 時 240 = 單行
  rule: "auto" | "exact" | "atLeast";
}

export interface Paragraph {
  type: "paragraph";
  id: string; // 例如 "p12"，俾報告同用家選擇參照
  heading: number; // 0 = 內文，1–6 = 標題層級
  align: Align;
  indentLeft: number; // twip
  indentFirstLine: number; // twip，負數 = 懸掛縮排
  spacingBefore: number; // twip
  spacingAfter: number; // twip
  lineSpacing: LineSpacing;
  list?: ListRef;
  inlines: Inline[];
}

export interface TableCell {
  blocks: Block[];
  colSpan: number;
  rowSpan: number; // 垂直合併嘅起點格；被合併嘅格唔會出現
}

export interface Table {
  type: "table";
  columnWidths: number[]; // twip
  rows: TableCell[][];
  borders: boolean;
}

export interface ImageSlot {
  type: "image";
  index: number; // 1 起，全文件順序
  widthEmu: number;
  heightEmu: number;
  context: string; // 位置提示：前面最近嘅標題，冇標題就用前一段文字
  mediaPath: string; // 原檔入面嘅路徑，例如 "word/media/image1.png"
  contentType: "png" | "jpg" | "gif" | "bmp" | "webp" | "tiff" | "other";
  keep: boolean; // true = 保留原圖；false = 佔位框（由 applyChoices 設定）
}

/** 可以原圖放入新檔案嘅圖片格式 */
export const KEEPABLE_IMAGE_TYPES: ReadonlySet<ImageSlot["contentType"]> = new Set(["png", "jpg", "gif", "bmp", "webp", "tiff"]);

export interface PageBreak {
  type: "pageBreak";
}

export type Block = Paragraph | Table | ImageSlot | PageBreak;

export interface ListLevel {
  format: string; // OOXML numFmt，例如 "decimal"、"bullet"、"lowerLetter"
  text: string; // OOXML lvlText，例如 "%1."、"•"
  start: number;
  indentLeft: number; // twip
  hanging: number; // twip
}

export interface PageSetup {
  widthTwip: number;
  heightTwip: number;
  landscape: boolean;
  margins: { top: number; right: number; bottom: number; left: number };
}

export interface CmDocument {
  kind: "docx";
  page: PageSetup;
  body: Block[];
  header: Block[]; // 只有段落同表格
  footer: Block[];
  footnotes: Record<number, Block[]>;
  lists: Record<string, ListLevel[]>; // numId → 每層設定
  media: Record<string, Uint8Array>; // mediaPath → 原圖 bytes（只喺保留原圖時用）
}

// ── 報告 ──────────────────────────────────────────────

export type VisibleCode =
  | "wordTextWatermark" // Word 內建文字水印
  | "wordPictureWatermark" // Word 內建圖片水印
  | "headerShape" // 頁首頁尾嘅圖形或者文字方塊
  | "headerImage" // 頁首頁尾嘅圖片
  | "behindText" // 擺喺文字後面嘅圖形
  | "pageBackground" // 頁面背景
  | "textArt"; // 內文入面嘅 WordArt

export interface VisibleItem {
  code: VisibleCode;
  part: string;
  text?: string;
}

export interface MetadataItem {
  part: string;
  field: string;
  value: string;
}

export interface PhraseItem {
  paragraphId: string;
  text: string; // 段落文字（最多 160 字）
  match: string;
}

export interface ImageItem {
  index: number;
  context: string;
  keepable: boolean; // 原圖格式可唔可以放入新檔案
  tags: ("exif" | "xmp" | "c2pa")[];
}

export interface HiddenCharSummary {
  removed: Record<string, number>; // 例如 { "U+200B": 23 }
  replaced: Record<string, number>; // 換成普通空格
  kept: number; // 因為語言需要而保留
}

export interface Report {
  fileKind: "docx";
  visible: VisibleItem[];
  metadata: MetadataItem[];
  c2pa: string[]; // 搵到 C2PA 嘅 part
  hidden: HiddenCharSummary;
  phrases: PhraseItem[];
  images: ImageItem[];
}

// ── 用家選擇 ─────────────────────────────────────────

export interface Choices {
  removeParagraphs: string[]; // 要刪除嘅段落 id
  keepImages: number[]; // 保留原圖嘅圖片序號，其餘用佔位框
}
