// 清走圖片入面嘅 metadata（EXIF、XMP、C2PA、文字註解），影像同色彩設定檔保留

type Kind = "png" | "jpg" | "gif" | "bmp";

// PNG：只保留影像同顯示需要嘅 chunk
const PNG_KEEP = new Set(["IHDR", "PLTE", "IDAT", "IEND", "tRNS", "gAMA", "cHRM", "sRGB", "iCCP", "sBIT", "pHYs", "bKGD", "acTL", "fcTL", "fdAT"]);
const PNG_SIG = [137, 80, 78, 71, 13, 10, 26, 10];

function concat(parts: Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
  let o = 0;
  for (const p of parts) {
    out.set(p, o);
    o += p.length;
  }
  return out;
}

const startsWithAscii = (b: Uint8Array, s: string, at = 0) => Array.from(s).every((c, i) => b[at + i] === c.charCodeAt(0));

function stripPng(b: Uint8Array): Uint8Array | null {
  if (!PNG_SIG.every((x, i) => b[i] === x)) return null;
  const dv = new DataView(b.buffer, b.byteOffset, b.byteLength);
  const parts: Uint8Array[] = [b.slice(0, 8)];
  for (let o = 8; o + 12 <= b.length; ) {
    const len = dv.getUint32(o);
    const type = String.fromCharCode(...b.slice(o + 4, o + 8));
    const end = o + 12 + len;
    if (end > b.length) return null;
    if (PNG_KEEP.has(type)) parts.push(b.slice(o, end));
    o = end;
    if (type === "IEND") break;
  }
  return concat(parts);
}

function stripJpeg(b: Uint8Array): Uint8Array | null {
  if (b[0] !== 0xff || b[1] !== 0xd8) return null;
  const parts: Uint8Array[] = [b.slice(0, 2)];
  let o = 2;
  while (o + 4 <= b.length) {
    if (b[o] !== 0xff) return null;
    const marker = b[o + 1];
    if (marker === 0xda) {
      // 開始影像資料：之後全部照抄
      parts.push(b.slice(o));
      return concat(parts);
    }
    const len = (b[o + 2] << 8) | b[o + 3];
    const seg = b.slice(o, o + 2 + len);
    const body = seg.slice(4);
    const keep =
      marker === 0xe0 || // JFIF
      marker === 0xee || // Adobe 色彩轉換
      (marker === 0xe2 && startsWithAscii(body, "ICC_PROFILE\0")) || // 色彩設定檔
      marker < 0xe0 || // DQT、SOF、DHT、DRI 等影像必需嘅段
      (marker > 0xef && marker !== 0xfe); // 唔係 APPn 亦唔係註解
    if (keep) parts.push(seg);
    o += 2 + len;
  }
  return null;
}

function stripGif(b: Uint8Array): Uint8Array | null {
  if (!startsWithAscii(b, "GIF8")) return null;
  const flags = b[10];
  let o = 13 + (flags & 0x80 ? 3 * 2 ** ((flags & 7) + 1) : 0);
  const parts: Uint8Array[] = [b.slice(0, o)];
  const skipSubBlocks = (from: number) => {
    let p = from;
    while (p < b.length && b[p] !== 0) p += b[p] + 1;
    return p + 1;
  };
  while (o < b.length) {
    const id = b[o];
    if (id === 0x3b) {
      parts.push(b.slice(o, o + 1));
      return concat(parts);
    }
    if (id === 0x2c) {
      const lflags = b[o + 9];
      let p = o + 10 + (lflags & 0x80 ? 3 * 2 ** ((lflags & 7) + 1) : 0);
      p = skipSubBlocks(p + 1); // LZW 最小碼長之後係資料
      parts.push(b.slice(o, p));
      o = p;
    } else if (id === 0x21) {
      const label = b[o + 1];
      const end = skipSubBlocks(o + 2);
      const isLoop = label === 0xff && (startsWithAscii(b, "NETSCAPE2.0", o + 3) || startsWithAscii(b, "ANIMEXTS1.0", o + 3));
      // 保留：圖形控制（動畫時間、透明）、純文字、循環播放；清走：註解、其他應用程式擴充（例如 XMP）
      if (label === 0xf9 || label === 0x01 || isLoop) parts.push(b.slice(o, end));
      o = end;
    } else return null;
  }
  return null;
}

function stripBmp(b: Uint8Array): Uint8Array | null {
  if (!startsWithAscii(b, "BM")) return null;
  const size = new DataView(b.buffer, b.byteOffset).getUint32(2, true);
  return b.slice(0, Math.min(size, b.length)); // 檔案尾嘅多餘資料唔要
}

/** 清走圖片 metadata；格式唔啱或者讀唔到就回傳 null（呼叫者會改用佔位圖） */
export function stripImageMetadata(bytes: Uint8Array, kind: Kind): Uint8Array | null {
  switch (kind) {
    case "png":
      return stripPng(bytes);
    case "jpg":
      return stripJpeg(bytes);
    case "gif":
      return stripGif(bytes);
    case "bmp":
      return stripBmp(bytes);
  }
}

// ── 佔位圖 ───────────────────────────────────────────

const CRC_TABLE = Array.from({ length: 256 }, (_, n) => {
  let c = n;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  return c >>> 0;
});

function crc32(data: Uint8Array): number {
  let c = 0xffffffff;
  for (const x of data) c = CRC_TABLE[(c ^ x) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function chunk(type: string, data: Uint8Array): Uint8Array {
  const out = new Uint8Array(12 + data.length);
  const dv = new DataView(out.buffer);
  dv.setUint32(0, data.length);
  out.set(Uint8Array.from(type, (c) => c.charCodeAt(0)), 4);
  out.set(data, 8);
  dv.setUint32(8 + data.length, crc32(out.slice(4, 8 + data.length)));
  return out;
}

async function deflate(data: Uint8Array): Promise<Uint8Array> {
  const stream = new Blob([data as BlobPart]).stream().pipeThrough(new CompressionStream("deflate"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

/** 淺灰色、虛線邊嘅佔位圖，同原圖一樣比例（最長一邊 400px） */
export async function placeholderPng(widthEmu: number, heightEmu: number): Promise<Uint8Array> {
  const ratio = widthEmu > 0 && heightEmu > 0 ? widthEmu / heightEmu : 2;
  const w = Math.max(8, Math.round(ratio >= 1 ? 400 : 400 * ratio));
  const h = Math.max(8, Math.round(ratio >= 1 ? 400 / ratio : 400));
  const raw = new Uint8Array((w + 1) * h);
  for (let y = 0; y < h; y++) {
    raw[y * (w + 1)] = 0; // 每行嘅過濾類型
    for (let x = 0; x < w; x++) {
      const edge = x < 3 || y < 3 || x >= w - 3 || y >= h - 3;
      const dash = Math.floor((x + y) / 12) % 2 === 0;
      raw[y * (w + 1) + 1 + x] = edge ? (dash ? 0x99 : 0xf2) : 0xf2;
    }
  }
  const ihdr = new Uint8Array(13);
  const dv = new DataView(ihdr.buffer);
  dv.setUint32(0, w);
  dv.setUint32(4, h);
  ihdr.set([8, 0, 0, 0, 0], 8); // 8-bit 灰階
  return concat([Uint8Array.from(PNG_SIG), chunk("IHDR", ihdr), chunk("IDAT", await deflate(raw)), chunk("IEND", new Uint8Array())]);
}
