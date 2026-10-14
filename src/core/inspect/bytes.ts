// 喺 bytes 入面搜尋 ASCII 標記，用嚟認 EXIF / XMP / C2PA

function indexOf(hay: Uint8Array, needle: string): number {
  const n = Array.from(needle, (c) => c.charCodeAt(0));
  outer: for (let i = 0; i <= hay.length - n.length; i++) {
    for (let j = 0; j < n.length; j++) if (hay[i + j] !== n[j]) continue outer;
    return i;
  }
  return -1;
}

export function containsAscii(hay: Uint8Array, needle: string): boolean {
  return indexOf(hay, needle) >= 0;
}

export function hasC2pa(bytes: Uint8Array): boolean {
  // JUMBF 盒子類型 "jumb"、C2PA 標籤 "c2pa"、PNG 嘅 C2PA chunk "caBX"
  return containsAscii(bytes, "c2pa") || containsAscii(bytes, "jumb") || containsAscii(bytes, "caBX");
}

export function imageTags(bytes: Uint8Array): ("exif" | "xmp" | "c2pa")[] {
  const tags: ("exif" | "xmp" | "c2pa")[] = [];
  if (containsAscii(bytes, "Exif\0\0") || containsAscii(bytes, "eXIf")) tags.push("exif");
  if (containsAscii(bytes, "http://ns.adobe.com/xap/1.0/") || containsAscii(bytes, "XML:com.adobe.xmp")) tags.push("xmp");
  if (hasC2pa(bytes)) tags.push("c2pa");
  return tags;
}
