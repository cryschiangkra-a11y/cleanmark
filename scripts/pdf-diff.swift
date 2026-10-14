// 逐頁比較兩個 PDF，用 macOS 內建 PDFKit，唔使額外安裝
// 用法：swift scripts/pdf-diff.swift a.pdf b.pdf [輸出差異圖嘅 folder] [--strict]
//   一般模式：108 dpi，顏色差距大嘅像素先算唔同（忽略抗鋸齒）
//   --strict：216 dpi，任何一個像素唔同都算；另外逐行比較文字、位置、闊度、字型同字體大小
import AppKit
import PDFKit

func render(_ page: PDFPage, scale: CGFloat) -> (pixels: [UInt8], w: Int, h: Int) {
  let box = page.bounds(for: .mediaBox)
  let w = Int(box.width * scale), h = Int(box.height * scale)
  var pixels = [UInt8](repeating: 255, count: w * h * 4)
  let ctx = CGContext(data: &pixels, width: w, height: h, bitsPerComponent: 8, bytesPerRow: w * 4,
                      space: CGColorSpaceCreateDeviceRGB(), bitmapInfo: CGImageAlphaInfo.premultipliedLast.rawValue)!
  ctx.setFillColor(CGColor(red: 1, green: 1, blue: 1, alpha: 1))
  ctx.fill(CGRect(x: 0, y: 0, width: w, height: h))
  ctx.scaleBy(x: scale, y: scale)
  page.draw(with: .mediaBox, to: ctx)
  return (pixels, w, h)
}

func savePNG(_ pixels: [UInt8], _ w: Int, _ h: Int, _ path: String) {
  var p = pixels
  let ctx = CGContext(data: &p, width: w, height: h, bitsPerComponent: 8, bytesPerRow: w * 4,
                      space: CGColorSpaceCreateDeviceRGB(), bitmapInfo: CGImageAlphaInfo.premultipliedLast.rawValue)!
  let rep = NSBitmapImageRep(cgImage: ctx.makeImage()!)
  try? rep.representation(using: .png, properties: [:])!.write(to: URL(fileURLWithPath: path))
}

/** 每行：文字、左邊位置、底部位置、闊度、高度，同埋每段字型 + 大細 */
func lines(_ page: PDFPage) -> [String] {
  guard let all = page.selection(for: page.bounds(for: .mediaBox)) else { return [] }
  return all.selectionsByLine().map { line in
    let b = line.bounds(for: page)
    var fonts: [String] = []
    if let attr = line.attributedString {
      attr.enumerateAttribute(.font, in: NSRange(location: 0, length: attr.length)) { value, range, _ in
        if let f = value as? NSFont { fonts.append("\(f.fontName)@\(String(format: "%.2f", f.pointSize))×\(range.length)") }
      }
    }
    return String(format: "x=%.2f y=%.2f w=%.2f h=%.2f ", b.minX, b.minY, b.width, b.height)
      + "[\(fonts.joined(separator: ","))] " + (line.string ?? "")
  }
}

let args = CommandLine.arguments.filter { $0 != "--strict" }
let strict = CommandLine.arguments.contains("--strict")
guard args.count >= 3, let a = PDFDocument(url: URL(fileURLWithPath: args[1])),
      let b = PDFDocument(url: URL(fileURLWithPath: args[2])) else {
  print("用法：swift pdf-diff.swift a.pdf b.pdf [diff-folder] [--strict]"); exit(2)
}
let outDir = args.count > 3 ? args[3] : nil
let scale: CGFloat = strict ? 3 : 1.5
let threshold = strict ? 0 : 60
print("pages \(a.pageCount) vs \(b.pageCount)\(strict ? "（嚴格模式：216 dpi、零容差、逐行比較）" : "")")
var differing: [Int] = []
var totalLines = 0
var lineDiffs = 0
for i in 0..<min(a.pageCount, b.pageCount) {
  let pa = a.page(at: i)!, pb = b.page(at: i)!
  let ra = render(pa, scale: scale), rb = render(pb, scale: scale)
  if ra.w != rb.w || ra.h != rb.h { print("page \(i + 1): size differs"); differing.append(i + 1); continue }
  var changed = 0
  var diff = [UInt8](repeating: 255, count: ra.pixels.count)
  for px in stride(from: 0, to: ra.pixels.count, by: 4) {
    let d = abs(Int(ra.pixels[px]) - Int(rb.pixels[px])) + abs(Int(ra.pixels[px + 1]) - Int(rb.pixels[px + 1])) + abs(Int(ra.pixels[px + 2]) - Int(rb.pixels[px + 2]))
    if d > threshold { changed += 1; diff[px] = 255; diff[px + 1] = 0; diff[px + 2] = 0 }
    else { let g = UInt8(min(255, 200 + Int(ra.pixels[px]) / 5)); diff[px] = g; diff[px + 1] = g; diff[px + 2] = g }
  }
  var pageBad = changed > 0
  if changed > 0 {
    print(String(format: "page %d: %d pixels differ (%.4f%%)", i + 1, changed, Double(changed) / Double(ra.w * ra.h) * 100))
    if let o = outDir { savePNG(diff, ra.w, ra.h, "\(o)/diff-p\(i + 1).png") }
  }
  if strict {
    let la = lines(pa), lb = lines(pb)
    totalLines += la.count
    if la.count != lb.count { print("page \(i + 1): \(la.count) vs \(lb.count) lines"); pageBad = true }
    for (x, y) in zip(la, lb) where x != y {
      lineDiffs += 1
      pageBad = true
      if lineDiffs <= 10 { print("page \(i + 1):\n  原檔 \(x)\n  新檔 \(y)") }
    }
  }
  if pageBad { differing.append(i + 1) }
}
if strict { print("compared \(totalLines) lines, \(lineDiffs) differ") }
print(differing.isEmpty && a.pageCount == b.pageCount ? "IDENTICAL" : "DIFFERENT pages: \(differing)")
