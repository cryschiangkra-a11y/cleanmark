import type { Block, CmDocument, PageSetup } from "../../model";
import { OoxmlPackage } from "../../ooxml/package";
import { kid, kids, NS, attr, wnum, type Element } from "../../ooxml/xml";
import { paragraphIds, partName } from "../../ooxml/ids";
import { readBlocks, type ReadState } from "./body";
import { readNumbering } from "./numbering";
import { StyleSheet } from "./styles";

export const MAIN_PART = "word/document.xml";

function readPage(sectPr: Element | undefined): PageSetup {
  const pgSz = kid(sectPr, NS.w, "pgSz");
  const pgMar = kid(sectPr, NS.w, "pgMar");
  return {
    widthTwip: wnum(pgSz, "w") ?? 11906,
    heightTwip: wnum(pgSz, "h") ?? 16838,
    landscape: attr(pgSz, NS.w, "orient") === "landscape",
    margins: {
      top: wnum(pgMar, "top") ?? 1440,
      right: wnum(pgMar, "right") ?? 1440,
      bottom: wnum(pgMar, "bottom") ?? 1440,
      left: wnum(pgMar, "left") ?? 1440,
    },
  };
}

async function readHeaderFooter(
  st: ReadState,
  sectPr: Element | undefined,
  kind: "header" | "footer",
  docRels: Awaited<ReturnType<OoxmlPackage["rels"]>>,
): Promise<Block[]> {
  const refs = sectPr ? kids(sectPr, NS.w, `${kind}Reference`) : [];
  const ref = refs.find((r) => (attr(r, NS.w, "type") ?? "default") === "default");
  const rel = ref ? docRels.get(attr(ref, NS.r, "id") ?? "") : undefined;
  if (!rel) return [];
  const root = await st.pkg.xml(rel.target);
  if (!root) return [];
  return readBlocks(root, st, {
    rels: await st.pkg.rels(rel.target),
    allowGraphics: false,
    part: partName(rel.target),
    ids: paragraphIds(root, partName(rel.target)),
  });
}

async function readFootnotes(st: ReadState): Promise<Record<number, Block[]>> {
  const path = "word/footnotes.xml";
  const root = await st.pkg.xml(path);
  const out: Record<number, Block[]> = {};
  if (!root) return out;
  const rels = await st.pkg.rels(path);
  const ids = paragraphIds(root, "footnotes");
  for (const fn of kids(root, NS.w, "footnote")) {
    if (attr(fn, NS.w, "type")) continue; // separator 等
    const id = wnum(fn, "id");
    if (id === undefined) continue;
    out[id] = await readBlocks(fn, st, { rels, allowGraphics: false, part: "footnotes", ids });
  }
  return out;
}

export async function readDocx(bytes: Uint8Array): Promise<{ doc: CmDocument; pkg: OoxmlPackage }> {
  const pkg = await OoxmlPackage.open(bytes);
  const root = await pkg.xml(MAIN_PART);
  const body = root ? kid(root, NS.w, "body") : undefined;
  if (!body) throw new Error("word/document.xml 冇 <w:body>");

  const st: ReadState = {
    pkg,
    styles: new StyleSheet(await pkg.xml("word/styles.xml")),
    media: {},
    imageCount: 0,
    paraCount: {},
    lastHeading: "",
    lastText: "",
  };
  const docRels = await pkg.rels(MAIN_PART);
  const blocks = await readBlocks(body, st, {
    rels: docRels,
    allowGraphics: true,
    part: "document",
    ids: paragraphIds(root!, "document"),
  });
  const sectPr = kid(body, NS.w, "sectPr");

  const doc: CmDocument = {
    kind: "docx",
    page: readPage(sectPr),
    body: blocks,
    header: await readHeaderFooter(st, sectPr, "header", docRels),
    footer: await readHeaderFooter(st, sectPr, "footer", docRels),
    footnotes: await readFootnotes(st),
    lists: readNumbering(await pkg.xml("word/numbering.xml")),
    media: st.media,
  };
  return { doc, pkg };
}
