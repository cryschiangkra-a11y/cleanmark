import { cleanParagraphSummary } from "../clean/apply-choices";
import { emptySummary, mergeSummary } from "../clean/hidden-chars";
import { KEEPABLE_IMAGE_TYPES, type CmDocument, type ImageItem, type Report } from "../model";
import type { OoxmlPackage } from "../ooxml/package";
import { allImages, allParagraphs } from "../walk";
import { hasC2pa, imageTags } from "./bytes";
import { readOoxmlMetadata } from "./metadata";
import { findPhrases } from "./phrases";
import { findVisibleDocx } from "./visible-docx";

export async function inspectDocx(pkg: OoxmlPackage, doc: CmDocument): Promise<Report> {
  const c2pa: string[] = [];
  for (const path of pkg.paths()) {
    const bytes = await pkg.bytes(path);
    if (bytes && hasC2pa(bytes)) c2pa.push(path);
  }

  const images: ImageItem[] = allImages(doc).map((img) => ({
    index: img.index,
    context: img.context,
    keepable: KEEPABLE_IMAGE_TYPES.has(img.contentType) && !!doc.media[img.mediaPath],
    tags: doc.media[img.mediaPath] ? imageTags(doc.media[img.mediaPath]) : [],
  }));

  let hidden = emptySummary();
  for (const p of allParagraphs(doc)) hidden = mergeSummary(hidden, cleanParagraphSummary(p));

  return {
    fileKind: "docx",
    visible: await findVisibleDocx(pkg),
    metadata: await readOoxmlMetadata(pkg),
    c2pa,
    hidden,
    phrases: findPhrases(doc),
    images,
  };
}
