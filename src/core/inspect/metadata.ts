import type { MetadataItem } from "../model";
import type { OoxmlPackage } from "../ooxml/package";
import { attr, kids, NS, type Element } from "../ooxml/xml";

function leaves(el: Element): Element[] {
  const children = kids(el);
  return children.length === 0 ? [el] : children.flatMap(leaves);
}

/** Word / PowerPoint 共用：docProps 同 customXml */
export async function readOoxmlMetadata(pkg: OoxmlPackage): Promise<MetadataItem[]> {
  const out: MetadataItem[] = [];
  for (const part of ["docProps/core.xml", "docProps/app.xml"]) {
    const root = await pkg.xml(part);
    if (!root) continue;
    for (const leaf of kids(root).flatMap(leaves)) {
      const value = (leaf.textContent ?? "").trim();
      if (value) out.push({ part, field: leaf.nodeName, value });
    }
  }
  const custom = await pkg.xml("docProps/custom.xml");
  if (custom) {
    for (const prop of kids(custom, NS.custom, "property")) {
      out.push({ part: "docProps/custom.xml", field: attr(prop, null, "name") ?? "?", value: (prop.textContent ?? "").trim() });
    }
  }
  for (const path of pkg.paths()) {
    if (/^customXml\/item\d+\.xml$/.test(path)) {
      const text = ((await pkg.text(path)) ?? "").replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
      out.push({ part: path, field: "customXml", value: text.slice(0, 120) });
    }
  }
  return out;
}
