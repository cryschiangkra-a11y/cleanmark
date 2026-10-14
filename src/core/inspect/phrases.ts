import type { CmDocument, PhraseItem } from "../model";
import { allParagraphs, paragraphText } from "../walk";

// 可疑字句：中英文，唔分大細楷
export const PHRASE_PATTERNS: RegExp[] = [
  /由\s*AI\s*(?:協助|輔助)?\s*(?:生成|產生|撰寫|製作)/i,
  /AI\s*(?:協助|輔助)/i,
  /人工智能\s*(?:生成|產生|協助)/,
  /人工智慧\s*(?:生成|產生|協助)/,
  /generated\s+(?:by|with|using)\b/i,
  /\bAI[-\s]generated\b/i,
  /\bwritten\s+(?:by|with)\s+AI\b/i,
  /\b(?:ChatGPT|OpenAI|Claude|Anthropic|Gemini|Bard|Copilot|DeepSeek|Perplexity|Grok|Llama)\b/i,
];

export function findPhrases(doc: CmDocument): PhraseItem[] {
  const out: PhraseItem[] = [];
  for (const p of allParagraphs(doc)) {
    const text = paragraphText(p);
    for (const re of PHRASE_PATTERNS) {
      const m = re.exec(text);
      if (m) {
        out.push({ paragraphId: p.id, text: text.length > 160 ? text.slice(0, 160) + "…" : text, match: m[0] });
        break;
      }
    }
  }
  return out;
}
