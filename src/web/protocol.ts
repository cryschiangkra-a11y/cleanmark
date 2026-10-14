import type { Choices, CmDocument, Report } from "../core/model";
import type { ErrorCode } from "../core/pipeline";
import type { Problem } from "../core/verify/self-check-docx";

export type WorkerRequest =
  | { id: number; type: "analyze"; name: string; bytes: ArrayBuffer }
  | { id: number; type: "build"; choices: Choices };

export type AnalyzeResult = { doc: CmDocument; report: Report };
export type BuildResult = { name: string; bytes: Uint8Array };

export type WorkerResponse =
  | { id: number; ok: true; result: AnalyzeResult | BuildResult }
  | { id: number; ok: false; code: ErrorCode | "unknown"; problems: Problem[] };
