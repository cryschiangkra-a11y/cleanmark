import { analyze, build, CleanMarkError } from "../core/pipeline";
import type { WorkerRequest, WorkerResponse } from "./protocol";

// 所有讀取、檢查、輸出都喺呢個背景執行緒做，畫面唔會卡

let current: { name: string; bytes: Uint8Array } | undefined;

function reply(msg: WorkerResponse, transfer: Transferable[] = []): void {
  (self as unknown as Worker).postMessage(msg, transfer);
}

self.onmessage = async (ev: MessageEvent<WorkerRequest>) => {
  const req = ev.data;
  try {
    if (req.type === "analyze") {
      const bytes = new Uint8Array(req.bytes);
      const result = await analyze(req.name, bytes);
      current = { name: req.name, bytes };
      reply({ id: req.id, ok: true, result });
    } else {
      if (!current) throw new CleanMarkError("corrupt");
      const result = await build(current.name, current.bytes, req.choices);
      reply({ id: req.id, ok: true, result }, [result.bytes.buffer as ArrayBuffer]);
    }
  } catch (e) {
    if (e instanceof CleanMarkError) reply({ id: req.id, ok: false, code: e.code, problems: e.problems });
    else {
      console.error(e);
      reply({ id: req.id, ok: false, code: "unknown", problems: [] });
    }
  }
};
