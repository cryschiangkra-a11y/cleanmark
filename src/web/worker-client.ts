import type { Choices } from "../core/model";
import type { AnalyzeResult, BuildResult, WorkerRequest, WorkerResponse } from "./protocol";

type DistributiveOmit<T, K extends keyof T> = T extends unknown ? Omit<T, K> : never;

export class WorkerError extends Error {
  constructor(
    public code: string,
    public problems: { code: string; detail: string }[],
  ) {
    super(code);
  }
}

type Pending = { resolve: (v: unknown) => void; reject: (e: unknown) => void };

export class CleanMarkWorker {
  private worker = new Worker(new URL("./worker.ts", import.meta.url), { type: "module" });
  private nextId = 1;
  private pending = new Map<number, Pending>();

  constructor() {
    this.worker.onmessage = (ev: MessageEvent<WorkerResponse>) => {
      const msg = ev.data;
      const p = this.pending.get(msg.id);
      if (!p) return;
      this.pending.delete(msg.id);
      if (msg.ok) p.resolve(msg.result);
      else p.reject(new WorkerError(msg.code, msg.problems));
    };
  }

  private send<T>(req: DistributiveOmit<WorkerRequest, "id">, transfer: Transferable[] = []): Promise<T> {
    const id = this.nextId++;
    return new Promise<T>((resolve, reject) => {
      this.pending.set(id, { resolve: resolve as (v: unknown) => void, reject });
      this.worker.postMessage({ ...req, id }, transfer);
    });
  }

  analyze(name: string, bytes: ArrayBuffer): Promise<AnalyzeResult> {
    return this.send({ type: "analyze", name, bytes }, [bytes]);
  }

  build(choices: Choices): Promise<BuildResult> {
    return this.send({ type: "build", choices });
  }
}
