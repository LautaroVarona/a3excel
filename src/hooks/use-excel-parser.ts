"use client";

import { useCallback, useEffect, useRef, useState } from "react";

import type { ParseProgress, ParsedExcel } from "@/lib/excel-types";

export type EngineState = "loading" | "ready" | "error";

type WorkerResponse =
  | { type: "pong" }
  | { type: "progress"; jobId: string; progress: ParseProgress }
  | { type: "result"; jobId: string; data: ParsedExcel }
  | { type: "error"; jobId: string; message: string };

type PendingJob = {
  onProgress: (progress: ParseProgress) => void;
  resolve: (data: ParsedExcel) => void;
  reject: (error: Error) => void;
};

export function useExcelParser() {
  const workerRef = useRef<Worker | null>(null);
  const jobsRef = useRef<Map<string, PendingJob>>(new Map());
  const jobCounterRef = useRef(0);
  const [engineState, setEngineState] = useState<EngineState>("loading");
  const [engineError, setEngineError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    setEngineState("loading");
    setEngineError(null);

    try {
      const worker = new Worker("/workers/excel-parser.js");
      workerRef.current = worker;

      worker.onmessage = (event: MessageEvent<WorkerResponse>) => {
        if (cancelled) return;

        const data = event.data;

        if (data.type === "pong") {
          setEngineState("ready");
          return;
        }

        if (data.type === "progress") {
          jobsRef.current.get(data.jobId)?.onProgress(data.progress);
          return;
        }

        const job = jobsRef.current.get(data.jobId);
        if (!job) return;

        jobsRef.current.delete(data.jobId);

        if (data.type === "result") {
          job.resolve(data.data);
          return;
        }

        job.reject(new Error(data.message));
      };

      worker.onerror = () => {
        if (cancelled) return;
        setEngineState("error");
        setEngineError(
          "No se pudo iniciar el motor de Excel. Recargá la página."
        );
      };

      worker.postMessage({ type: "ping" });
    } catch {
      setEngineState("error");
      setEngineError(
        "Tu navegador no pudo crear el proceso de análisis. Recargá la página."
      );
    }

    return () => {
      cancelled = true;
      workerRef.current?.terminate();
      workerRef.current = null;
      jobsRef.current.clear();
    };
  }, []);

  const parseInWorker = useCallback(
    (
      buffer: ArrayBuffer,
      startTime: number,
      onProgress: (progress: ParseProgress) => void,
      password?: string
    ): Promise<ParsedExcel> => {
      const worker = workerRef.current;
      if (!worker || engineState !== "ready") {
        return Promise.reject(
          new Error("El motor de Excel aún no está listo. Esperá un momento.")
        );
      }

      const jobId = `job-${++jobCounterRef.current}`;

      return new Promise((resolve, reject) => {
        jobsRef.current.set(jobId, { onProgress, resolve, reject });
        worker.postMessage(
          { type: "parse", jobId, buffer, startTime, password },
          [buffer]
        );
      });
    },
    [engineState]
  );

  return { engineState, engineError, parseInWorker };
}
