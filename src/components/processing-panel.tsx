"use client";

import { CheckCircle2, Loader2 } from "lucide-react";
import { useEffect, useRef, useState } from "react";

import {
  PHASE_LABELS,
  type ParsePhase,
  type ParseProgress,
  formatDuration,
} from "@/lib/excel";
import { cn } from "@/lib/utils";

const PHASES: ParsePhase[] = [
  "reading",
  "parsing",
  "converting",
  "indexing",
  "complete",
];

const STALL_WARNING_MS = 15_000;

interface ProcessingPanelProps {
  fileName: string;
  progress: ParseProgress;
}

export function ProcessingPanel({ fileName, progress }: ProcessingPanelProps) {
  const currentIndex = PHASES.indexOf(progress.phase);
  const [showStallWarning, setShowStallWarning] = useState(false);
  const lastProgressRef = useRef({
    percent: progress.percent,
    phase: progress.phase,
    updatedAt: Date.now(),
  });

  useEffect(() => {
    const last = lastProgressRef.current;
    if (last.percent !== progress.percent || last.phase !== progress.phase) {
      lastProgressRef.current = {
        percent: progress.percent,
        phase: progress.phase,
        updatedAt: Date.now(),
      };
      setShowStallWarning(false);
      return;
    }

    if (progress.phase === "complete") {
      setShowStallWarning(false);
      return;
    }

    const elapsedSinceUpdate = Date.now() - last.updatedAt;
    if (elapsedSinceUpdate >= STALL_WARNING_MS) {
      setShowStallWarning(true);
      return;
    }

    const timer = setTimeout(() => {
      setShowStallWarning(true);
    }, STALL_WARNING_MS - elapsedSinceUpdate);

    return () => clearTimeout(timer);
  }, [progress.percent, progress.phase]);

  return (
    <div className="flex flex-col gap-6 rounded-sm border border-border bg-card p-6 shadow-sm">
      <div className="flex items-start gap-4">
        <div className="flex h-10 w-10 shrink-0 items-center justify-center rounded-sm border border-border bg-secondary">
          {progress.phase === "complete" ? (
            <CheckCircle2 className="h-5 w-5 text-foreground" />
          ) : (
            <Loader2 className="h-5 w-5 animate-spin text-muted-foreground" />
          )}
        </div>
        <div className="min-w-0 flex-1">
          <p className="text-sm font-medium truncate">{fileName}</p>
          <p className="mt-0.5 text-sm text-muted-foreground">
            {progress.message}
          </p>
        </div>
        <span className="shrink-0 font-mono text-sm tabular-nums text-foreground">
          {progress.percent}%
        </span>
      </div>

      <div className="space-y-2">
        <div className="h-2 overflow-hidden rounded-full bg-secondary">
          <div
            className="h-full rounded-full bg-foreground transition-all duration-300 ease-out"
            style={{ width: `${progress.percent}%` }}
          />
        </div>
        <div className="flex flex-wrap justify-between gap-2 text-xs text-muted-foreground">
          <span>
            {progress.processed > 0 || progress.total > 0 ? (
              <>
                Procesado:{" "}
                <span className="text-foreground tabular-nums">
                  {progress.processed.toLocaleString("es-ES")}
                </span>
                {progress.total > 0 && progress.phase !== "reading" && (
                  <>
                    {" "}
                    /{" "}
                    <span className="text-foreground tabular-nums">
                      {progress.total.toLocaleString("es-ES")}
                    </span>
                  </>
                )}
              </>
            ) : (
              "Iniciando…"
            )}
          </span>
          <span className="tabular-nums">
            {progress.estimatedRemainingMs !== null
              ? `Restante: ${formatDuration(progress.estimatedRemainingMs)}`
              : `Transcurrido: ${formatDuration(progress.elapsedMs)}`}
          </span>
        </div>
      </div>

      {showStallWarning && progress.phase !== "complete" && (
        <p className="rounded-sm border border-border bg-secondary/60 px-3 py-2 text-xs text-muted-foreground">
          Sigue procesando… si tarda mucho, el formato puede no ser compatible.
          Probá guardar el archivo como .xlsx.
        </p>
      )}

      <ol className="grid gap-2 sm:grid-cols-2 lg:grid-cols-5">
        {PHASES.map((phase, index) => {
          const isDone = index < currentIndex || progress.phase === "complete";
          const isActive = index === currentIndex && progress.phase !== "complete";

          return (
            <li
              key={phase}
              className={cn(
                "flex items-center gap-2 rounded-sm border px-3 py-2 text-xs transition-colors",
                isDone && "border-foreground/20 bg-foreground/5 text-foreground",
                isActive && "border-foreground/40 bg-accent text-foreground",
                !isDone && !isActive && "border-border text-muted-foreground"
              )}
            >
              <span
                className={cn(
                  "flex h-5 w-5 shrink-0 items-center justify-center rounded-full text-[10px] font-medium",
                  isDone && "bg-foreground text-background",
                  isActive && "border border-foreground bg-transparent",
                  !isDone && !isActive && "border border-border"
                )}
              >
                {isDone ? "✓" : index + 1}
              </span>
              <span className="leading-tight">{PHASE_LABELS[phase]}</span>
            </li>
          );
        })}
      </ol>
    </div>
  );
}
