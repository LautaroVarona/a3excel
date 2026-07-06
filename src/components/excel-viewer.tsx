"use client";

import { FileSpreadsheet, Upload, X } from "lucide-react";
import { useCallback, useEffect, useRef, useState, startTransition } from "react";
import { flushSync } from "react-dom";

import { ExcelDataTable } from "@/components/excel-data-table";
import { ProcessingPanel } from "@/components/processing-panel";
import { ThemeToggle } from "@/components/theme-toggle";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { useExcelParser } from "@/hooks/use-excel-parser";
import {
  validateExcelFile,
  type ParsedExcel,
  type ParseProgress,
  flushUI,
  parseExcelFileWithProgress,
} from "@/lib/excel";
import { isBrokenA3Reexport } from "@/lib/a3-xls-read-strategies";
import { cn } from "@/lib/utils";

const INITIAL_PROGRESS: ParseProgress = {
  phase: "reading",
  message: "Preparando…",
  percent: 0,
  processed: 0,
  total: 0,
  elapsedMs: 0,
  estimatedRemainingMs: null,
};

export function ExcelViewer() {
  const { engineState, engineError, parseInWorker } = useExcelParser();
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [parsedData, setParsedData] = useState<ParsedExcel | null>(null);
  const [sourceBuffer, setSourceBuffer] = useState<ArrayBuffer | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [progress, setProgress] = useState<ParseProgress>(INITIAL_PROGRESS);
  const [error, setError] = useState<string | null>(null);
  const [filePassword, setFilePassword] = useState("");
  const [isDragging, setIsDragging] = useState(false);

  const updateProgress = useCallback((next: ParseProgress) => {
    setProgress(next);
  }, []);

  useEffect(() => {
    document.body.style.cursor = isProcessing ? "progress" : "";
    return () => {
      document.body.style.cursor = "";
    };
  }, [isProcessing]);

  const beginProcessing = useCallback((name: string) => {
    flushSync(() => {
      setIsProcessing(true);
      setFileName(name);
      setError(null);
      setParsedData(null);
      setSourceBuffer(null);
      setProgress({
        ...INITIAL_PROGRESS,
        message: `Archivo seleccionado: ${name}`,
      });
    });
  }, []);

  const processFile = useCallback(
    async (file: File) => {
      if (engineState !== "ready") {
        setError("El motor de Excel aún se está cargando. Esperá unos segundos.");
        return;
      }

      const validationError = validateExcelFile(file);
      if (validationError) {
        setError(validationError);
        return;
      }

      beginProcessing(file.name);
      await new Promise((resolve) => setTimeout(resolve, 50));

      try {
        const result = await parseExcelFileWithProgress(
          file,
          updateProgress,
          parseInWorker,
          filePassword.trim() || undefined
        );

        if (
          isBrokenA3Reexport(file.name, file.size) &&
          result.data.layout?.kind !== "ymant"
        ) {
          throw new Error(
            "Este .XLS (~33 KB) no es un export válido de A3NOM. " +
              "Importá el archivo original de A3 (≈245 KB). No regrabes en Excel."
          );
        }

        setProgress({
          phase: "complete",
          message: `Preparando tabla con ${result.data.totalRows.toLocaleString("es-ES")} filas…`,
          percent: 100,
          processed: result.data.totalRows,
          total: result.data.totalRows,
          elapsedMs: 0,
          estimatedRemainingMs: null,
        });

        await flushUI();

        startTransition(() => {
          setParsedData({
            ...result.data,
            originalRows: structuredClone(result.data.rows),
          });
          setSourceBuffer(result.sourceBuffer);
          setIsProcessing(false);
        });
      } catch (err) {
        setParsedData(null);
        setSourceBuffer(null);
        setFileName(null);
        setError(
          err instanceof Error
            ? err.message
            : "No se pudo procesar el archivo Excel."
        );
        setIsProcessing(false);
      }
    },
    [beginProcessing, engineState, filePassword, parseInWorker, updateProgress]
  );

  const queueFile = useCallback(
    (file: File | undefined) => {
      if (!file || isProcessing) return;
      window.setTimeout(() => {
        void processFile(file);
      }, 0);
    },
    [isProcessing, processFile]
  );

  const handleFileChange = useCallback(
    (event: React.ChangeEvent<HTMLInputElement>) => {
      const file = event.target.files?.[0];
      event.target.value = "";
      queueFile(file);
    },
    [queueFile]
  );

  const handleDrop = useCallback(
    (event: React.DragEvent<HTMLDivElement>) => {
      event.preventDefault();
      setIsDragging(false);
      queueFile(event.dataTransfer.files?.[0]);
    },
    [queueFile]
  );

  const handleClear = useCallback(() => {
    setParsedData(null);
    setSourceBuffer(null);
    setFileName(null);
    setError(null);
    setProgress(INITIAL_PROGRESS);
  }, []);

  const showUploadZone = !parsedData && !isProcessing;
  const canUpload = engineState === "ready" && !isProcessing;

  const engineLabel =
    engineState === "ready"
      ? "Motor listo"
      : engineState === "loading"
        ? "Inicializando motor…"
        : "Motor con error";

  return (
    <div className="flex min-h-screen flex-col">
      <header className="border-b border-border bg-card/80 backdrop-blur-sm">
        <div className="mx-auto flex max-w-[1600px] items-center justify-between px-6 py-4">
          <div className="flex items-center gap-3">
            <div className="flex h-9 w-9 items-center justify-center rounded-sm border border-border bg-secondary">
              <FileSpreadsheet className="h-4 w-4 text-muted-foreground" />
            </div>
            <div>
              <h1 className="text-sm font-semibold tracking-wide uppercase">
                A3 Excel
              </h1>
              <p className="text-xs text-muted-foreground">
                Análisis local · Archivos a3ERP vía Excel en tu PC
              </p>
            </div>
          </div>

          <div className="flex items-center gap-3">
            <span
              className={cn(
                "hidden rounded-sm border px-2 py-1 text-xs sm:inline",
                engineState === "ready" && "border-foreground/20 text-foreground",
                engineState === "loading" &&
                  "border-border text-muted-foreground",
                engineState === "error" && "border-destructive/40 text-destructive"
              )}
            >
              {engineLabel}
            </span>
            {fileName && !isProcessing && (
              <span className="hidden max-w-[240px] truncate text-xs text-muted-foreground sm:inline">
                {fileName}
              </span>
            )}
            {parsedData && !isProcessing && (
              <Button variant="outline" size="sm" onClick={handleClear}>
                <X className="h-3.5 w-3.5" />
                Limpiar
              </Button>
            )}
            <ThemeToggle />
          </div>
        </div>
      </header>

      <main className="mx-auto flex w-full max-w-[1600px] flex-1 flex-col gap-6 px-6 py-8">
        {engineState === "loading" && (
          <div className="rounded-sm border border-border bg-card px-4 py-3 text-sm text-muted-foreground">
            Inicializando motor de Excel en segundo plano…
          </div>
        )}

        {engineError && (
          <div className="rounded-sm border border-destructive/40 bg-destructive/10 px-4 py-3 text-sm text-destructive">
            {engineError}
          </div>
        )}

        {isProcessing && fileName && (
          <ProcessingPanel fileName={fileName} progress={progress} />
        )}

        {showUploadZone && (
          <div
            role="button"
            tabIndex={canUpload ? 0 : -1}
            aria-disabled={!canUpload}
            onKeyDown={(e) => {
              if (!canUpload) return;
              if (e.key === "Enter" || e.key === " ")
                fileInputRef.current?.click();
            }}
            onDragOver={(e) => {
              if (!canUpload) return;
              e.preventDefault();
              setIsDragging(true);
            }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={handleDrop}
            onClick={() => {
              if (canUpload) fileInputRef.current?.click();
            }}
            className={cn(
              "noir-grid flex flex-col items-center justify-center rounded-sm border border-dashed px-8 py-20 transition-colors",
              canUpload && "cursor-pointer",
              !canUpload && "cursor-wait opacity-80",
              isDragging && canUpload
                ? "border-foreground/40 bg-accent"
                : "border-border bg-card",
              canUpload &&
                "hover:border-muted-foreground/30 hover:bg-accent/50"
            )}
          >
            <input
              ref={fileInputRef}
              type="file"
              accept=".xls,.xlsx,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              onChange={handleFileChange}
              className="hidden"
            />
            <Upload className="mb-4 h-8 w-8 text-muted-foreground" />
            <p className="text-sm font-medium">
              {engineState === "ready"
                ? "Arrastra un archivo Excel o haz clic para seleccionar"
                : "Preparando motor de lectura…"}
            </p>
            <p className="mt-1 text-xs text-muted-foreground">
              Formatos admitidos: .XLS, .XLSX
            </p>
          </div>
        )}

        {showUploadZone && (
          <div className="mx-auto w-full max-w-md space-y-2">
            <label
              htmlFor="excel-password"
              className="text-xs font-medium text-muted-foreground"
            >
              Contraseña de apertura (opcional)
            </label>
            <Input
              id="excel-password"
              type="password"
              placeholder="Solo si el archivo la pide al abrir"
              value={filePassword}
              onChange={(e) => setFilePassword(e.target.value)}
              disabled={!canUpload}
              className="h-9"
            />
            <p className="text-xs text-muted-foreground">
              Los .XLS de a3ERP suelen estar cifrados. Si el navegador no puede
              leerlos, se envían al servidor para descifrarlos (clave XOR de
              a3ERP o la contraseña que ingreses arriba).
            </p>
          </div>
        )}

        {error && (
          <div className="rounded-sm border border-destructive/40 bg-destructive/10 px-4 py-3 text-sm text-destructive">
            {error}
          </div>
        )}

      {parsedData && !isProcessing && (
        <div className="rounded-sm border border-amber-500/40 bg-amber-500/10 px-4 py-3 text-sm text-amber-900 dark:text-amber-100">
          {parsedData.layout?.kind === "ymant" ? (
            <>
              Editá los importes en la tabla y exportá directo a A3.{" "}
              <strong>No abras el .XLS en Excel</strong> — Excel destruye el
              formato interno que A3NOM necesita.
            </>
          ) : (
            <>
              Este archivo no se reconoció como export YMANT de A3. Para
              importar de vuelta en A3NOM usá el .XLS original (~245 KB).
            </>
          )}
        </div>
      )}

      {parsedData && !isProcessing && (
        <ExcelDataTable
            data={parsedData}
            sourceFileName={fileName}
            sourceBuffer={sourceBuffer}
            filePassword={filePassword}
            onUploadAnother={() => fileInputRef.current?.click()}
          />
        )}

        <input
          ref={fileInputRef}
          type="file"
          accept=".xls,.xlsx"
          onChange={handleFileChange}
          className="hidden"
        />
      </main>

      <footer className="border-t border-border py-4 text-center text-xs text-muted-foreground">
        Lectura inicial en el navegador · Los archivos protegidos se procesan en
        el servidor
      </footer>
    </div>
  );
}
