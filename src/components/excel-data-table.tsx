"use client";

import { useCallback, useMemo, useRef, useState } from "react";
import { AgGridReact } from "ag-grid-react";
import {
  AllCommunityModule,
  ModuleRegistry,
  type ColDef,
  type FilterChangedEvent,
  type GridApi,
  type GridReadyEvent,
} from "ag-grid-community";
import { Download, FileSpreadsheet, Upload } from "lucide-react";

import { ExcelAutoFilter } from "@/components/grid/excel-auto-filter";
import { ExcelExportHeader } from "@/components/excel-export-header";
import { Button } from "@/components/ui/button";
import { useTheme } from "@/components/theme-provider";
import { isA3NativeExportLayout } from "@/lib/is-a3-native-export";
import { type ParsedExcel, formatCellValue } from "@/lib/excel";
import { exportForExcelEditing } from "@/lib/export-for-excel-editing";
import { exportForA3Import } from "@/lib/export-parsed-excel";
import {
  parseEditedA3Rows,
  resolveLayoutFromEditedFile,
} from "@/lib/merge-edited-excel";
import { cn } from "@/lib/utils";

import "ag-grid-community/styles/ag-grid.css";
import "ag-grid-community/styles/ag-theme-alpine.css";

ModuleRegistry.registerModules([AllCommunityModule]);

interface ExcelDataTableProps {
  data: ParsedExcel;
  sourceFileName?: string | null;
  sourceBuffer?: ArrayBuffer | null;
  filePassword?: string;
  onDataUpdated: (data: ParsedExcel) => void;
  onUploadAnother: () => void;
}

export function ExcelDataTable({
  data,
  sourceFileName,
  sourceBuffer,
  filePassword,
  onDataUpdated,
  onUploadAnother,
}: ExcelDataTableProps) {
  const { theme } = useTheme();
  const gridApiRef = useRef<GridApi | null>(null);
  const editedFileInputRef = useRef<HTMLInputElement>(null);
  const [filteredCount, setFilteredCount] = useState(data.totalRows);
  const [isExportingExcel, setIsExportingExcel] = useState(false);
  const [isExportingA3, setIsExportingA3] = useState(false);
  const [isImportingEdited, setIsImportingEdited] = useState(false);
  const [exportError, setExportError] = useState<string | null>(null);
  const [importNotice, setImportNotice] = useState<string | null>(null);

  const isA3Export = isA3NativeExportLayout(data.layout);

  const defaultColDef = useMemo<ColDef>(
    () => ({
      filter: true,
      sortable: true,
      resizable: true,
      editable: true,
      minWidth: 140,
    }),
    []
  );

  const columnDefs = useMemo<ColDef[]>(
    () =>
      data.columns.map((columnId) => ({
        field: columnId,
        headerName: columnId,
        filter: ExcelAutoFilter,
        valueFormatter: (params) =>
          formatCellValue(params.value as ParsedExcel["rows"][number][string]),
        tooltipValueGetter: (params) =>
          formatCellValue(params.value as ParsedExcel["rows"][number][string]),
      })),
    [data.columns]
  );

  const updateFilteredCount = useCallback((api: GridApi) => {
    setFilteredCount(api.getDisplayedRowCount());
  }, []);

  const onGridReady = useCallback(
    (event: GridReadyEvent) => {
      gridApiRef.current = event.api;
      updateFilteredCount(event.api);
    },
    [updateFilteredCount]
  );

  const onFilterChanged = useCallback(
    (event: FilterChangedEvent) => {
      updateFilteredCount(event.api);
    },
    [updateFilteredCount]
  );

  const handleExportForExcel = useCallback(async () => {
    setExportError(null);
    setIsExportingExcel(true);
    try {
      await exportForExcelEditing(data, sourceFileName);
    } catch (err) {
      setExportError(
        err instanceof Error ? err.message : "No se pudo generar el Excel."
      );
    } finally {
      setIsExportingExcel(false);
    }
  }, [data, sourceFileName]);

  const handleExportForA3 = useCallback(async () => {
    setExportError(null);
    setIsExportingA3(true);
    try {
      await exportForA3Import(data, {
        sourceFileName,
        sourceBuffer,
        password: filePassword?.trim() || undefined,
      });
    } catch (err) {
      setExportError(
        err instanceof Error
          ? err.message
          : "No se pudo generar el import para A3."
      );
    } finally {
      setIsExportingA3(false);
    }
  }, [data, filePassword, sourceBuffer, sourceFileName]);

  const handleEditedFileChange = useCallback(
    async (event: React.ChangeEvent<HTMLInputElement>) => {
      const file = event.target.files?.[0];
      event.target.value = "";
      if (!file || !isA3Export || !data.layout) return;

      setExportError(null);
      setImportNotice(null);
      setIsImportingEdited(true);

      try {
        const buffer = await file.arrayBuffer();
        const layout = resolveLayoutFromEditedFile(buffer, data.layout);
        const columns = Object.keys(layout.columnIndices);
        const rows = parseEditedA3Rows(buffer, layout, columns);

        onDataUpdated({
          ...data,
          layout,
          columns,
          rows,
          totalRows: rows.length,
          originalRows: data.originalRows ?? data.rows,
        });

        setImportNotice(
          `Excel editado cargado — ${rows.length.toLocaleString("es-ES")} filas. ` +
            "Ahora podés generar el .XLS para A3."
        );
      } catch (err) {
        setExportError(
          err instanceof Error
            ? err.message
            : "No se pudo leer el Excel editado."
        );
      } finally {
        setIsImportingEdited(false);
      }
    },
    [data, isA3Export, onDataUpdated]
  );

  const exportBusy = isExportingExcel || isExportingA3 || isImportingEdited;

  return (
    <div className="flex flex-col gap-4">
      <ExcelExportHeader metadata={data.metadata} />

      {isA3Export && (
        <div className="rounded-sm border border-border bg-muted/40 px-4 py-3 text-sm text-muted-foreground">
          <p className="font-medium text-foreground">Flujo con Excel de escritorio</p>
          <ol className="mt-2 list-decimal space-y-1 pl-5">
            <li>
              <strong>Descargar para Excel</strong> — .xlsx con filtros, conservando
              la misma disposición del export original (cabecera y metadatos).
            </li>
            <li>Editá y filtrá en Excel. Guardá como .xlsx (no reordenes filas ni columnas).</li>
            <li>
              <strong>Subir Excel editado</strong> — la app fusiona tus cambios con el
              original de A3.
            </li>
            <li>
              <strong>Generar .XLS para A3</strong> — archivo cifrado listo para importar
              (~245 KB).
            </li>
          </ol>
        </div>
      )}

      <div className="flex flex-wrap items-center justify-between gap-3 text-xs text-muted-foreground">
        <div className="flex flex-wrap items-center gap-4">
          <span>
            Hoja: <span className="text-foreground">{data.sheetName}</span>
          </span>
          <span>
            Registros:{" "}
            <span className="text-foreground">
              {filteredCount.toLocaleString("es-ES")}
            </span>
            {filteredCount !== data.totalRows && (
              <> de {data.totalRows.toLocaleString("es-ES")}</>
            )}
          </span>
          <span>
            Columnas:{" "}
            <span className="text-foreground">{data.columns.length}</span>
          </span>
          {isA3Export && (
            <span className="rounded-sm border border-foreground/20 px-2 py-0.5 text-foreground">
              Export nativo A3
            </span>
          )}
        </div>

        <div className="flex flex-wrap items-center gap-2">
          {isA3Export && (
            <>
              <Button
                variant="outline"
                size="sm"
                onClick={() => void handleExportForExcel()}
                disabled={exportBusy}
              >
                <FileSpreadsheet className="h-3.5 w-3.5" />
                {isExportingExcel ? "Generando…" : "Descargar para Excel"}
              </Button>
              <Button
                variant="outline"
                size="sm"
                onClick={() => editedFileInputRef.current?.click()}
                disabled={exportBusy}
              >
                <Upload className="h-3.5 w-3.5" />
                {isImportingEdited ? "Leyendo…" : "Subir Excel editado"}
              </Button>
              <Button
                variant="default"
                size="sm"
                onClick={() => void handleExportForA3()}
                disabled={exportBusy || !sourceBuffer}
              >
                <Download className="h-3.5 w-3.5" />
                {isExportingA3 ? "Generando…" : "Generar .XLS para A3"}
              </Button>
            </>
          )}
          {!isA3Export && (
            <Button
              variant="default"
              size="sm"
              onClick={() => void handleExportForA3()}
              disabled={exportBusy}
            >
              <Download className="h-3.5 w-3.5" />
              Exportar Excel
            </Button>
          )}
          <Button variant="outline" size="sm" onClick={onUploadAnother}>
            Otro archivo
          </Button>
        </div>
      </div>

      <input
        ref={editedFileInputRef}
        type="file"
        accept=".xls,.xlsx,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        className="hidden"
        onChange={(event) => void handleEditedFileChange(event)}
      />

      {importNotice && (
        <div className="rounded-sm border border-foreground/20 bg-muted/50 px-4 py-3 text-sm text-foreground">
          {importNotice}
        </div>
      )}

      {exportError && (
        <div className="rounded-sm border border-destructive/40 bg-destructive/10 px-4 py-3 text-sm text-destructive">
          {exportError}
        </div>
      )}

      <div
        className={cn(
          "overflow-hidden rounded-sm border border-border",
          theme === "dark" ? "ag-theme-alpine-dark" : "ag-theme-alpine"
        )}
        style={{ width: "100%", height: "min(72vh, 900px)" }}
      >
        <AgGridReact
          theme="legacy"
          rowData={data.rows}
          columnDefs={columnDefs}
          defaultColDef={defaultColDef}
          animateRows={false}
          suppressCellFocus
          enableCellTextSelection
          ensureDomOrder
          onGridReady={onGridReady}
          onFilterChanged={onFilterChanged}
          overlayNoRowsTemplate="Sin resultados para los filtros aplicados."
        />
      </div>
    </div>
  );
}
