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
import { Download, Upload } from "lucide-react";

import { ExcelAutoFilter } from "@/components/grid/excel-auto-filter";
import { ExcelExportHeader } from "@/components/excel-export-header";
import { Button } from "@/components/ui/button";
import { useTheme } from "@/components/theme-provider";
import { type ParsedExcel, formatCellValue } from "@/lib/excel";
import { exportParsedExcelToFile } from "@/lib/export-parsed-excel";
import { cn } from "@/lib/utils";

import "ag-grid-community/styles/ag-grid.css";
import "ag-grid-community/styles/ag-theme-alpine.css";

ModuleRegistry.registerModules([AllCommunityModule]);

interface ExcelDataTableProps {
  data: ParsedExcel;
  sourceFileName?: string | null;
  sourceBuffer?: ArrayBuffer | null;
  filePassword?: string;
  onUploadAnother: () => void;
}

export function ExcelDataTable({
  data,
  sourceFileName,
  sourceBuffer,
  filePassword,
  onUploadAnother,
}: ExcelDataTableProps) {
  const { theme } = useTheme();
  const gridApiRef = useRef<GridApi | null>(null);
  const [filteredCount, setFilteredCount] = useState(data.totalRows);
  const [isExporting, setIsExporting] = useState(false);
  const [exportError, setExportError] = useState<string | null>(null);

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

  const collectRowsForExport = useCallback((): ParsedExcel["rows"] => {
    const api = gridApiRef.current;
    if (!api) return data.rows;

    const rows: ParsedExcel["rows"] = [];
    api.forEachNodeAfterFilterAndSort((node) => {
      if (node.data) {
        rows.push(node.data as ParsedExcel["rows"][number]);
      }
    });
    return rows.length > 0 ? rows : data.rows;
  }, [data.rows]);

  const handleExport = useCallback(async () => {
    setExportError(null);
    setIsExporting(true);
    try {
      await exportParsedExcelToFile(
        {
          ...data,
          rows: collectRowsForExport(),
          originalRows: data.originalRows ?? data.rows,
        },
        {
          sourceFileName,
          sourceBuffer,
          password: filePassword?.trim() || undefined,
        }
      );
    } catch (err) {
      setExportError(
        err instanceof Error
          ? err.message
          : "No se pudo generar el archivo Excel."
      );
    } finally {
      setIsExporting(false);
    }
  }, [collectRowsForExport, data, filePassword, sourceBuffer, sourceFileName]);

  return (
    <div className="flex flex-col gap-4">
      <ExcelExportHeader metadata={data.metadata} />

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
        </div>

        <div className="flex flex-wrap items-center gap-2">
          <Button
            variant="default"
            size="sm"
            onClick={() => void handleExport()}
            disabled={isExporting}
          >
            <Download className="h-3.5 w-3.5" />
            {isExporting ? "Exportando…" : "Exportar Excel"}
          </Button>
          <Button variant="outline" size="sm" onClick={onUploadAnother}>
            <Upload className="h-3.5 w-3.5" />
            Otro archivo
          </Button>
        </div>
      </div>

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
