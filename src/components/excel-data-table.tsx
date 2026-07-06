"use client";

import { useCallback, useMemo, useState } from "react";
import { AgGridReact } from "ag-grid-react";
import {
  AllCommunityModule,
  ModuleRegistry,
  type ColDef,
  type FilterChangedEvent,
  type GridApi,
  type GridReadyEvent,
} from "ag-grid-community";
import { Upload } from "lucide-react";

import { ExcelAutoFilter } from "@/components/grid/excel-auto-filter";
import { Button } from "@/components/ui/button";
import { useTheme } from "@/components/theme-provider";
import { type ParsedExcel, formatCellValue } from "@/lib/excel";
import { cn } from "@/lib/utils";

import "ag-grid-community/styles/ag-grid.css";
import "ag-grid-community/styles/ag-theme-alpine.css";

ModuleRegistry.registerModules([AllCommunityModule]);

interface ExcelDataTableProps {
  data: ParsedExcel;
  onUploadAnother: () => void;
}

export function ExcelDataTable({ data, onUploadAnother }: ExcelDataTableProps) {
  const { theme } = useTheme();
  const [filteredCount, setFilteredCount] = useState(data.totalRows);

  const defaultColDef = useMemo<ColDef>(
    () => ({
      filter: true,
      sortable: true,
      resizable: true,
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

  return (
    <div className="flex flex-col gap-4">
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

        <Button variant="outline" size="sm" onClick={onUploadAnother}>
          <Upload className="h-3.5 w-3.5" />
          Otro archivo
        </Button>
      </div>

      <div
        className={cn(
          "overflow-hidden rounded-sm border border-border",
          theme === "dark" ? "ag-theme-alpine-dark" : "ag-theme-alpine"
        )}
        style={{ width: "100%", height: "min(72vh, 900px)" }}
      >
        <AgGridReact
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
