"use client";

import {
  type ColumnDef,
  type ColumnFiltersState,
  type PaginationState,
  type SortingState,
  flexRender,
  getCoreRowModel,
  getFilteredRowModel,
  getPaginationRowModel,
  getSortedRowModel,
  useReactTable,
} from "@tanstack/react-table";
import {
  ArrowDown,
  ArrowUp,
  ArrowUpDown,
  ChevronLeft,
  ChevronRight,
  ChevronsLeft,
  ChevronsRight,
  FileSpreadsheet,
  Upload,
  X,
} from "lucide-react";
import { useCallback, useMemo, useRef, useState } from "react";

import { ProcessingPanel } from "@/components/processing-panel";
import { ThemeToggle } from "@/components/theme-toggle";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import {
  type ExcelRow,
  type ParseProgress,
  type ParsedExcel,
  formatCellValue,
  parseExcelFileWithProgress,
} from "@/lib/excel";
import { cn } from "@/lib/utils";

const PAGE_SIZE_OPTIONS = [25, 50, 100, 250] as const;

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
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [parsedData, setParsedData] = useState<ParsedExcel | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [progress, setProgress] = useState<ParseProgress>(INITIAL_PROGRESS);
  const [error, setError] = useState<string | null>(null);
  const [sorting, setSorting] = useState<SortingState>([]);
  const [columnFilters, setColumnFilters] = useState<ColumnFiltersState>([]);
  const [pagination, setPagination] = useState<PaginationState>({
    pageIndex: 0,
    pageSize: 50,
  });
  const [isDragging, setIsDragging] = useState(false);

  const columns = useMemo<ColumnDef<ExcelRow>[]>(() => {
    if (!parsedData) return [];

    return parsedData.columns.map((columnId) => ({
      accessorKey: columnId,
      header: columnId,
      cell: ({ getValue }) => formatCellValue(getValue() as ExcelRow[string]),
      filterFn: "includesString",
      sortingFn: "alphanumeric",
    }));
  }, [parsedData]);

  const table = useReactTable({
    data: parsedData?.rows ?? [],
    columns,
    state: { sorting, columnFilters, pagination },
    onSortingChange: setSorting,
    onColumnFiltersChange: setColumnFilters,
    onPaginationChange: setPagination,
    getCoreRowModel: getCoreRowModel(),
    getFilteredRowModel: getFilteredRowModel(),
    getSortedRowModel: getSortedRowModel(),
    getPaginationRowModel: getPaginationRowModel(),
  });

  const processFile = useCallback(async (file: File) => {
    const validExtensions = [".xls", ".xlsx"];
    const extension = file.name.slice(file.name.lastIndexOf(".")).toLowerCase();

    if (!validExtensions.includes(extension)) {
      setError("Formato no válido. Solo se admiten archivos .XLS y .XLSX.");
      return;
    }

    setIsProcessing(true);
    setFileName(file.name);
    setError(null);
    setParsedData(null);
    setProgress(INITIAL_PROGRESS);

    try {
      const result = await parseExcelFileWithProgress(file, setProgress);
      setParsedData(result);
      setSorting([]);
      setColumnFilters([]);
      setPagination((prev) => ({ ...prev, pageIndex: 0 }));
    } catch (err) {
      setParsedData(null);
      setFileName(null);
      setError(
        err instanceof Error
          ? err.message
          : "No se pudo procesar el archivo Excel."
      );
    } finally {
      setIsProcessing(false);
    }
  }, []);

  const handleFileChange = useCallback(
    (event: React.ChangeEvent<HTMLInputElement>) => {
      const file = event.target.files?.[0];
      if (file) void processFile(file);
      event.target.value = "";
    },
    [processFile]
  );

  const handleDrop = useCallback(
    (event: React.DragEvent<HTMLDivElement>) => {
      event.preventDefault();
      setIsDragging(false);
      if (isProcessing) return;
      const file = event.dataTransfer.files?.[0];
      if (file) void processFile(file);
    },
    [isProcessing, processFile]
  );

  const handleClear = useCallback(() => {
    setParsedData(null);
    setFileName(null);
    setError(null);
    setProgress(INITIAL_PROGRESS);
    setSorting([]);
    setColumnFilters([]);
    setPagination({ pageIndex: 0, pageSize: 50 });
  }, []);

  const filteredCount = table.getFilteredRowModel().rows.length;
  const totalCount = parsedData?.totalRows ?? 0;
  const showUploadZone = !parsedData && !isProcessing;

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
                Análisis local · Sin envío al servidor
              </p>
            </div>
          </div>

          <div className="flex items-center gap-3">
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
        {isProcessing && fileName && (
          <ProcessingPanel fileName={fileName} progress={progress} />
        )}

        {showUploadZone && (
          <div
            role="button"
            tabIndex={0}
            onKeyDown={(e) => {
              if (e.key === "Enter" || e.key === " ")
                fileInputRef.current?.click();
            }}
            onDragOver={(e) => {
              e.preventDefault();
              setIsDragging(true);
            }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={handleDrop}
            onClick={() => fileInputRef.current?.click()}
            className={cn(
              "noir-grid flex cursor-pointer flex-col items-center justify-center rounded-sm border border-dashed px-8 py-20 transition-colors",
              isDragging
                ? "border-foreground/40 bg-accent"
                : "border-border bg-card hover:border-muted-foreground/30 hover:bg-accent/50"
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
              Arrastra un archivo Excel o haz clic para seleccionar
            </p>
            <p className="mt-1 text-xs text-muted-foreground">
              Formatos admitidos: .XLS, .XLSX
            </p>
          </div>
        )}

        {error && (
          <div className="rounded-sm border border-destructive/40 bg-destructive/10 px-4 py-3 text-sm text-destructive">
            {error}
          </div>
        )}

        {parsedData && !isProcessing && (
          <div className="flex flex-col gap-4">
            <div className="flex flex-wrap items-center justify-between gap-3 text-xs text-muted-foreground">
              <div className="flex flex-wrap items-center gap-4">
                <span>
                  Hoja:{" "}
                  <span className="text-foreground">{parsedData.sheetName}</span>
                </span>
                <span>
                  Registros:{" "}
                  <span className="text-foreground">
                    {filteredCount.toLocaleString("es-ES")}
                  </span>
                  {filteredCount !== totalCount && (
                    <> de {totalCount.toLocaleString("es-ES")}</>
                  )}
                </span>
                <span>
                  Columnas:{" "}
                  <span className="text-foreground">
                    {parsedData.columns.length}
                  </span>
                </span>
              </div>

              <Button
                variant="outline"
                size="sm"
                onClick={() => fileInputRef.current?.click()}
              >
                <Upload className="h-3.5 w-3.5" />
                Otro archivo
              </Button>
              <input
                ref={fileInputRef}
                type="file"
                accept=".xls,.xlsx"
                onChange={handleFileChange}
                className="hidden"
              />
            </div>

            <div className="overflow-hidden rounded-sm border border-border bg-card">
              <div className="overflow-x-auto">
                <Table>
                  <TableHeader>
                    {table.getHeaderGroups().map((headerGroup) => (
                      <TableRow
                        key={headerGroup.id}
                        className="border-border hover:bg-transparent"
                      >
                        {headerGroup.headers.map((header) => (
                          <TableHead key={header.id} className="min-w-[140px]">
                            {header.isPlaceholder ? null : (
                              <div className="flex flex-col gap-2 py-1">
                                <button
                                  type="button"
                                  className="flex items-center gap-1.5 text-left hover:text-foreground"
                                  onClick={header.column.getToggleSortingHandler()}
                                >
                                  {flexRender(
                                    header.column.columnDef.header,
                                    header.getContext()
                                  )}
                                  {header.column.getIsSorted() === "asc" ? (
                                    <ArrowUp className="h-3.5 w-3.5 text-foreground" />
                                  ) : header.column.getIsSorted() === "desc" ? (
                                    <ArrowDown className="h-3.5 w-3.5 text-foreground" />
                                  ) : (
                                    <ArrowUpDown className="h-3.5 w-3.5 opacity-40" />
                                  )}
                                </button>
                                <Input
                                  placeholder="Filtrar..."
                                  value={
                                    (header.column.getFilterValue() as string) ??
                                    ""
                                  }
                                  onChange={(e) =>
                                    header.column.setFilterValue(e.target.value)
                                  }
                                  className="h-7 text-xs"
                                />
                              </div>
                            )}
                          </TableHead>
                        ))}
                      </TableRow>
                    ))}
                  </TableHeader>
                  <TableBody>
                    {table.getRowModel().rows.length > 0 ? (
                      table.getRowModel().rows.map((row) => (
                        <TableRow key={row.id}>
                          {row.getVisibleCells().map((cell) => (
                            <TableCell
                              key={cell.id}
                              className="max-w-[280px] truncate font-mono text-xs"
                              title={String(cell.getValue() ?? "")}
                            >
                              {flexRender(
                                cell.column.columnDef.cell,
                                cell.getContext()
                              )}
                            </TableCell>
                          ))}
                        </TableRow>
                      ))
                    ) : (
                      <TableRow>
                        <TableCell
                          colSpan={columns.length}
                          className="h-24 text-center text-muted-foreground"
                        >
                          Sin resultados para los filtros aplicados.
                        </TableCell>
                      </TableRow>
                    )}
                  </TableBody>
                </Table>
              </div>

              <div className="flex flex-wrap items-center justify-between gap-4 border-t border-border px-4 py-3">
                <div className="flex items-center gap-2 text-xs text-muted-foreground">
                  <span>Filas por página</span>
                  <select
                    value={table.getState().pagination.pageSize}
                    onChange={(e) => table.setPageSize(Number(e.target.value))}
                    className="h-8 rounded-sm border border-input bg-input px-2 text-xs text-foreground"
                  >
                    {PAGE_SIZE_OPTIONS.map((size) => (
                      <option key={size} value={size}>
                        {size}
                      </option>
                    ))}
                  </select>
                </div>

                <div className="flex items-center gap-1 text-xs text-muted-foreground">
                  Página{" "}
                  <span className="text-foreground">
                    {table.getState().pagination.pageIndex + 1}
                  </span>{" "}
                  de{" "}
                  <span className="text-foreground">
                    {table.getPageCount().toLocaleString("es-ES")}
                  </span>
                </div>

                <div className="flex items-center gap-1">
                  <Button
                    variant="outline"
                    size="icon"
                    className="h-8 w-8"
                    onClick={() => table.setPageIndex(0)}
                    disabled={!table.getCanPreviousPage()}
                  >
                    <ChevronsLeft className="h-4 w-4" />
                  </Button>
                  <Button
                    variant="outline"
                    size="icon"
                    className="h-8 w-8"
                    onClick={() => table.previousPage()}
                    disabled={!table.getCanPreviousPage()}
                  >
                    <ChevronLeft className="h-4 w-4" />
                  </Button>
                  <Button
                    variant="outline"
                    size="icon"
                    className="h-8 w-8"
                    onClick={() => table.nextPage()}
                    disabled={!table.getCanNextPage()}
                  >
                    <ChevronRight className="h-4 w-4" />
                  </Button>
                  <Button
                    variant="outline"
                    size="icon"
                    className="h-8 w-8"
                    onClick={() => table.setPageIndex(table.getPageCount() - 1)}
                    disabled={!table.getCanNextPage()}
                  >
                    <ChevronsRight className="h-4 w-4" />
                  </Button>
                </div>
              </div>
            </div>
          </div>
        )}
      </main>

      <footer className="border-t border-border py-4 text-center text-xs text-muted-foreground">
        Procesamiento 100% en el navegador · Los datos nunca salen de tu equipo
      </footer>
    </div>
  );
}
