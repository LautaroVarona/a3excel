"use client";

import { useCallback, useEffect, useMemo, useState } from "react";
import { useGridFilter, type CustomFilterProps } from "ag-grid-react";
import type { IDoesFilterPassParams } from "ag-grid-community";

import { formatCellValue } from "@/lib/excel";
import type { ExcelRow } from "@/lib/excel-types";

export type ExcelFilterModel = {
  text: string;
  textType: "contains" | "equals" | "startsWith" | "endsWith";
  values: string[] | null;
};

const TEXT_OPTIONS: { value: ExcelFilterModel["textType"]; label: string }[] = [
  { value: "contains", label: "Contiene" },
  { value: "equals", label: "Es igual a" },
  { value: "startsWith", label: "Empieza por" },
  { value: "endsWith", label: "Termina en" },
];

function matchesText(
  display: string,
  text: string,
  textType: ExcelFilterModel["textType"]
): boolean {
  const hay = display.toLowerCase();
  const needle = text.toLowerCase().trim();
  if (!needle) return true;

  switch (textType) {
    case "equals":
      return hay === needle;
    case "startsWith":
      return hay.startsWith(needle);
    case "endsWith":
      return hay.endsWith(needle);
    default:
      return hay.includes(needle);
  }
}

export function ExcelAutoFilter({
  model,
  onModelChange,
  getValue,
  api,
}: CustomFilterProps<ExcelRow, unknown, ExcelFilterModel>) {
  const [text, setText] = useState(model?.text ?? "");
  const [textType, setTextType] = useState<ExcelFilterModel["textType"]>(
    model?.textType ?? "contains"
  );
  const [selectedValues, setSelectedValues] = useState<string[] | null>(
    model?.values ?? null
  );
  const [miniFilter, setMiniFilter] = useState("");

  const uniqueValues = useMemo(() => {
    const values = new Set<string>();
    api.forEachNode((node) => {
      if (!node.data) return;
      const raw = getValue(node);
      values.add(formatCellValue(raw as ExcelRow[string]));
    });
    return Array.from(values).sort((a, b) =>
      a.localeCompare(b, "es", { numeric: true, sensitivity: "base" })
    );
  }, [api, getValue]);

  const visibleValues = useMemo(() => {
    const query = miniFilter.trim().toLowerCase();
    if (!query) return uniqueValues;
    return uniqueValues.filter((value) => value.toLowerCase().includes(query));
  }, [miniFilter, uniqueValues]);

  const isAllSelected =
    selectedValues === null ||
    (selectedValues.length === uniqueValues.length &&
      uniqueValues.every((value) => selectedValues.includes(value)));

  const publishModel = useCallback(
    (next: ExcelFilterModel | null) => {
      onModelChange(next);
    },
    [onModelChange]
  );

  useEffect(() => {
    setText(model?.text ?? "");
    setTextType(model?.textType ?? "contains");
    setSelectedValues(model?.values ?? null);
  }, [model]);

  const doesFilterPass = useCallback(
    ({ node }: IDoesFilterPassParams<ExcelRow>) => {
      if (!model) return true;

      const raw = getValue(node);
      const display = formatCellValue(raw as ExcelRow[string]);

      if (model.text.trim() && !matchesText(display, model.text, model.textType)) {
        return false;
      }

      if (model.values !== null && !model.values.includes(display)) {
        return false;
      }

      return true;
    },
    [getValue, model]
  );

  useGridFilter({ doesFilterPass });

  const applyFilters = () => {
    const hasText = text.trim().length > 0;
    const hasValueFilter = selectedValues !== null;

    if (!hasText && !hasValueFilter) {
      publishModel(null);
      return;
    }

    publishModel({
      text,
      textType,
      values: selectedValues,
    });
  };

  const resetFilters = () => {
    setText("");
    setTextType("contains");
    setSelectedValues(null);
    setMiniFilter("");
    publishModel(null);
  };

  const toggleValue = (value: string) => {
    setSelectedValues((current) => {
      const base = current === null ? [...uniqueValues] : [...current];
      const index = base.indexOf(value);
      if (index >= 0) {
        base.splice(index, 1);
      } else {
        base.push(value);
      }
      if (base.length === 0) return [];
      if (base.length === uniqueValues.length) return null;
      return base;
    });
  };

  const toggleSelectAll = () => {
    setSelectedValues(isAllSelected ? [] : null);
  };

  return (
    <div className="flex w-[260px] flex-col gap-3 p-3 text-xs text-foreground">
      <div className="space-y-2">
        <p className="font-medium">Filtro de texto</p>
        <select
          value={textType}
          onChange={(event) =>
            setTextType(event.target.value as ExcelFilterModel["textType"])
          }
          className="h-8 w-full rounded-sm border border-input bg-input px-2"
        >
          {TEXT_OPTIONS.map((option) => (
            <option key={option.value} value={option.value}>
              {option.label}
            </option>
          ))}
        </select>
        <input
          type="text"
          value={text}
          onChange={(event) => setText(event.target.value)}
          placeholder="Buscar texto…"
          className="h-8 w-full rounded-sm border border-input bg-input px-2"
        />
      </div>

      <div className="space-y-2 border-t border-border pt-3">
        <p className="font-medium">Valores únicos</p>
        <input
          type="text"
          value={miniFilter}
          onChange={(event) => setMiniFilter(event.target.value)}
          placeholder="Buscar en la lista…"
          className="h-8 w-full rounded-sm border border-input bg-input px-2"
        />
        <label className="flex items-center gap-2">
          <input
            type="checkbox"
            checked={isAllSelected}
            onChange={toggleSelectAll}
          />
          <span>Seleccionar todo</span>
        </label>
        <div className="max-h-48 space-y-1 overflow-y-auto">
          {visibleValues.map((value) => {
            const checked =
              selectedValues === null || selectedValues.includes(value);
            return (
              <label key={value} className="flex items-center gap-2 truncate">
                <input
                  type="checkbox"
                  checked={checked}
                  onChange={() => toggleValue(value)}
                />
                <span className="truncate" title={value}>
                  {value}
                </span>
              </label>
            );
          })}
          {visibleValues.length === 0 && (
            <p className="text-muted-foreground">Sin coincidencias.</p>
          )}
        </div>
      </div>

      <div className="flex justify-end gap-2 border-t border-border pt-3">
        <button
          type="button"
          onClick={resetFilters}
          className="rounded-sm border border-border px-2 py-1 hover:bg-accent"
        >
          Limpiar
        </button>
        <button
          type="button"
          onClick={applyFilters}
          className="rounded-sm border border-border bg-secondary px-2 py-1 hover:bg-accent"
        >
          Aplicar
        </button>
      </div>
    </div>
  );
}
