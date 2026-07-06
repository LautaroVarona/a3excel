import type { ExcelExportMetadata } from "@/lib/excel-types";
import { hasA3Metadata } from "@/lib/extract-a3-metadata";

interface ExcelExportHeaderProps {
  metadata: ExcelExportMetadata;
}

export function ExcelExportHeader({ metadata }: ExcelExportHeaderProps) {
  if (!hasA3Metadata(metadata)) return null;

  return (
    <div className="rounded-sm border border-border bg-card px-4 py-3">
      <dl className="grid gap-4 text-xs sm:grid-cols-3">
        {(metadata.companyCode || metadata.companyName) && (
          <div className="space-y-1">
            <dt className="font-medium uppercase tracking-wide text-muted-foreground">
              Empresa
            </dt>
            <dd className="text-sm text-foreground">
              {metadata.companyCode && (
                <span className="font-mono">{metadata.companyCode}</span>
              )}
              {metadata.companyCode && metadata.companyName && (
                <span className="text-muted-foreground"> · </span>
              )}
              {metadata.companyName && <span>{metadata.companyName}</span>}
            </dd>
          </div>
        )}

        {metadata.selection && (
          <div className="space-y-1">
            <dt className="font-medium uppercase tracking-wide text-muted-foreground">
              Selección
            </dt>
            <dd className="text-sm text-foreground">{metadata.selection}</dd>
          </div>
        )}

        {metadata.exportDate && (
          <div className="space-y-1">
            <dt className="font-medium uppercase tracking-wide text-muted-foreground">
              Fecha
            </dt>
            <dd className="text-sm text-foreground">{metadata.exportDate}</dd>
          </div>
        )}
      </dl>
    </div>
  );
}
