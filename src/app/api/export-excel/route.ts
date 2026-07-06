import { NextResponse } from "next/server";

import type { ParsedExcel } from "@/lib/excel-types";
import {
  buildEditableExportName,
  exportYmantPreservingBuffer,
} from "@/lib/xls-preserving-export";
import { MAX_FILE_SIZE_BYTES } from "@/lib/excel-types";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";
export const maxDuration = 60;

export async function POST(request: Request) {
  try {
    const formData = await request.formData();
    const file = formData.get("file");
    const payloadField = formData.get("payload");
    const passwordField = formData.get("password");

    if (!(file instanceof File)) {
      return NextResponse.json(
        { error: "No se recibió el archivo original." },
        { status: 400 }
      );
    }

    if (typeof payloadField !== "string") {
      return NextResponse.json(
        { error: "No se recibieron los datos a exportar." },
        { status: 400 }
      );
    }

    if (file.size <= 0 || file.size > MAX_FILE_SIZE_BYTES) {
      return NextResponse.json(
        { error: "El archivo original no es válido." },
        { status: 400 }
      );
    }

    const payload = JSON.parse(payloadField) as Pick<
      ParsedExcel,
      "rows" | "originalRows" | "layout" | "columns" | "sheetName" | "totalRows" | "metadata"
    >;

    const password =
      typeof passwordField === "string" && passwordField.trim()
        ? passwordField.trim()
        : undefined;

    const input = Buffer.from(await file.arrayBuffer());
    const output = exportYmantPreservingBuffer(
      input,
      {
        sheetName: payload.sheetName ?? "Conceptos",
        columns: payload.columns ?? [],
        rows: payload.rows ?? [],
        originalRows: payload.originalRows,
        totalRows: payload.totalRows ?? payload.rows?.length ?? 0,
        metadata: payload.metadata ?? {
          companyCode: null,
          companyName: null,
          selection: null,
          exportDate: null,
          sectionTitle: null,
        },
        layout: payload.layout,
      },
      password
    );

    const fileName = buildEditableExportName(file.name);

    return new NextResponse(new Uint8Array(output), {
      status: 200,
      headers: {
        "Content-Type": "application/vnd.ms-excel",
        "Content-Disposition": `attachment; filename="${fileName}"`,
        "Cache-Control": "no-store",
      },
    });
  } catch (error) {
    const message =
      error instanceof Error
        ? error.message
        : "No se pudo exportar el archivo Excel.";

    return NextResponse.json({ error: message }, { status: 422 });
  }
}
