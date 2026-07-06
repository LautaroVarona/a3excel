import { randomUUID } from "crypto";
import { spawn } from "child_process";
import { readFile, unlink, writeFile } from "fs/promises";
import { tmpdir } from "os";
import { join } from "path";

function runPowerShell(
  scriptPath: string,
  inputPath: string,
  outputPath: string
): Promise<void> {
  return new Promise((resolve, reject) => {
    const child = spawn(
      "powershell.exe",
      [
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        scriptPath,
        "-InputPath",
        inputPath,
        "-OutputPath",
        outputPath,
      ],
      { windowsHide: true }
    );

    let stderr = "";
    child.stderr.on("data", (chunk: Buffer) => {
      stderr += chunk.toString();
    });

    const timeout = setTimeout(() => {
      child.kill();
      reject(
        new Error(
          "Excel tardó demasiado en abrir el archivo. Cerrá instancias de Excel e intentá de nuevo."
        )
      );
    }, 90_000);

    child.on("error", (error) => {
      clearTimeout(timeout);
      reject(error);
    });

    child.on("close", (code) => {
      clearTimeout(timeout);
      if (code === 0) {
        resolve();
        return;
      }
      reject(
        new Error(
          stderr.trim() ||
            "No se pudo abrir el archivo con Excel. Verificá que Microsoft Excel esté instalado."
        )
      );
    });
  });
}

export function canUseExcelCom(): boolean {
  return process.platform === "win32";
}

export async function convertEncryptedWorkbookViaExcel(
  input: Buffer,
  extension: ".xls" | ".xlsx"
): Promise<Buffer> {
  if (!canUseExcelCom()) {
    throw new Error(
      "Este archivo requiere Microsoft Excel en Windows para descifrarlo."
    );
  }

  const id = randomUUID();
  const inputPath = join(tmpdir(), `a3excel-in-${id}${extension}`);
  const outputPath = join(tmpdir(), `a3excel-out-${id}.xlsx`);
  const scriptPath = join(process.cwd(), "scripts", "excel-com-convert.ps1");

  await writeFile(inputPath, input);

  try {
    await runPowerShell(scriptPath, inputPath, outputPath);
    return await readFile(outputPath);
  } finally {
    await Promise.allSettled([unlink(inputPath), unlink(outputPath)]);
  }
}
