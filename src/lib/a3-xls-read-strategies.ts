/** Contraseñas habituales en exports .XLS cifrados de a3ERP / Excel BIFF. */
export const A3_XLS_READ_PASSWORDS = [
  "VelvetSweatshop",
  " ",
  "",
] as const;

export function buildXlsReadPasswordStrategies(
  userPassword?: string
): Array<{ password?: string }> {
  const strategies: Array<{ password?: string }> = [];
  const seen = new Set<string | undefined>();

  const add = (password?: string) => {
    const key = password ?? "";
    if (seen.has(key)) return;
    seen.add(key);
    strategies.push(password ? { password } : {});
  };

  if (userPassword) add(userPassword);
  for (const password of A3_XLS_READ_PASSWORDS) {
    add(password || undefined);
  }

  return strategies;
}

/** Exports .XLS cifrados de A3NOM suelen superar este tamaño. */
export const A3_ENCRYPTED_XLS_MIN_BYTES = 80_000;

export function isLikelyEncryptedA3Xls(
  fileName: string | null | undefined,
  byteLength: number
): boolean {
  return (
    byteLength >= A3_ENCRYPTED_XLS_MIN_BYTES &&
    (fileName?.toLowerCase().endsWith(".xls") ?? false)
  );
}

export function needsA3ServerParse(
  fileName: string | null | undefined,
  byteLength: number,
  hasA3NativeLayout: boolean
): boolean {
  return isLikelyEncryptedA3Xls(fileName, byteLength) && !hasA3NativeLayout;
}

/** Reexports vía Excel/SheetJS suelen quedar en este rango (inválidos para A3). */
export const A3_BROKEN_XLS_MAX_BYTES = 45_000;

export function isBrokenA3Reexport(
  fileName: string | null | undefined,
  byteLength: number
): boolean {
  return (
    (fileName?.toLowerCase().endsWith(".xls") ?? false) &&
    byteLength > 0 &&
    byteLength < A3_BROKEN_XLS_MAX_BYTES
  );
}
