import Link from "next/link";

export default function NotFound() {
  return (
    <div className="flex min-h-screen flex-col items-center justify-center gap-4 px-6 text-center">
      <h1 className="text-2xl font-semibold">Página no encontrada</h1>
      <p className="text-sm text-muted-foreground">
        La ruta solicitada no existe en esta aplicación.
      </p>
      <Link
        href="/"
        className="rounded-sm border border-border px-4 py-2 text-sm hover:bg-accent"
      >
        Volver al inicio
      </Link>
    </div>
  );
}
