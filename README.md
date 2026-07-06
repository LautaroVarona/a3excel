# A3 Excel

Herramienta web de carga y filtrado de datos para archivos Excel (.XLS / .XLSX).

## Stack

- **Next.js 15** (App Router) + TypeScript
- **Tailwind CSS 4** + Shadcn UI
- **SheetJS (xlsx)** — lectura local en el navegador
- **TanStack Table** — filtrado, ordenamiento y paginación

## Inicio rápido

```bash
npm install
npm run dev
```

Abre [http://localhost:3000](http://localhost:3000).

## Características

- Carga por clic o arrastrar y soltar
- Procesamiento 100% en el cliente (sin subida al servidor)
- Filtrado por columna con inputs individuales
- Ordenamiento interactivo por columnas
- Paginación configurable (25 / 50 / 100 / 250 filas)
- Diseño noir minimalista de alto contraste

## Estructura

```
src/
├── app/
│   ├── layout.tsx      # Layout raíz con tema oscuro
│   ├── page.tsx        # Página principal
│   └── globals.css     # Variables CSS noir
├── components/
│   ├── excel-viewer.tsx  # Componente principal
│   └── ui/               # Componentes Shadcn
└── lib/
    ├── excel.ts          # Parser SheetJS
    └── utils.ts
```
