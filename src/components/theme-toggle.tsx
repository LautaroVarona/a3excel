"use client";

import { Moon, Sun } from "lucide-react";

import { Button } from "@/components/ui/button";
import { useTheme } from "@/components/theme-provider";
import { cn } from "@/lib/utils";

export function ThemeToggle() {
  const { theme, toggleTheme, mounted } = useTheme();

  return (
    <Button
      variant="outline"
      size="icon"
      className="relative h-9 w-9"
      onClick={toggleTheme}
      aria-label={theme === "dark" ? "Activar modo diurno" : "Activar modo nocturno"}
      title={theme === "dark" ? "Modo diurno" : "Modo nocturno"}
    >
      <Sun
        className={cn(
          "h-4 w-4 transition-all",
          mounted && theme === "dark"
            ? "rotate-0 scale-100"
            : "absolute rotate-90 scale-0"
        )}
      />
      <Moon
        className={cn(
          "h-4 w-4 transition-all",
          mounted && theme === "light"
            ? "rotate-0 scale-100"
            : "absolute -rotate-90 scale-0"
        )}
      />
    </Button>
  );
}
