"use client";

import { useEffect, useState } from "react";
import { Moon, Sun } from "lucide-react";

export default function ThemeToggle() {
  const [isDark, setIsDark] = useState(false);

  useEffect(() => {
    try {
      const saved = window.localStorage.getItem("theme");
      const prefersDark = window.matchMedia("(prefers-color-scheme: dark)").matches;
      const initial = saved === "dark" || saved === "light" ? saved : prefersDark ? "dark" : "light";
      setIsDark(initial === "dark");
    } catch {
      setIsDark(false);
    }
  }, []);

  const toggle = () => {
    const next = isDark ? "light" : "dark";
    document.documentElement.classList.remove("light", "dark");
    document.documentElement.classList.add(next);
    localStorage.setItem("theme", next);
    setIsDark(next === "dark");
  };

  const buttonClass = isDark
    ? "fixed right-4 top-4 z-50 inline-flex items-center gap-2 rounded-full border border-white/20 bg-white/10 px-3 py-2 text-xs font-semibold text-white backdrop-blur-md transition hover:bg-white/20 focus:outline-none focus:ring-2 focus:ring-blue-400"
    : "fixed right-4 top-4 z-50 inline-flex items-center gap-2 rounded-full border border-slate-300 bg-white px-3 py-2 text-xs font-semibold text-slate-700 shadow-md transition hover:bg-slate-100 focus:outline-none focus:ring-2 focus:ring-blue-400";

  return (
    <button
      onClick={toggle}
      aria-label="Toggle theme"
      className={buttonClass}
    >
      {isDark ? <Sun className="h-4 w-4" /> : <Moon className="h-4 w-4" />}
      <span>{isDark ? "Light mode" : "Dark mode"}</span>
    </button>
  );
}
