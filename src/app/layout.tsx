import type { Metadata } from "next";
import "./globals.css";
import ThemeToggle from "@/components/ThemeToggle";

export const metadata: Metadata = {
    title: "PO Cutting",
    description: "Madison 88 PO Cutting with AI extraction and Nexgen validation",
};

export default function RootLayout({
    children,
}: Readonly<{
    children: React.ReactNode;
}>) {
    return (
        <html lang="en" suppressHydrationWarning>
            <head>
                <script
                    dangerouslySetInnerHTML={{
                        __html: `
  (function() {
    try {
      const saved = window.localStorage.getItem('theme');
      const systemDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
      const theme = saved === 'light' || saved === 'dark' ? saved : (systemDark ? 'dark' : 'light');
      document.documentElement.classList.remove('light', 'dark');
      document.documentElement.classList.add(theme);
    } catch (e) {
      console.error('Theme init error', e);
    }
  })();
                        `,
                    }}
                />
            </head>
            <body className="font-sans min-h-screen bg-[hsl(var(--background))] text-[hsl(var(--foreground))]">
                <ThemeToggle />
                {children}
            </body>
        </html>
    );
}
