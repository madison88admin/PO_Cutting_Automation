import type { Metadata } from "next";
import { Inter, Outfit } from "next/font/google";
import "./globals.css";
import ThemeToggle from "@/components/ThemeToggle";

const inter = Inter({ subsets: ["latin"], variable: "--font-inter" });
const outfit = Outfit({ subsets: ["latin"], variable: "--font-outfit" });

export const metadata: Metadata = {
    title: "PO Cutting Automation",
    description: "Madison 88 Business Solutions - PO Mass Uploading System",
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
            <body className={`${inter.variable} ${outfit.variable} font-sans min-h-screen bg-[hsl(var(--background))] text-[hsl(var(--foreground))]`}>
                <ThemeToggle />
                {children}
            </body>
        </html>
    );
}
