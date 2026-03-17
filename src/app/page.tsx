"use client";

import { useEffect, useState } from "react";
import Workflow from "@/components/Workflow";
import AdminPanel from "@/components/AdminPanel";
import TemplateReference from "@/components/TemplateReference";
import { Settings, ArrowLeft, Shield, Globe, Zap, Cpu, LayoutDashboard, Database, BookOpen } from "lucide-react";
import { motion, AnimatePresence } from "framer-motion";

export default function Home() {
    const [view, setView] = useState<"WORKFLOW" | "ADMIN" | "REFERENCE">("WORKFLOW");
    const [healthState, setHealthState] = useState<"checking" | "ok" | "down">("checking");
    const [isAdminAuthenticated, setIsAdminAuthenticated] = useState(false);
    const [showAdminLogin, setShowAdminLogin] = useState(false);
    const [adminPassword, setAdminPassword] = useState("");
    const [adminLoginError, setAdminLoginError] = useState("");
    const [isAdminLoginLoading, setIsAdminLoginLoading] = useState(false);

    const checkSystemHealth = async () => {
        setHealthState("checking");
        try {
            const response = await fetch("/api/health", { cache: "no-store" });
            if (!response.ok) {
                setHealthState("down");
                return;
            }

            const payload = await response.json();
            setHealthState(payload?.status === "ok" ? "ok" : "down");
        } catch {
            setHealthState("down");
        }
    };

    useEffect(() => {
        checkSystemHealth();
        const timer = window.setInterval(checkSystemHealth, 30000);
        return () => window.clearInterval(timer);
    }, []);

    useEffect(() => {
        const checkAdminSession = async () => {
            try {
                const response = await fetch("/api/admin/session", { cache: "no-store" });
                if (!response.ok) return;
                const payload = await response.json();
                setIsAdminAuthenticated(Boolean(payload?.authenticated));
            } catch {
                setIsAdminAuthenticated(false);
            }
        };

        checkAdminSession();
    }, []);

    const attemptAdminLogin = async () => {
        setIsAdminLoginLoading(true);
        setAdminLoginError("");
        try {
            const response = await fetch("/api/admin/session", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ password: adminPassword }),
            });

            if (!response.ok) {
                setAdminLoginError("Invalid admin password.");
                return;
            }

            setIsAdminAuthenticated(true);
            setShowAdminLogin(false);
            setAdminPassword("");
            setView("ADMIN");
        } catch {
            setAdminLoginError("Unable to sign in. Please try again.");
        } finally {
            setIsAdminLoginLoading(false);
        }
    };

    const logoutAdmin = async () => {
        try {
            await fetch("/api/admin/session", { method: "DELETE" });
        } finally {
            setIsAdminAuthenticated(false);
            if (view === "ADMIN") {
                setView("WORKFLOW");
            }
        }
    };

    const footerLinks = [
        {
            label: "Documentation",
            onClick: () => {
                setView("REFERENCE");
                window.scrollTo({ top: 0, behavior: "smooth" });
            },
        },
        {
            label: "Technical Support",
            onClick: () => {
                window.open("https://m88itsm.netlify.app/login", "_blank", "noopener,noreferrer");
            },
        },
        {
            label: "System Health",
            onClick: checkSystemHealth,
        },
    ];

    const healthStyle = {
        checking: {
            dot: "bg-amber-400",
            text: "Checking",
            textClass: "text-amber-300",
        },
        ok: {
            dot: "bg-emerald-400",
            text: "Operational",
            textClass: "text-emerald-300",
        },
        down: {
            dot: "bg-rose-400",
            text: "Degraded",
            textClass: "text-rose-300",
        },
    }[healthState];

    return (
        <main className="min-h-screen relative flex flex-col items-center px-6 py-12 md:py-20 overflow-x-hidden selection:bg-blue-500/30 bg-[hsl(var(--background))] text-[hsl(var(--foreground))]">
            {/* Ultra-Premium Ambient Background */}
            <div className="fixed inset-0 pointer-events-none -z-10 overflow-hidden bg-[hsl(var(--background))]">
                <div className="absolute top-[-10%] left-[-10%] w-[60vw] h-[60vw] bg-blue-600/10 rounded-full blur-[180px] animate-pulse-glow" />
                <div className="absolute bottom-[10%] right-[-5%] w-[50vw] h-[50vw] bg-indigo-600/10 rounded-full blur-[150px] animate-float" />
                <div className="absolute top-[40%] left-[70%] w-[30vw] h-[30vw] bg-sky-500/5 rounded-full blur-[120px]" />

                {/* Noise and Texture */}
                <div className="absolute inset-0 opacity-[0.1] mix-blend-overlay pointer-events-none bg-[url('https://grainy-gradients.vercel.app/noise.svg')]" />
                <div className="absolute inset-0 bg-gradient-to-b from-transparent via-black/20 to-[rgba(0,0,0,0.4)] pointer-events-none" />
            </div>

            {/* Top Navigation Bar */}
            <nav className="w-full max-w-7xl flex flex-col md:flex-row justify-between items-center gap-10 mb-24 relative z-50">
                <motion.div
                    initial={{ opacity: 0, x: -30 }}
                    animate={{ opacity: 1, x: 0 }}
                    transition={{ duration: 0.8, ease: "easeOut" }}
                    className="space-y-6 text-center md:text-left"
                >
                    <div className="inline-flex items-center gap-4 px-5 py-2.5 rounded-2xl bg-[hsl(var(--card)/0.6)] border border-[hsl(var(--border)/0.2)] backdrop-blur-xl shadow-2xl group">
                        <div className="flex h-2.5 w-2.5 relative">
                            <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-blue-400 opacity-75"></span>
                            <span className="relative inline-flex rounded-full h-2.5 w-2.5 bg-blue-500 shadow-[0_0_10px_rgba(59,130,246,0.8)]"></span>
                        </div>
                        <div className="flex items-center gap-3 text-[10px] tracking-[0.4em] font-black uppercase">
                            <span className="text-[hsl(var(--foreground)/0.9)]">Madison 88 Solutions</span>
                            <span className="w-px h-3 bg-[hsl(var(--foreground)/0.15)]" />
                            <span className="text-blue-400">System v2.0</span>
                        </div>
                    </div>

                    <h1 className="text-7xl md:text-9xl font-black tracking-tighter leading-[0.8] text-[hsl(var(--foreground))]">
                        PO CUTTING<br />
                        <span className="bg-clip-text text-transparent bg-gradient-to-r from-blue-500 via-sky-400 to-indigo-500">
                            AUTOMATION
                        </span>
                    </h1>
                </motion.div>

                <motion.div
                    initial={{ opacity: 0, x: 30 }}
                    animate={{ opacity: 1, x: 0 }}
                    transition={{ duration: 0.8, ease: "easeOut" }}
                    className="flex flex-col items-center md:items-end gap-8"
                >
                    <div className="flex gap-4">
                        <button
                            type="button"
                            onClick={() => {
                                if (isAdminAuthenticated) {
                                    logoutAdmin();
                                    return;
                                }
                                setShowAdminLogin(true);
                            }}
                            className={`group relative flex items-center gap-4 px-6 py-5 border rounded-2xl text-[10px] font-black uppercase tracking-[0.3em] transition-all shadow-2xl overflow-hidden ${isAdminAuthenticated
                                    ? "bg-emerald-500/10 border-emerald-500/40 text-emerald-200"
                                    : "bg-[hsl(var(--card))] border-[hsl(var(--border))] text-[hsl(var(--muted))] hover:text-[hsl(var(--foreground))] hover:bg-[rgba(255,255,255,0.08)]"
                                }`}
                        >
                            <Shield className="w-4 h-4" />
                            <span>{isAdminAuthenticated ? "Admin Logout" : "Admin Login"}</span>
                        </button>

                        <button
                            onClick={() => setView("REFERENCE")}
                            className={`group relative flex items-center gap-4 px-6 py-5 border rounded-2xl text-[10px] font-black uppercase tracking-[0.3em] transition-all shadow-2xl overflow-hidden ${view === "REFERENCE"
                                    ? "bg-blue-500/10 border-blue-500/50 text-[hsl(var(--foreground))]"
                                    : "bg-[hsl(var(--card))] border-[hsl(var(--border))] text-[hsl(var(--muted))] hover:text-[hsl(var(--foreground))] hover:bg-[rgba(255,255,255,0.08)]"
                                }`}
                        >
                            <BookOpen className={`w-4 h-4 transition-transform ${view === "REFERENCE" ? "scale-110" : "group-hover:scale-110"}`} />
                            <span>System Reference</span>
                        </button>

                        <button
                            onClick={() => {
                                if (view === "ADMIN") {
                                    setView("WORKFLOW");
                                    return;
                                }

                                if (isAdminAuthenticated) {
                                    setView("ADMIN");
                                    return;
                                }

                                setShowAdminLogin(true);
                            }}
                            className={`group relative flex items-center gap-4 px-10 py-5 border rounded-2xl text-[10px] font-black uppercase tracking-[0.4em] transition-all shadow-2xl overflow-hidden ${view === "ADMIN"
                                    ? "bg-blue-500/10 border-blue-500/50 text-[hsl(var(--foreground))]"
                                    : "bg-[hsl(var(--card))] border-[hsl(var(--border))] text-[hsl(var(--muted))] hover:text-[hsl(var(--foreground))] hover:bg-[rgba(255,255,255,0.12)]"
                                }`}
                        >
                            <div className="absolute inset-0 bg-gradient-to-r from-blue-500/10 to-transparent opacity-0 group-hover:opacity-100 transition-opacity" />
                            {view === "WORKFLOW" || view === "REFERENCE" ? (
                                <>
                                    <LayoutDashboard className="w-4 h-4 group-hover:rotate-12 transition-transform duration-500 text-blue-500" />
                                    <span>Management Console</span>
                                </>
                            ) : (
                                <>
                                    <ArrowLeft className="w-4 h-4 group-hover:-translate-x-2 transition-transform text-blue-500" />
                                    <span>Return to Workflow</span>
                                </>
                            )}
                        </button>
                    </div>

                    <div className="flex gap-10">
                        {[
                            { icon: Shield, label: "Secure" },
                            { icon: Zap, label: "Turbo" },
                            { icon: Database, label: "NG-Sync" }
                        ].map((item, i) => (
                            <div key={i} className="flex items-center gap-2.5 text-[10px] font-black tracking-[0.2em] uppercase text-[hsl(var(--muted))] group cursor-default">
                                <item.icon className="w-4 h-4 text-[hsl(var(--muted))] group-hover:text-blue-500 transition-colors" />
                                <span className="group-hover:text-[hsl(var(--foreground))] transition-colors">{item.label}</span>
                            </div>
                        ))}
                    </div>
                </motion.div>
            </nav>

            {/* Main Application Shell */}
            <AnimatePresence mode="wait">
                <motion.div
                    key={view}
                    initial={{ opacity: 0, scale: 0.98, y: 40 }}
                    animate={{ opacity: 1, scale: 1, y: 0 }}
                    exit={{ opacity: 0, scale: 1.02, y: -40 }}
                    transition={{ type: "spring", damping: 25, stiffness: 120 }}
                    className="w-full relative z-20"
                >
                    {view === "WORKFLOW" ? <Workflow /> : view === "ADMIN" ? <AdminPanel /> : <TemplateReference />}
                </motion.div>
            </AnimatePresence>

            {/* Premium Footer */}
            <footer className="mt-40 w-full max-w-7xl pt-16 border-t border-[hsl(var(--border))] flex flex-col md:flex-row justify-between items-center gap-12 group">
                <div className="flex flex-col md:flex-row items-center gap-6">
                    <div className="text-[10px] font-black uppercase tracking-[0.5em] text-[hsl(var(--muted))] group-hover:text-[hsl(var(--foreground))] transition-colors">
                        Proprietary Automation Framework &bull; M88 Solutions
                    </div>
                </div>

                <div className="flex gap-10">
                    {footerLinks.map(link => (
                        <button
                            key={link.label}
                            type="button"
                            onClick={link.onClick}
                            className="text-[10px] font-black uppercase tracking-[0.3em] text-[hsl(var(--muted))] hover:text-[hsl(var(--foreground))] transition-all relative group/link"
                        >
                            {link.label}
                            <span className="absolute -bottom-2 left-0 w-0 h-px bg-blue-500 transition-all group-hover/link:w-full" />
                        </button>
                    ))}

                    <button
                        type="button"
                        onClick={checkSystemHealth}
                        className="inline-flex items-center gap-2 text-[10px] font-black uppercase tracking-[0.25em] text-[hsl(var(--muted))] hover:text-[hsl(var(--foreground))] transition-colors"
                        title="Refresh system health"
                    >
                        <span className={`inline-flex h-2 w-2 rounded-full ${healthStyle.dot}`} />
                        <span className={`${healthStyle.textClass} text-[hsl(var(--foreground))]`}>{healthStyle.text}</span>
                    </button>
                </div>
            </footer>

            <AnimatePresence>
                {showAdminLogin && (
                    <motion.div
                        initial={{ opacity: 0 }}
                        animate={{ opacity: 1 }}
                        exit={{ opacity: 0 }}
                        className="fixed inset-0 z-[120] bg-[rgba(0,0,0,0.55)] backdrop-blur-sm flex items-center justify-center p-6"
                    >
                        <motion.div
                            initial={{ opacity: 0, y: 20, scale: 0.98 }}
                            animate={{ opacity: 1, y: 0, scale: 1 }}
                            exit={{ opacity: 0, y: 14, scale: 0.98 }}
                            className="w-full max-w-md rounded-3xl border border-[hsl(var(--border))] bg-[hsl(var(--panel))] p-7 shadow-2xl"
                        >
                            <h3 className="text-xl font-black uppercase tracking-tight text-[hsl(var(--foreground))]">Admin Login</h3>
                            <p className="mt-2 text-[11px] text-[hsl(var(--muted))] uppercase tracking-[0.16em] font-black">Dashboard and mapping tables require authentication.</p>

                            <input
                                type="password"
                                value={adminPassword}
                                onChange={(e) => setAdminPassword(e.target.value)}
                                onKeyDown={(e) => {
                                    if (e.key === "Enter") attemptAdminLogin();
                                }}
                                placeholder="Enter admin password"
                                className="mt-6 w-full h-12 rounded-xl bg-[hsl(var(--background)/0.45)] border border-[hsl(var(--border))] px-4 text-sm text-[hsl(var(--foreground))] placeholder:text-[hsl(var(--muted))] outline-none focus:border-blue-500/50"
                            />

                            {adminLoginError && <p className="mt-3 text-xs text-rose-400">{adminLoginError}</p>}

                            <div className="mt-6 flex justify-end gap-3">
                                <button
                                    type="button"
                                    onClick={() => {
                                        setShowAdminLogin(false);
                                        setAdminLoginError("");
                                        setAdminPassword("");
                                    }}
                                    className="h-11 px-5 rounded-xl border border-[hsl(var(--border))] text-[hsl(var(--muted))] hover:text-[hsl(var(--foreground))] hover:border-[hsl(var(--foreground)/0.24)] text-xs font-black uppercase tracking-[0.15em]"
                                >
                                    Cancel
                                </button>
                                <button
                                    type="button"
                                    onClick={attemptAdminLogin}
                                    disabled={isAdminLoginLoading}
                                    className="h-11 px-5 rounded-xl bg-blue-600 text-[hsl(var(--foreground))] disabled:opacity-60 text-xs font-black uppercase tracking-[0.15em]"
                                >
                                    {isAdminLoginLoading ? "Signing in..." : "Sign In"}
                                </button>
                            </div>
                        </motion.div>
                    </motion.div>
                )}
            </AnimatePresence>
        </main>
    );
}
