"use client";

import { useState, useEffect } from "react";
import {
    Settings, Shield, Edit2, Save, X, ChevronRight, Search, History,
    Users, BarChart3, Clock, Database, Lock, UserCircle2, Filter, Zap, Activity, ArrowLeft
} from "lucide-react";
import { clsx, type ClassValue } from "clsx";
import { twMerge } from "tailwind-merge";
import { motion, AnimatePresence } from "framer-motion";

function cn(...inputs: ClassValue[]) {
    return twMerge(clsx(inputs));
}

type Tab = "DASHBOARD" | "MAPPINGS" | "USERS" | "SECURITY";

type FactoryMappingRow = {
    id: string;
    brand: string;
    category: string;
    product_supplier: string;
    updated_at?: string;
};

type AdminPanelProps = {
    onReturnToWorkflow?: () => void;
};

export default function AdminPanel({ onReturnToWorkflow }: AdminPanelProps) {
    const [activeTab, setActiveTab] = useState<Tab>("DASHBOARD");
    const [isLoading, setIsLoading] = useState(false);
    const [searchTerm, setSearchTerm] = useState("");
    const [mappingRows, setMappingRows] = useState<FactoryMappingRow[]>([]);
    const [mappingError, setMappingError] = useState<string>("");
    const [showNewRecordModal, setShowNewRecordModal] = useState(false);
    const [newBrand, setNewBrand] = useState("");
    const [newCategory, setNewCategory] = useState("");
    const [newSupplier, setNewSupplier] = useState("");
    const [isSavingRecord, setIsSavingRecord] = useState(false);

    const fetchMappings = async () => {
        setIsLoading(true);
        setMappingError("");
        try {
            const res = await fetch("/api/admin/factory", { cache: "no-store" });
            if (!res.ok) {
                const text = await res.text();
                throw new Error(text || `Request failed (${res.status})`);
            }
            const data = (await res.json()) as FactoryMappingRow[];
            setMappingRows(Array.isArray(data) ? data : []);
        } catch (err) {
            const message = err instanceof Error ? err.message : "Failed to load mappings";
            setMappingError(message);
        } finally {
            setIsLoading(false);
        }
    };

    useEffect(() => {
        if (activeTab !== "MAPPINGS") return;
        fetchMappings();
    }, [activeTab]);

    const submitNewRecord = async () => {
        const brand = newBrand.trim();
        const category = newCategory.trim();
        const product_supplier = newSupplier.trim();
        if (!brand || !category || !product_supplier) {
            setMappingError("Brand, Category, and ProductSupplier are required.");
            return;
        }

        setIsSavingRecord(true);
        setMappingError("");
        try {
            const res = await fetch("/api/admin/factory", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ brand, category, product_supplier }),
            });

            if (!res.ok) {
                const text = await res.text();
                throw new Error(text || `Failed to save record (${res.status})`);
            }

            setShowNewRecordModal(false);
            setNewBrand("");
            setNewCategory("");
            setNewSupplier("");
            await fetchMappings();
        } catch (err) {
            const message = err instanceof Error ? err.message : "Failed to save new record";
            setMappingError(message);
        } finally {
            setIsSavingRecord(false);
        }
    };

    const filteredRows = mappingRows.filter((row) => {
        const q = searchTerm.trim().toLowerCase();
        if (!q) return true;
        return [row.brand, row.category, row.product_supplier]
            .filter(Boolean)
            .some((field) => field.toLowerCase().includes(q));
    });

    return (
        <div className="w-full space-y-16">
            {/* Control Center Header */}
            <div className="flex flex-col md:flex-row justify-between items-center gap-12">
                <div className="text-center md:text-left space-y-2">
                    <h2 className="text-5xl font-black tracking-tighter uppercase text-white">Control Center</h2>
                    <p className="text-blue-500/60 text-[10px] uppercase font-black tracking-[0.4em]">Madison 88 Operations Governance</p>
                </div>

                <div className="flex flex-col sm:flex-row items-center gap-4">
                    {onReturnToWorkflow && (
                        <button
                            type="button"
                            onClick={onReturnToWorkflow}
                            className="secondary-button light-strong-text px-6 py-3 rounded-2xl text-[10px] tracking-[0.3em] flex items-center gap-3"
                        >
                            <ArrowLeft className="w-4 h-4" />
                            RETURN TO WORKFLOW
                        </button>
                    )}

                    <div className="flex bg-slate-900/40 backdrop-blur-xl p-1.5 rounded-[24px] border border-white/5 shadow-2xl">
                        {(["DASHBOARD", "MAPPINGS", "USERS", "SECURITY"] as const).map(tab => (
                            <button
                                key={tab}
                                onClick={() => setActiveTab(tab)}
                                className={cn(
                                    "px-8 py-3.5 rounded-[18px] text-[10px] font-black transition-all duration-500 uppercase tracking-widest whitespace-nowrap relative overflow-hidden group",
                                    activeTab === tab ? "text-white" : "text-slate-500 hover:text-slate-300"
                                )}
                            >
                                <span className="relative z-10">{tab}</span>
                                {activeTab === tab && (
                                    <motion.div
                                        layoutId="activeTab"
                                        className="absolute inset-0 bg-blue-600 shadow-[0_0_20px_rgba(37,99,235,0.4)]"
                                        transition={{ type: "spring", bounce: 0.2, duration: 0.6 }}
                                    />
                                )}
                            </button>
                        ))}
                    </div>
                </div>
            </div>

            {/* Main Admin Content Area */}
            <div className="glass-panel rounded-[56px] min-h-[700px] overflow-hidden relative border-white/[0.08]">
                <div className="scanline" />

                <AnimatePresence mode="wait">
                    {activeTab === "DASHBOARD" && (
                        <motion.div
                            key="dash"
                            initial={{ opacity: 0, y: 30 }}
                            animate={{ opacity: 1, y: 0 }}
                            exit={{ opacity: 0, y: -30 }}
                            className="p-12 md:p-20 space-y-20"
                        >
                            {/* Stat Grid */}
                            <div className="grid grid-cols-1 md:grid-cols-4 gap-8">
                                {[
                                    { label: "Throughput", value: "84.2k", sub: "Last 30 days", icon: BarChart3, color: "text-blue-500", glow: "shadow-blue-500/10" },
                                    { label: "Active Nodes", value: "3", sub: "Production", icon: Database, color: "text-indigo-500", glow: "shadow-indigo-500/10" },
                                    { label: "Uptime", value: "99.99%", sub: "High Priority", icon: Zap, color: "text-amber-500", glow: "shadow-amber-500/10" },
                                    { label: "Incidents", value: "0", sub: "Critical Level", icon: Shield, color: "text-emerald-500", glow: "shadow-emerald-500/10" },
                                ].map(stat => (
                                    <div key={stat.label} className={cn("glass-card p-10 rounded-[40px] relative group overflow-hidden border-white/[0.05]", stat.glow)}>
                                        <div className="absolute top-0 right-0 w-32 h-32 bg-white/[0.02] blur-3xl -z-10" />
                                        <stat.icon className={cn("w-7 h-7 mb-6 transition-all duration-500 group-hover:scale-110 group-hover:rotate-6", stat.color)} />
                                        <p className="text-4xl font-black text-white tracking-tighter">{stat.value}</p>
                                        <p className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-500 mt-3">{stat.label}</p>
                                        <div className="flex items-center gap-2 mt-2">
                                            <div className="w-1 h-1 rounded-full bg-slate-700" />
                                            <p className="text-[9px] text-slate-700 font-bold uppercase tracking-widest">{stat.sub}</p>
                                        </div>
                                    </div>
                                ))}
                            </div>

                            {/* Activity Section */}
                            <div className="space-y-10">
                                <div className="flex items-center justify-between px-6">
                                    <div className="space-y-1">
                                        <h3 className="text-[11px] font-black uppercase tracking-[0.4em] flex items-center gap-4 text-slate-400">
                                            <Activity className="w-4 h-4 text-blue-500" /> Recent Activity Cloud
                                        </h3>
                                        <p className="text-[9px] text-slate-600 font-bold uppercase tracking-widest pl-8">Real-time system state monitoring</p>
                                    </div>
                                    <button className="text-[9px] font-black text-blue-500 hover:text-blue-400 transition-colors tracking-[0.2em] uppercase bg-blue-500/5 px-4 py-2 rounded-xl border border-blue-500/10 hover:border-blue-500/30">EXTRACT FULL LOGS</button>
                                </div>

                                <div className="space-y-4">
                                    {[1, 2, 3, 4].map(i => (
                                        <div key={i} className="glass-card flex items-center gap-8 p-8 rounded-[32px] hover:bg-white/[0.03] border-white/[0.03] group">
                                            <div className="w-14 h-14 bg-slate-950 border border-white/5 rounded-2xl flex items-center justify-center transition-all duration-500 group-hover:border-blue-500/30 group-hover:shadow-[0_0_20px_rgba(59,130,246,0.1)]">
                                                <UserCircle2 className="w-7 h-7 text-slate-600 group-hover:text-blue-400 transition-colors" />
                                            </div>
                                            <div className="flex-1 space-y-2">
                                                <div className="flex items-center gap-4">
                                                    <span className="font-black text-base tracking-tight text-white uppercase">System Synchronized</span>
                                                    <span className="text-[9px] px-3 py-1 rounded-lg bg-blue-500/10 text-blue-400 border border-blue-500/20 font-black uppercase tracking-widest shadow-[0_0_15px_rgba(59,130,246,0.1)]">NG_PUSH</span>
                                                </div>
                                                <p className="text-[9px] text-slate-600 font-black uppercase tracking-[0.2em]">AuditID: 88-XF-{i}920 &bull; Operator: <span className="text-slate-400">System_Core_v2</span></p>
                                            </div>
                                            <div className="text-right space-y-1">
                                                <p className="text-[11px] font-black text-slate-600 group-hover:text-slate-400 transition-colors">03:42:0{i} AM</p>
                                                <div className="flex items-center justify-end gap-2">
                                                    <div className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse" />
                                                    <p className="text-[10px] font-black text-emerald-500 uppercase tracking-widest">Success</p>
                                                </div>
                                            </div>
                                        </div>
                                    ))}
                                </div>
                            </div>
                        </motion.div>
                    )}

                    {activeTab === "MAPPINGS" && (
                        <motion.div
                            key="mappings"
                            initial={{ opacity: 0, scale: 0.98 }}
                            animate={{ opacity: 1, scale: 1 }}
                            className="p-12 md:p-20 space-y-12"
                        >
                            <div className="flex flex-col md:flex-row justify-between items-center gap-8 bg-slate-950/40 p-6 rounded-[32px] border border-white/5 backdrop-blur-md">
                                <div className="relative w-full md:w-[400px]">
                                    <Search className="absolute left-6 top-1/2 -translate-y-1/2 w-5 h-5 text-slate-600" />
                                    <input
                                        type="text"
                                        placeholder="QUERY REPOSITORY..."
                                        value={searchTerm}
                                        onChange={(e) => setSearchTerm(e.target.value)}
                                        className="w-full h-16 pl-16 pr-8 bg-slate-900/60 border border-white/5 rounded-2xl text-[11px] font-black uppercase tracking-[0.2em] placeholder:text-slate-700 outline-none focus:border-blue-500/40 transition-all focus:shadow-[0_0_40px_rgba(59,130,246,0.1)]"
                                    />
                                </div>
                                <div className="flex gap-4 w-full md:w-auto">
                                    <button type="button" className="secondary-button light-strong-text h-16 px-10 flex-1 md:flex-none flex items-center gap-3 text-[10px] tracking-[0.3em]">
                                        <Filter className="w-5 h-5" /> FILTERS
                                    </button>
                                    <button
                                        type="button"
                                        onClick={() => setShowNewRecordModal(true)}
                                        className="primary-button h-16 px-10 flex-1 md:flex-none flex items-center gap-3 text-[10px] tracking-[0.3em]"
                                    >
                                        + NEW RECORD
                                    </button>
                                </div>
                            </div>

                            <div className="glass-card rounded-[48px] overflow-hidden border-white/[0.08] bg-slate-950/20 relative min-h-[400px]">
                                <div className="absolute inset-0 bg-blue-500/5 blur-[120px] pointer-events-none" />

                                {isLoading ? (
                                    <div className="flex flex-col items-center justify-center p-24 text-center gap-6 relative z-10">
                                        <Database className="w-14 h-14 text-blue-500/50 animate-pulse" />
                                        <p className="text-sm font-black uppercase tracking-[0.2em] text-slate-300">Loading Factory Mappings...</p>
                                    </div>
                                ) : mappingError ? (
                                    <div className="flex flex-col items-center justify-center p-24 text-center gap-4 relative z-10">
                                        <X className="w-12 h-12 text-rose-400" />
                                        <p className="text-sm font-black uppercase tracking-[0.2em] text-rose-300">Unable to load data</p>
                                        <p className="text-[10px] text-slate-500 max-w-xl">{mappingError}</p>
                                    </div>
                                ) : filteredRows.length === 0 ? (
                                    <div className="flex flex-col items-center justify-center p-24 text-center gap-4 relative z-10">
                                        <Search className="w-12 h-12 text-slate-500" />
                                        <p className="text-sm font-black uppercase tracking-[0.2em] text-slate-300">No records found</p>
                                        <p className="text-[10px] text-slate-500">Try a different search query.</p>
                                    </div>
                                ) : (
                                    <div className="relative z-10 p-6 md:p-8">
                                        <div className="overflow-auto rounded-2xl border border-white/10">
                                            <table className="w-full min-w-[760px] text-left">
                                                <thead className="bg-slate-900/70 border-b border-white/10">
                                                    <tr>
                                                        {[
                                                            "Brand",
                                                            "Category",
                                                            "ProductSupplier",
                                                            "Last Updated",
                                                        ].map((h) => (
                                                            <th key={h} className="px-5 py-4 text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">
                                                                {h}
                                                            </th>
                                                        ))}
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    {filteredRows.map((row) => (
                                                        <tr key={row.id} className="border-b border-white/5 hover:bg-white/[0.03]">
                                                            <td className="px-5 py-4 text-sm font-bold text-white">{row.brand || "-"}</td>
                                                            <td className="px-5 py-4 text-sm text-slate-300">{row.category || "-"}</td>
                                                            <td className="px-5 py-4 text-sm text-blue-300">{row.product_supplier || "-"}</td>
                                                            <td className="px-5 py-4 text-xs text-slate-500">{row.updated_at ? new Date(row.updated_at).toLocaleString() : "-"}</td>
                                                        </tr>
                                                    ))}
                                                </tbody>
                                            </table>
                                        </div>
                                        <p className="mt-4 text-[10px] uppercase tracking-[0.2em] text-slate-500 font-black">
                                            {filteredRows.length} records loaded
                                        </p>
                                    </div>
                                )}
                            </div>

                            <AnimatePresence>
                                {showNewRecordModal && (
                                    <motion.div
                                        initial={{ opacity: 0 }}
                                        animate={{ opacity: 1 }}
                                        exit={{ opacity: 0 }}
                                        className="fixed inset-0 z-[120] bg-slate-950/75 backdrop-blur-sm flex items-center justify-center p-6"
                                    >
                                        <motion.div
                                            initial={{ opacity: 0, y: 20, scale: 0.98 }}
                                            animate={{ opacity: 1, y: 0, scale: 1 }}
                                            exit={{ opacity: 0, y: 16, scale: 0.98 }}
                                            className="w-full max-w-xl rounded-3xl border border-white/10 bg-slate-950 p-8 shadow-2xl space-y-6"
                                        >
                                            <div className="flex items-start justify-between gap-4">
                                                <div>
                                                    <h4 className="text-xl font-black tracking-tight text-white uppercase">New Factory Mapping</h4>
                                                    <p className="text-[10px] uppercase tracking-[0.2em] text-slate-500 font-black mt-1">Create or update Brand + Category supplier map</p>
                                                </div>
                                                <button
                                                    type="button"
                                                    onClick={() => setShowNewRecordModal(false)}
                                                    className="h-9 w-9 rounded-xl border border-white/10 text-slate-400 hover:text-white hover:border-white/30 transition-colors"
                                                >
                                                    <X className="w-4 h-4 mx-auto" />
                                                </button>
                                            </div>

                                            <div className="grid grid-cols-1 gap-4">
                                                <input
                                                    value={newBrand}
                                                    onChange={(e) => setNewBrand(e.target.value)}
                                                    placeholder="Brand"
                                                    className="h-12 rounded-xl bg-slate-900/70 border border-white/10 px-4 text-sm text-white placeholder:text-slate-500 outline-none focus:border-blue-500/50"
                                                />
                                                <input
                                                    value={newCategory}
                                                    onChange={(e) => setNewCategory(e.target.value)}
                                                    placeholder="Category"
                                                    className="h-12 rounded-xl bg-slate-900/70 border border-white/10 px-4 text-sm text-white placeholder:text-slate-500 outline-none focus:border-blue-500/50"
                                                />
                                                <input
                                                    value={newSupplier}
                                                    onChange={(e) => setNewSupplier(e.target.value)}
                                                    placeholder="ProductSupplier"
                                                    className="h-12 rounded-xl bg-slate-900/70 border border-white/10 px-4 text-sm text-white placeholder:text-slate-500 outline-none focus:border-blue-500/50"
                                                />
                                            </div>

                                            <div className="flex justify-end gap-3">
                                                <button
                                                    type="button"
                                                    onClick={() => setShowNewRecordModal(false)}
                                                    className="h-11 px-5 rounded-xl border border-white/10 text-slate-300 hover:text-white hover:border-white/30 text-xs font-black uppercase tracking-[0.15em]"
                                                >
                                                    Cancel
                                                </button>
                                                <button
                                                    type="button"
                                                    disabled={isSavingRecord}
                                                    onClick={submitNewRecord}
                                                    className="h-11 px-5 rounded-xl bg-blue-600 text-white disabled:opacity-60 text-xs font-black uppercase tracking-[0.15em]"
                                                >
                                                    {isSavingRecord ? "Saving..." : "Save Record"}
                                                </button>
                                            </div>
                                        </motion.div>
                                    </motion.div>
                                )}
                            </AnimatePresence>
                        </motion.div>
                    )}

                    {activeTab === "SECURITY" && (
                        <motion.div
                            key="sec"
                            initial={{ opacity: 0, scale: 1.02 }}
                            animate={{ opacity: 1, scale: 1 }}
                            className="p-12 md:p-20"
                        >
                            <div className="max-w-4xl mx-auto space-y-16">
                                <div className="text-center space-y-6">
                                    <div className="relative w-24 h-24 mx-auto">
                                        <div className="absolute inset-0 bg-blue-600/20 blur-[50px] animate-pulse rounded-full" />
                                        <div className="relative w-full h-full bg-slate-900 border border-blue-500/30 rounded-[30%] flex items-center justify-center shadow-2xl">
                                            <Shield className="w-12 h-12 text-blue-500" />
                                        </div>
                                    </div>
                                    <div className="space-y-2">
                                        <h3 className="text-4xl font-black tracking-tighter uppercase text-white">Access Matrix</h3>
                                        <p className="text-slate-500 text-sm font-black uppercase tracking-widest opacity-60">Multi-layer security protocols for PO administration</p>
                                    </div>
                                </div>

                                <div className="grid grid-cols-1 gap-6">
                                    {[
                                        { title: "Supabase RLS", desc: "Row Level Security enforcing role-based data isolation", status: "Active" },
                                        { title: "Audit Persistence", desc: "Non-volatile storage of system state transitions", status: "Enabled" },
                                        { title: "Data Anonymization", desc: "Masking customer identifiers in transit", status: "Disabled" },
                                    ].map(prop => (
                                        <div key={prop.title} className="glass-card p-10 rounded-[32px] flex flex-col md:flex-row items-center justify-between gap-8 border-white/[0.05] hover:border-white/20">
                                            <div className="space-y-3 text-center md:text-left">
                                                <h4 className="font-black text-lg uppercase tracking-widest text-white">{prop.title}</h4>
                                                <p className="text-[10px] text-slate-500 font-bold uppercase tracking-widest italic">{prop.desc}</p>
                                            </div>
                                            <div className={cn(
                                                "px-10 py-3 rounded-2xl text-[10px] font-black uppercase tracking-[0.3em] border transition-all duration-500",
                                                prop.status === "Active" || prop.status === "Enabled" ? "bg-emerald-500/10 text-emerald-500 border-emerald-500/20 shadow-[0_0_30px_rgba(16,185,129,0.1)]" : "bg-slate-900 text-slate-700 border-white/5"
                                            )}>
                                                {prop.status}
                                            </div>
                                        </div>
                                    ))}
                                </div>
                            </div>
                        </motion.div>
                    )}
                </AnimatePresence>
            </div>
        </div>
    );
}
