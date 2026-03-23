"use client";

import { useState, useEffect, useCallback } from "react";
import {
    Settings, Shield, Edit2, Save, X, Search, History,
    Users, BarChart3, Database, Lock, UserCircle2, Filter, Zap, Activity, ArrowLeft,
    Trash2, Plus, ChevronDown
} from "lucide-react";
import { clsx, type ClassValue } from "clsx";
import { twMerge } from "tailwind-merge";
import { motion, AnimatePresence } from "framer-motion";

function cn(...inputs: ClassValue[]) {
    return twMerge(clsx(inputs));
}

type Tab = "DASHBOARD" | "MAPPINGS" | "USERS" | "SECURITY";
type MappingTab = "factory" | "columns" | "mlo";

// ── Types ──────────────────────────────────────────────────────────────────────

type FactoryRow = {
    id: string;
    brand: string;
    category: string;
    product_supplier: string;
    updated_at?: string;
};

type ColumnRow = {
    id: string;
    customer: string;
    buy_file_column: string;
    internal_field: string;
    notes?: string;
    updated_at?: string;
};

type MloRow = {
    id: string;
    brand: string;
    keyuser1: string;
    keyuser2: string;
    keyuser4: string;
    keyuser5: string;
    orders_template?: string | null;
    lines_template?: string | null;
    valid_statuses?: string[] | null;
    updated_at?: string;
};

type AdminPanelProps = {
    onReturnToWorkflow?: () => void;
};

// ── Helpers ────────────────────────────────────────────────────────────────────

function ModalOverlay({ children, onClose }: { children: React.ReactNode; onClose: () => void }) {
    return (
        <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 z-[120] bg-slate-950/75 backdrop-blur-sm flex items-center justify-center p-6"
            onClick={(e) => { if (e.target === e.currentTarget) onClose(); }}
        >
            <motion.div
                initial={{ opacity: 0, y: 20, scale: 0.98 }}
                animate={{ opacity: 1, y: 0, scale: 1 }}
                exit={{ opacity: 0, y: 16, scale: 0.98 }}
                className="w-full max-w-xl rounded-3xl border border-white/10 bg-slate-950 p-8 shadow-2xl space-y-6"
            >
                {children}
            </motion.div>
        </motion.div>
    );
}

function ModalHeader({ title, sub, onClose }: { title: string; sub?: string; onClose: () => void }) {
    return (
        <div className="flex items-start justify-between gap-4">
            <div>
                <h4 className="text-xl font-black tracking-tight text-white uppercase">{title}</h4>
                {sub && <p className="text-[10px] uppercase tracking-[0.2em] text-slate-500 font-black mt-1">{sub}</p>}
            </div>
            <button type="button" onClick={onClose}
                className="h-9 w-9 rounded-xl border border-white/10 text-slate-400 hover:text-white hover:border-white/30 transition-colors flex items-center justify-center">
                <X className="w-4 h-4" />
            </button>
        </div>
    );
}

function Field({ label, value, onChange, placeholder }: { label: string; value: string; onChange: (v: string) => void; placeholder?: string }) {
    return (
        <div className="space-y-1.5">
            <label className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-400">{label}</label>
            <input
                value={value}
                onChange={(e) => onChange(e.target.value)}
                placeholder={placeholder || label}
                className="w-full h-11 rounded-xl bg-slate-900/70 border border-white/10 px-4 text-sm text-white placeholder:text-slate-600 outline-none focus:border-blue-500/50 transition-colors"
            />
        </div>
    );
}

function ModalActions({ onCancel, onSave, saving }: { onCancel: () => void; onSave: () => void; saving: boolean }) {
    return (
        <div className="flex justify-end gap-3 pt-2">
            <button type="button" onClick={onCancel}
                className="h-11 px-5 rounded-xl border border-white/10 text-slate-300 hover:text-white hover:border-white/30 text-xs font-black uppercase tracking-[0.15em] transition-colors">
                Cancel
            </button>
            <button type="button" disabled={saving} onClick={onSave}
                className="h-11 px-6 rounded-xl bg-blue-600 hover:bg-blue-500 text-white disabled:opacity-60 text-xs font-black uppercase tracking-[0.15em] transition-colors">
                {saving ? "Saving..." : "Save"}
            </button>
        </div>
    );
}

function TableHeader({ cols }: { cols: string[] }) {
    return (
        <thead className="bg-slate-900/70 border-b border-white/10">
            <tr>
                {cols.map(h => (
                    <th key={h} className="px-5 py-4 text-[10px] font-black uppercase tracking-[0.2em] text-slate-400 whitespace-nowrap">{h}</th>
                ))}
            </tr>
        </thead>
    );
}

function EmptyState({ icon: Icon, message }: { icon: React.ElementType; message: string }) {
    return (
        <div className="flex flex-col items-center justify-center p-24 text-center gap-4 relative z-10">
            <Icon className="w-12 h-12 text-slate-500" />
            <p className="text-sm font-black uppercase tracking-[0.2em] text-slate-300">{message}</p>
        </div>
    );
}

// ── Sub-tabs ───────────────────────────────────────────────────────────────────

function MappingSubTabs({ active, onChange }: { active: MappingTab; onChange: (t: MappingTab) => void }) {
    const tabs: { id: MappingTab; label: string }[] = [
        { id: "factory", label: "Factory" },
        { id: "columns", label: "Column Aliases" },
        { id: "mlo", label: "MLO Config" },
    ];
    return (
        <div className="flex gap-2 bg-slate-900/60 p-1 rounded-2xl border border-white/5 w-fit">
            {tabs.map(t => (
                <button key={t.id} type="button" onClick={() => onChange(t.id)}
                    className={cn(
                        "px-6 py-2.5 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all duration-300",
                        active === t.id ? "bg-blue-600 text-white shadow-[0_0_20px_rgba(37,99,235,0.3)]" : "text-slate-500 hover:text-slate-300"
                    )}>
                    {t.label}
                </button>
            ))}
        </div>
    );
}

// ── Factory Mapping Panel ──────────────────────────────────────────────────────

function FactoryPanel() {
    const [rows, setRows] = useState<FactoryRow[]>([]);
    const [loading, setLoading] = useState(false);
    const [error, setError] = useState("");
    const [search, setSearch] = useState("");
    const [showModal, setShowModal] = useState(false);
    const [saving, setSaving] = useState(false);
    const [form, setForm] = useState({ brand: "", category: "", product_supplier: "" });

    const load = useCallback(async () => {
        setLoading(true); setError("");
        try {
            const res = await fetch("/api/admin/factory", { cache: "no-store" });
            if (!res.ok) throw new Error(await res.text());
            setRows(await res.json());
        } catch (e) { setError(e instanceof Error ? e.message : "Failed to load"); }
        finally { setLoading(false); }
    }, []);

    useEffect(() => { load(); }, [load]);

    const save = async () => {
        if (!form.brand || !form.category || !form.product_supplier) { setError("All fields required."); return; }
        setSaving(true); setError("");
        try {
            const res = await fetch("/api/admin/factory", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(form) });
            if (!res.ok) throw new Error(await res.text());
            setShowModal(false); setForm({ brand: "", category: "", product_supplier: "" }); await load();
        } catch (e) { setError(e instanceof Error ? e.message : "Save failed"); }
        finally { setSaving(false); }
    };

    const del = async (id: string) => {
        if (!confirm("Delete this factory mapping?")) return;
        try {
            const res = await fetch("/api/admin/factory", { method: "DELETE", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ id }) });
            if (!res.ok) throw new Error(await res.text());
            await load();
        } catch (e) { setError(e instanceof Error ? e.message : "Delete failed"); }
    };

    const filtered = rows.filter(r => {
        const q = search.toLowerCase();
        return !q || [r.brand, r.category, r.product_supplier].some(v => v?.toLowerCase().includes(q));
    });

    return (
        <div className="space-y-6">
            <div className="flex flex-col sm:flex-row gap-4 items-center">
                <div className="relative flex-1">
                    <Search className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-600" />
                    <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Search..."
                        className="w-full h-12 pl-12 pr-4 bg-slate-900/60 border border-white/5 rounded-xl text-[11px] font-black uppercase tracking-[0.15em] placeholder:text-slate-700 outline-none focus:border-blue-500/40 transition-all" />
                </div>
                <button type="button" onClick={() => setShowModal(true)}
                    className="primary-button h-12 px-8 flex items-center gap-2 text-[10px] tracking-[0.3em] whitespace-nowrap">
                    <Plus className="w-4 h-4" /> NEW RECORD
                </button>
            </div>

            {error && <p className="text-xs text-rose-400 font-bold px-2">{error}</p>}

            <div className="rounded-2xl border border-white/10 overflow-auto">
                <table className="w-full min-w-[600px] text-left">
                    <TableHeader cols={["Brand", "Category", "Product Supplier", "Updated", ""]} />
                    <tbody>
                        {loading ? (
                            <tr><td colSpan={5}><EmptyState icon={Database} message="Loading..." /></td></tr>
                        ) : filtered.length === 0 ? (
                            <tr><td colSpan={5}><EmptyState icon={Search} message="No records found" /></td></tr>
                        ) : filtered.map(r => (
                            <tr key={r.id} className="border-b border-white/5 hover:bg-white/[0.03]">
                                <td className="px-5 py-3.5 text-sm font-bold text-white">{r.brand}</td>
                                <td className="px-5 py-3.5 text-sm text-slate-300">{r.category}</td>
                                <td className="px-5 py-3.5 text-sm text-blue-300">{r.product_supplier}</td>
                                <td className="px-5 py-3.5 text-xs text-slate-500">{r.updated_at ? new Date(r.updated_at).toLocaleDateString() : "-"}</td>
                                <td className="px-5 py-3.5">
                                    <button type="button" onClick={() => del(r.id)}
                                        className="text-slate-600 hover:text-rose-400 transition-colors">
                                        <Trash2 className="w-4 h-4" />
                                    </button>
                                </td>
                            </tr>
                        ))}
                    </tbody>
                </table>
            </div>
            <p className="text-[10px] uppercase tracking-[0.2em] text-slate-600 font-black">{filtered.length} records</p>

            <AnimatePresence>
                {showModal && (
                    <ModalOverlay onClose={() => setShowModal(false)}>
                        <ModalHeader title="New Factory Mapping" sub="Brand + Category → ProductSupplier" onClose={() => setShowModal(false)} />
                        <div className="space-y-3">
                            <Field label="Brand" value={form.brand} onChange={v => setForm(f => ({ ...f, brand: v }))} />
                            <Field label="Category" value={form.category} onChange={v => setForm(f => ({ ...f, category: v }))} />
                            <Field label="Product Supplier" value={form.product_supplier} onChange={v => setForm(f => ({ ...f, product_supplier: v }))} placeholder="ProductSupplier" />
                        </div>
                        {error && <p className="text-xs text-rose-400 font-bold">{error}</p>}
                        <ModalActions onCancel={() => setShowModal(false)} onSave={save} saving={saving} />
                    </ModalOverlay>
                )}
            </AnimatePresence>
        </div>
    );
}

// ── Column Mapping Panel ───────────────────────────────────────────────────────

function ColumnPanel() {
    const [rows, setRows] = useState<ColumnRow[]>([]);
    const [loading, setLoading] = useState(false);
    const [error, setError] = useState("");
    const [search, setSearch] = useState("");
    const [customerFilter, setCustomerFilter] = useState("ALL");
    const [showModal, setShowModal] = useState(false);
    const [saving, setSaving] = useState(false);
    const [form, setForm] = useState({ customer: "", buy_file_column: "", internal_field: "", notes: "" });

    const load = useCallback(async () => {
        setLoading(true); setError("");
        try {
            const res = await fetch("/api/admin/columns", { cache: "no-store" });
            if (!res.ok) throw new Error(await res.text());
            setRows(await res.json());
        } catch (e) { setError(e instanceof Error ? e.message : "Failed to load"); }
        finally { setLoading(false); }
    }, []);

    useEffect(() => { load(); }, [load]);

    const save = async () => {
        if (!form.customer || !form.buy_file_column || !form.internal_field) { setError("Customer, buy_file_column, and internal_field are required."); return; }
        setSaving(true); setError("");
        try {
            const res = await fetch("/api/admin/columns", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(form) });
            if (!res.ok) throw new Error(await res.text());
            setShowModal(false); setForm({ customer: "", buy_file_column: "", internal_field: "", notes: "" }); await load();
        } catch (e) { setError(e instanceof Error ? e.message : "Save failed"); }
        finally { setSaving(false); }
    };

    const customers = ["ALL", ...Array.from(new Set(rows.map(r => r.customer))).sort()];

    const filtered = rows.filter(r => {
        const q = search.toLowerCase();
        const matchSearch = !q || [r.customer, r.buy_file_column, r.internal_field].some(v => v?.toLowerCase().includes(q));
        const matchCustomer = customerFilter === "ALL" || r.customer === customerFilter;
        return matchSearch && matchCustomer;
    });

    return (
        <div className="space-y-6">
            <div className="flex flex-col sm:flex-row gap-4 items-center">
                <div className="relative flex-1">
                    <Search className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-600" />
                    <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Search columns..."
                        className="w-full h-12 pl-12 pr-4 bg-slate-900/60 border border-white/5 rounded-xl text-[11px] font-black uppercase tracking-[0.15em] placeholder:text-slate-700 outline-none focus:border-blue-500/40 transition-all" />
                </div>
                <div className="relative">
                    <select value={customerFilter} onChange={e => setCustomerFilter(e.target.value)}
                        className="h-12 pl-4 pr-10 bg-slate-900/60 border border-white/5 rounded-xl text-[11px] font-black uppercase tracking-[0.15em] text-slate-300 outline-none focus:border-blue-500/40 appearance-none cursor-pointer">
                        {customers.map(c => <option key={c} value={c}>{c}</option>)}
                    </select>
                    <ChevronDown className="absolute right-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-500 pointer-events-none" />
                </div>
                <button type="button" onClick={() => setShowModal(true)}
                    className="primary-button h-12 px-8 flex items-center gap-2 text-[10px] tracking-[0.3em] whitespace-nowrap">
                    <Plus className="w-4 h-4" /> NEW ALIAS
                </button>
            </div>

            {error && <p className="text-xs text-rose-400 font-bold px-2">{error}</p>}

            <div className="rounded-2xl border border-white/10 overflow-auto max-h-[500px]">
                <table className="w-full min-w-[700px] text-left">
                    <TableHeader cols={["Customer", "Buy File Column", "Internal Field", "Notes", "Updated"]} />
                    <tbody>
                        {loading ? (
                            <tr><td colSpan={5}><EmptyState icon={Database} message="Loading..." /></td></tr>
                        ) : filtered.length === 0 ? (
                            <tr><td colSpan={5}><EmptyState icon={Search} message="No column mappings found" /></td></tr>
                        ) : filtered.map(r => (
                            <tr key={r.id} className="border-b border-white/5 hover:bg-white/[0.03]">
                                <td className="px-5 py-3 text-xs font-bold text-slate-400 whitespace-nowrap">{r.customer}</td>
                                <td className="px-5 py-3 text-sm text-white font-mono">{r.buy_file_column}</td>
                                <td className="px-5 py-3 text-sm text-blue-300 font-mono">{r.internal_field}</td>
                                <td className="px-5 py-3 text-xs text-slate-500 italic">{r.notes || "-"}</td>
                                <td className="px-5 py-3 text-xs text-slate-600">{r.updated_at ? new Date(r.updated_at).toLocaleDateString() : "-"}</td>
                            </tr>
                        ))}
                    </tbody>
                </table>
            </div>
            <p className="text-[10px] uppercase tracking-[0.2em] text-slate-600 font-black">{filtered.length} of {rows.length} mappings</p>

            <AnimatePresence>
                {showModal && (
                    <ModalOverlay onClose={() => setShowModal(false)}>
                        <ModalHeader title="New Column Alias" sub="Map a buy file column header to an internal field" onClose={() => setShowModal(false)} />
                        <div className="space-y-3">
                            <Field label="Customer" value={form.customer} onChange={v => setForm(f => ({ ...f, customer: v }))} placeholder="e.g. DEFAULT or ABC Retail" />
                            <Field label="Buy File Column" value={form.buy_file_column} onChange={v => setForm(f => ({ ...f, buy_file_column: v }))} placeholder="e.g. Vendor Confirmed CRD" />
                            <Field label="Internal Field" value={form.internal_field} onChange={v => setForm(f => ({ ...f, internal_field: v }))} placeholder="e.g. confirmedExFac" />
                            <Field label="Notes (optional)" value={form.notes} onChange={v => setForm(f => ({ ...f, notes: v }))} placeholder="Optional description" />
                        </div>
                        {error && <p className="text-xs text-rose-400 font-bold">{error}</p>}
                        <ModalActions onCancel={() => setShowModal(false)} onSave={save} saving={saving} />
                    </ModalOverlay>
                )}
            </AnimatePresence>
        </div>
    );
}

// ── MLO Config Panel ───────────────────────────────────────────────────────────

function MloPanel() {
    const [rows, setRows] = useState<MloRow[]>([]);
    const [loading, setLoading] = useState(false);
    const [error, setError] = useState("");
    const [showModal, setShowModal] = useState(false);
    const [saving, setSaving] = useState(false);
    const [editRow, setEditRow] = useState<MloRow | null>(null);
    const [form, setForm] = useState({ brand: "", keyuser1: "", keyuser2: "", keyuser4: "", keyuser5: "", orders_template: "", lines_template: "", valid_statuses: "" });

    const load = useCallback(async () => {
        setLoading(true); setError("");
        try {
            const res = await fetch("/api/admin/mlo", { cache: "no-store" });
            if (!res.ok) throw new Error(await res.text());
            setRows(await res.json());
        } catch (e) { setError(e instanceof Error ? e.message : "Failed to load"); }
        finally { setLoading(false); }
    }, []);

    useEffect(() => { load(); }, [load]);

    const openEdit = (row: MloRow) => {
        setEditRow(row);
        setForm({
            brand: row.brand,
            keyuser1: row.keyuser1,
            keyuser2: row.keyuser2,
            keyuser4: row.keyuser4,
            keyuser5: row.keyuser5,
            orders_template: row.orders_template || "",
            lines_template: row.lines_template || "",
            valid_statuses: (row.valid_statuses || []).join(", "),
        });
        setShowModal(true);
    };

    const openNew = () => {
        setEditRow(null);
        setForm({ brand: "", keyuser1: "", keyuser2: "", keyuser4: "", keyuser5: "", orders_template: "", lines_template: "", valid_statuses: "" });
        setShowModal(true);
    };

    const save = async () => {
        if (!form.brand) { setError("Brand is required."); return; }
        setSaving(true); setError("");
        try {
            const payload = {
                brand: form.brand,
                keyuser1: form.keyuser1,
                keyuser2: form.keyuser2,
                keyuser4: form.keyuser4,
                keyuser5: form.keyuser5,
                orders_template: form.orders_template || null,
                lines_template: form.lines_template || null,
                valid_statuses: form.valid_statuses.split(",").map(s => s.trim()).filter(Boolean),
            };
            const res = await fetch("/api/admin/mlo", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) });
            if (!res.ok) throw new Error(await res.text());
            setShowModal(false); await load();
        } catch (e) { setError(e instanceof Error ? e.message : "Save failed"); }
        finally { setSaving(false); }
    };

    return (
        <div className="space-y-6">
            <div className="flex justify-end">
                <button type="button" onClick={openNew}
                    className="primary-button h-12 px-8 flex items-center gap-2 text-[10px] tracking-[0.3em]">
                    <Plus className="w-4 h-4" /> NEW BRAND CONFIG
                </button>
            </div>

            {error && <p className="text-xs text-rose-400 font-bold px-2">{error}</p>}

            <div className="rounded-2xl border border-white/10 overflow-auto">
                <table className="w-full min-w-[900px] text-left">
                    <TableHeader cols={["Brand", "KeyUser1", "KeyUser2", "KeyUser4", "KeyUser5", "Orders Tmpl", "Lines Tmpl", "Valid Statuses", ""]} />
                    <tbody>
                        {loading ? (
                            <tr><td colSpan={9}><EmptyState icon={Database} message="Loading..." /></td></tr>
                        ) : rows.length === 0 ? (
                            <tr><td colSpan={9}><EmptyState icon={Database} message="No MLO configs found" /></td></tr>
                        ) : rows.map(r => (
                            <tr key={r.id} className="border-b border-white/5 hover:bg-white/[0.03]">
                                <td className="px-5 py-3 text-sm font-bold text-white">{r.brand}</td>
                                <td className="px-5 py-3 text-xs text-slate-300 font-mono">{r.keyuser1 || "-"}</td>
                                <td className="px-5 py-3 text-xs text-slate-300 font-mono">{r.keyuser2 || "-"}</td>
                                <td className="px-5 py-3 text-xs text-slate-300 font-mono">{r.keyuser4 || "-"}</td>
                                <td className="px-5 py-3 text-xs text-slate-300 font-mono">{r.keyuser5 || "-"}</td>
                                <td className="px-5 py-3 text-xs text-blue-300">{r.orders_template || "-"}</td>
                                <td className="px-5 py-3 text-xs text-blue-300">{r.lines_template || "-"}</td>
                                <td className="px-5 py-3 text-xs text-slate-400">{(r.valid_statuses || []).join(", ") || "-"}</td>
                                <td className="px-5 py-3">
                                    <button type="button" onClick={() => openEdit(r)}
                                        className="text-slate-600 hover:text-blue-400 transition-colors">
                                        <Edit2 className="w-4 h-4" />
                                    </button>
                                </td>
                            </tr>
                        ))}
                    </tbody>
                </table>
            </div>
            <p className="text-[10px] uppercase tracking-[0.2em] text-slate-600 font-black">{rows.length} brand configs</p>

            <AnimatePresence>
                {showModal && (
                    <ModalOverlay onClose={() => setShowModal(false)}>
                        <ModalHeader
                            title={editRow ? "Edit MLO Config" : "New MLO Config"}
                            sub="Brand-level KeyUsers, templates, and valid statuses"
                            onClose={() => setShowModal(false)}
                        />
                        <div className="space-y-3 max-h-[60vh] overflow-y-auto pr-1">
                            <Field label="Brand" value={form.brand} onChange={v => setForm(f => ({ ...f, brand: v }))} />
                            <div className="grid grid-cols-2 gap-3">
                                <Field label="KeyUser1" value={form.keyuser1} onChange={v => setForm(f => ({ ...f, keyuser1: v }))} />
                                <Field label="KeyUser2" value={form.keyuser2} onChange={v => setForm(f => ({ ...f, keyuser2: v }))} />
                                <Field label="KeyUser4" value={form.keyuser4} onChange={v => setForm(f => ({ ...f, keyuser4: v }))} />
                                <Field label="KeyUser5" value={form.keyuser5} onChange={v => setForm(f => ({ ...f, keyuser5: v }))} />
                            </div>
                            <div className="grid grid-cols-2 gap-3">
                                <Field label="Orders Template" value={form.orders_template} onChange={v => setForm(f => ({ ...f, orders_template: v }))} placeholder="e.g. BULK" />
                                <Field label="Lines Template" value={form.lines_template} onChange={v => setForm(f => ({ ...f, lines_template: v }))} placeholder="e.g. BULK" />
                            </div>
                            <Field label="Valid Statuses (comma-separated)" value={form.valid_statuses} onChange={v => setForm(f => ({ ...f, valid_statuses: v }))} placeholder="e.g. Confirmed, Approved" />
                        </div>
                        {error && <p className="text-xs text-rose-400 font-bold">{error}</p>}
                        <ModalActions onCancel={() => setShowModal(false)} onSave={save} saving={saving} />
                    </ModalOverlay>
                )}
            </AnimatePresence>
        </div>
    );
}

// ── Main AdminPanel ────────────────────────────────────────────────────────────

export default function AdminPanel({ onReturnToWorkflow }: AdminPanelProps) {
    const [activeTab, setActiveTab] = useState<Tab>("DASHBOARD");
    const [mappingTab, setMappingTab] = useState<MappingTab>("factory");

    return (
        <div className="w-full space-y-16">
            {/* Header */}
            <div className="flex flex-col md:flex-row justify-between items-center gap-12">
                <div className="text-center md:text-left space-y-2">
                    <h2 className="text-5xl font-black tracking-tighter uppercase text-white">Control Center</h2>
                    <p className="text-blue-500/60 text-[10px] uppercase font-black tracking-[0.4em]">Madison 88 Operations Governance</p>
                </div>

                <div className="flex flex-col sm:flex-row items-center gap-4">
                    {onReturnToWorkflow && (
                        <button type="button" onClick={onReturnToWorkflow}
                            className="secondary-button light-strong-text px-6 py-3 rounded-2xl text-[10px] tracking-[0.3em] flex items-center gap-3">
                            <ArrowLeft className="w-4 h-4" /> RETURN TO WORKFLOW
                        </button>
                    )}
                    <div className="flex bg-slate-900/40 backdrop-blur-xl p-1.5 rounded-[24px] border border-white/5 shadow-2xl">
                        {(["DASHBOARD", "MAPPINGS", "USERS", "SECURITY"] as const).map(tab => (
                            <button key={tab} onClick={() => setActiveTab(tab)}
                                className={cn(
                                    "px-8 py-3.5 rounded-[18px] text-[10px] font-black transition-all duration-500 uppercase tracking-widest whitespace-nowrap relative overflow-hidden",
                                    activeTab === tab ? "text-white" : "text-slate-500 hover:text-slate-300"
                                )}>
                                <span className="relative z-10">{tab}</span>
                                {activeTab === tab && (
                                    <motion.div layoutId="activeTab"
                                        className="absolute inset-0 bg-blue-600 shadow-[0_0_20px_rgba(37,99,235,0.4)]"
                                        transition={{ type: "spring", bounce: 0.2, duration: 0.6 }} />
                                )}
                            </button>
                        ))}
                    </div>
                </div>
            </div>

            {/* Content */}
            <div className="glass-panel rounded-[56px] min-h-[700px] overflow-hidden relative border-white/[0.08]">
                <div className="scanline" />

                <AnimatePresence mode="wait">
                    {activeTab === "DASHBOARD" && (
                        <motion.div key="dash" initial={{ opacity: 0, y: 30 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -30 }}
                            className="p-12 md:p-20 space-y-20">
                            <div className="grid grid-cols-1 md:grid-cols-4 gap-8">
                                {[
                                    { label: "Throughput", value: "84.2k", sub: "Last 30 days", icon: BarChart3, color: "text-blue-500" },
                                    { label: "Active Nodes", value: "3", sub: "Production", icon: Database, color: "text-indigo-500" },
                                    { label: "Uptime", value: "99.99%", sub: "High Priority", icon: Zap, color: "text-amber-500" },
                                    { label: "Incidents", value: "0", sub: "Critical Level", icon: Shield, color: "text-emerald-500" },
                                ].map(stat => (
                                    <div key={stat.label} className="glass-card p-10 rounded-[40px] relative group overflow-hidden border-white/[0.05]">
                                        <stat.icon className={cn("w-7 h-7 mb-6 transition-all duration-500 group-hover:scale-110", stat.color)} />
                                        <p className="text-4xl font-black text-white tracking-tighter">{stat.value}</p>
                                        <p className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-500 mt-3">{stat.label}</p>
                                        <p className="text-[9px] text-slate-700 font-bold uppercase tracking-widest mt-1">{stat.sub}</p>
                                    </div>
                                ))}
                            </div>

                            <div className="space-y-4">
                                <h3 className="text-[11px] font-black uppercase tracking-[0.4em] flex items-center gap-4 text-slate-400">
                                    <Activity className="w-4 h-4 text-blue-500" /> Recent Activity
                                </h3>
                                {[1, 2, 3, 4].map(i => (
                                    <div key={i} className="glass-card flex items-center gap-8 p-8 rounded-[32px] hover:bg-white/[0.03] border-white/[0.03] group">
                                        <div className="w-14 h-14 bg-slate-950 border border-white/5 rounded-2xl flex items-center justify-center">
                                            <UserCircle2 className="w-7 h-7 text-slate-600 group-hover:text-blue-400 transition-colors" />
                                        </div>
                                        <div className="flex-1 space-y-1">
                                            <span className="font-black text-base tracking-tight text-white uppercase">System Synchronized</span>
                                            <p className="text-[9px] text-slate-600 font-black uppercase tracking-[0.2em]">AuditID: 88-XF-{i}920</p>
                                        </div>
                                        <div className="text-right">
                                            <p className="text-[11px] font-black text-slate-600">03:42:0{i} AM</p>
                                            <div className="flex items-center justify-end gap-2 mt-1">
                                                <div className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse" />
                                                <p className="text-[10px] font-black text-emerald-500 uppercase tracking-widest">Success</p>
                                            </div>
                                        </div>
                                    </div>
                                ))}
                            </div>
                        </motion.div>
                    )}

                    {activeTab === "MAPPINGS" && (
                        <motion.div key="mappings" initial={{ opacity: 0, scale: 0.98 }} animate={{ opacity: 1, scale: 1 }}
                            className="p-12 md:p-20 space-y-10">
                            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-6">
                                <div>
                                    <h3 className="text-2xl font-black uppercase tracking-tight text-white">Mapping Tables</h3>
                                    <p className="text-[10px] text-slate-500 font-black uppercase tracking-[0.3em] mt-1">Manage factory, column alias, and MLO configurations</p>
                                </div>
                                <MappingSubTabs active={mappingTab} onChange={setMappingTab} />
                            </div>

                            <AnimatePresence mode="wait">
                                {mappingTab === "factory" && (
                                    <motion.div key="factory" initial={{ opacity: 0, x: 10 }} animate={{ opacity: 1, x: 0 }} exit={{ opacity: 0, x: -10 }}>
                                        <FactoryPanel />
                                    </motion.div>
                                )}
                                {mappingTab === "columns" && (
                                    <motion.div key="columns" initial={{ opacity: 0, x: 10 }} animate={{ opacity: 1, x: 0 }} exit={{ opacity: 0, x: -10 }}>
                                        <ColumnPanel />
                                    </motion.div>
                                )}
                                {mappingTab === "mlo" && (
                                    <motion.div key="mlo" initial={{ opacity: 0, x: 10 }} animate={{ opacity: 1, x: 0 }} exit={{ opacity: 0, x: -10 }}>
                                        <MloPanel />
                                    </motion.div>
                                )}
                            </AnimatePresence>
                        </motion.div>
                    )}

                    {activeTab === "SECURITY" && (
                        <motion.div key="sec" initial={{ opacity: 0, scale: 1.02 }} animate={{ opacity: 1, scale: 1 }}
                            className="p-12 md:p-20">
                            <div className="max-w-4xl mx-auto space-y-16">
                                <div className="text-center space-y-6">
                                    <div className="relative w-24 h-24 mx-auto">
                                        <div className="absolute inset-0 bg-blue-600/20 blur-[50px] animate-pulse rounded-full" />
                                        <div className="relative w-full h-full bg-slate-900 border border-blue-500/30 rounded-[30%] flex items-center justify-center">
                                            <Shield className="w-12 h-12 text-blue-500" />
                                        </div>
                                    </div>
                                    <div>
                                        <h3 className="text-4xl font-black tracking-tighter uppercase text-white">Access Matrix</h3>
                                        <p className="text-slate-500 text-sm font-black uppercase tracking-widest opacity-60 mt-2">Multi-layer security protocols</p>
                                    </div>
                                </div>
                                <div className="grid grid-cols-1 gap-6">
                                    {[
                                        { title: "Supabase RLS", desc: "Row Level Security enforcing role-based data isolation", status: "Active" },
                                        { title: "Audit Persistence", desc: "Non-volatile storage of system state transitions", status: "Enabled" },
                                        { title: "Data Anonymization", desc: "Masking customer identifiers in transit", status: "Disabled" },
                                    ].map(prop => (
                                        <div key={prop.title} className="glass-card p-10 rounded-[32px] flex flex-col md:flex-row items-center justify-between gap-8 border-white/[0.05]">
                                            <div className="space-y-2">
                                                <h4 className="font-black text-lg uppercase tracking-widest text-white">{prop.title}</h4>
                                                <p className="text-[10px] text-slate-500 font-bold uppercase tracking-widest">{prop.desc}</p>
                                            </div>
                                            <div className={cn(
                                                "px-10 py-3 rounded-2xl text-[10px] font-black uppercase tracking-[0.3em] border",
                                                prop.status === "Active" || prop.status === "Enabled"
                                                    ? "bg-emerald-500/10 text-emerald-500 border-emerald-500/20"
                                                    : "bg-slate-900 text-slate-700 border-white/5"
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
