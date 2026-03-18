"use client";

import { motion } from "framer-motion";
import { FileSpreadsheet, List, Ruler, Info, CheckCircle2 } from "lucide-react";

const TEMPLATES = [
    {
        title: "ORDERS.xlsx",
        subtitle: "PO Header & Factory Assignment",
        icon: FileSpreadsheet,
        color: "text-blue-400",
        cols: [
            { name: "PurchaseOrder", desc: "Internal PO Number (Required)", req: true },
            { name: "ProductSupplier", desc: "Factory Code (Automatic/Manual)", req: true },
            { name: "Customer", desc: "NextGen Customer Name", req: true },
            { name: "KeyUser1-5", desc: "MLO / Merchandiser Codes", req: false },
            { name: "KeyDate", desc: "PO Issuance Date", req: false },
        ]
    },
    {
        title: "LINES.xlsx",
        subtitle: "Style & Delivery Items",
        icon: List,
        color: "text-indigo-400",
        cols: [
            { name: "PurchaseOrder", desc: "Links to Header PO", req: true },
            { name: "LineItem", desc: "Sequential Row Index", req: true },
            { name: "Product", desc: "Style Number", req: true },
            { name: "ProductRange", desc: "Season Code (7-chars)", req: true },
            { name: "UDF-buyer_po_number", desc: "Raw PO Number", req: false },
        ]
    },
    {
        title: "ORDER_SIZES.xlsx",
        subtitle: "Size Breakdown & Quantities",
        icon: Ruler,
        color: "text-sky-400",
        cols: [
            { name: "PurchaseOrder", desc: "Links to Header PO", req: true },
            { name: "LineItem", desc: "Links to Style Row", req: true },
            { name: "Product", desc: "Style Number", req: true },
            { name: "SizeName", desc: "Size Label (S, M, L)", req: true },
            { name: "Quantity", desc: "Number of units", req: true },
        ]
    }
];

export default function TemplateReference() {
    return (
        <div className="w-full max-w-7xl mx-auto space-y-12 px-1 sm:px-2">
            <div className="flex flex-col gap-4 md:flex-row md:items-end md:justify-between">
                <div className="space-y-2">
                    <h2 className="text-3xl md:text-4xl font-black text-white tracking-tight">System Templates</h2>
                    <p className="text-slate-500 text-xs sm:text-sm tracking-[0.18em] uppercase font-bold">Standard NextGen Format Reference</p>
                </div>

                <div className="inline-flex items-center gap-2 self-start md:self-auto rounded-xl border border-white/10 bg-white/5 px-3 py-2 text-[10px] uppercase tracking-[0.2em] text-slate-300 font-bold">
                    <span className="h-2 w-2 rounded-full bg-emerald-400 shadow-[0_0_10px_rgba(74,222,128,0.7)]" />
                    Live Field Blueprint
                </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-6 xl:gap-8">
                {TEMPLATES.map((tpl, i) => (
                    <motion.div
                        key={tpl.title}
                        initial={{ opacity: 0, y: 20 }}
                        animate={{ opacity: 1, y: 0 }}
                        transition={{ delay: i * 0.1 }}
                        className="group relative h-full bg-white/5 border border-white/10 rounded-3xl p-6 sm:p-7 backdrop-blur-xl hover:bg-white/[0.08] hover:border-white/20 transition-all shadow-2xl overflow-hidden"
                    >
                        <div className="absolute -top-4 -right-4 p-6 opacity-10 group-hover:opacity-20 transition-opacity">
                            <tpl.icon className="w-24 h-24 sm:w-28 sm:h-28" />
                        </div>
                        <div className="absolute inset-0 bg-gradient-to-b from-transparent via-transparent to-blue-950/20 pointer-events-none" />

                        <div className="relative flex items-center gap-4 mb-6">
                            <div className={`p-3.5 rounded-2xl bg-white/5 border border-white/10 ${tpl.color} group-hover:scale-110 transition-transform`}>
                                <tpl.icon className="w-6 h-6" />
                            </div>
                            <div className="space-y-1">
                                <h3 className="text-xl sm:text-2xl font-black text-white leading-tight">{tpl.title}</h3>
                                <p className="text-[10px] text-slate-500 font-bold uppercase tracking-[0.14em]">{tpl.subtitle}</p>
                            </div>
                        </div>

                        <div className="relative space-y-3.5">
                            {tpl.cols.map((col, ci) => (
                                <div key={ci} className="template-row flex flex-col gap-1.5 p-3.5 rounded-xl bg-black/20 border border-white/5 hover:border-white/15 transition-colors">
                                    <div className="flex items-center justify-between">
                                        <span className="text-sm sm:text-[15px] font-extrabold text-slate-200">{col.name}</span>
                                        {col.req && <CheckCircle2 className="w-3.5 h-3.5 text-blue-500" />}
                                    </div>
                                    <p className="template-row-desc text-[10px] sm:text-[11px] text-slate-500 font-medium">{col.desc}</p>
                                </div>
                            ))}
                        </div>

                        <div className="relative mt-7 flex items-center justify-between gap-2 text-[10px] font-black uppercase tracking-widest text-blue-400 light-blue-text opacity-75">
                            <div className="flex items-center gap-2">
                                <Info className="w-3 h-3" />
                                <span>Round-Trip Compatible</span>
                            </div>
                            <span className="rounded-md border border-white/10 px-2 py-1 text-[9px] tracking-[0.16em] text-slate-300 light-strong-text">Ready</span>
                        </div>
                    </motion.div>
                ))}
            </div>
        </div>
    );
}
