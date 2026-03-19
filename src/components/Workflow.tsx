"use client";

import { useEffect, useState } from "react";
import { Upload, FileCheck, AlertCircle, Download, ChevronRight, ChevronLeft, Settings, History, Loader2, Info, CheckCircle2, CloudUpload, ArrowRight, ShieldCheck, FileText } from "lucide-react";
import { clsx, type ClassValue } from "clsx";
import { twMerge } from "tailwind-merge";
import { motion, AnimatePresence } from "framer-motion";

function cn(...inputs: ClassValue[]) {
    return twMerge(clsx(inputs));
}

type Step = "UPLOAD" | "RUN" | "VALIDATE" | "REVIEW" | "DOWNLOAD";

export default function Workflow() {
    const [currentStep, setCurrentStep] = useState<Step>("UPLOAD");
    const [isProcessing, setIsProcessing] = useState(false);
    const [errors, setErrors] = useState<any[]>([]);
    const [uploadData, setUploadData] = useState<any>(null);
    const [buyFiles, setBuyFiles] = useState<FileList | null>(null);
    const [productSheetFile, setProductSheetFile] = useState<File | null>(null);
    const [manualPo, setManualPo] = useState("");
    const [manualTemplate, setManualTemplate] = useState("");
    const [manualComments, setManualComments] = useState("");
    const [manualKeyDate, setManualKeyDate] = useState("");
    const [manualKeyUser1, setManualKeyUser1] = useState("");
    const [manualKeyUser2, setManualKeyUser2] = useState("");
    const [manualKeyUser3, setManualKeyUser3] = useState("");
    const [manualKeyUser4, setManualKeyUser4] = useState("");
    const [manualKeyUser5, setManualKeyUser5] = useState("");
    const applyTheme = (nextTheme: "dark" | "light") => {
      document.documentElement.classList.remove("light", "dark");
      document.documentElement.classList.add(nextTheme);
      window.localStorage.setItem("theme", nextTheme);
    };

    useEffect(() => {
      const saved = window.localStorage.getItem("theme");
      const systemPreferred = window.matchMedia("(prefers-color-scheme: dark)").matches;
      const initial = saved === "light" || saved === "dark" ? saved : (systemPreferred ? "dark" : "light");
      applyTheme(initial as "dark" | "light");
    }, []);

    const steps: { key: Step; label: string; icon: any }[] = [
        { key: "UPLOAD", label: "Acquisition", icon: CloudUpload },
        { key: "RUN", label: "Engine", icon: Loader2 },
        { key: "VALIDATE", label: "Audit", icon: ShieldCheck },
        { key: "REVIEW", label: "Review", icon: FileCheck },
        { key: "DOWNLOAD", label: "Export", icon: Download },
    ];

    const currentStepIndex = steps.findIndex(s => s.key === currentStep);

    const handleStartUpload = async () => {
        if (!buyFiles || buyFiles.length === 0) return;

        setIsProcessing(true);
        setCurrentStep("RUN");

        const formData = new FormData();
        for (let i = 0; i < buyFiles.length; i++) {
            formData.append("file", buyFiles[i]);
        }
        if (productSheetFile) {
            formData.append("file", productSheetFile);
        }
        if (manualPo.trim()) formData.append("manualPo", manualPo.trim());
        if (manualTemplate.trim()) formData.append("manualTemplate", manualTemplate.trim());
        if (manualComments.trim()) formData.append("manualComments", manualComments.trim());
        if (manualKeyDate.trim()) formData.append("manualKeyDate", manualKeyDate.trim());
        if (manualKeyUser1.trim()) formData.append("manualKeyUser1", manualKeyUser1.trim());
        if (manualKeyUser2.trim()) formData.append("manualKeyUser2", manualKeyUser2.trim());
        if (manualKeyUser3.trim()) formData.append("manualKeyUser3", manualKeyUser3.trim());
        if (manualKeyUser4.trim()) formData.append("manualKeyUser4", manualKeyUser4.trim());
        if (manualKeyUser5.trim()) formData.append("manualKeyUser5", manualKeyUser5.trim());

        try {
            const res = await fetch("/api/upload", {
                method: "POST",
                body: formData,
            });
            const result = await res.json();

            if (!res.ok || result.error) {
                console.error("Upload failed:", result.error || res.statusText);
                setErrors([{
                    field: "System",
                    row: 0,
                    message: result.error || "Internal Server Error during data acquisition",
                    severity: "CRITICAL"
                }]);
                setIsProcessing(false);
                setCurrentStep("VALIDATE"); // Go to audit to show the error
                return;
            }

            setUploadData(result);
            setErrors(result.errors || []);

            // Simulate progress for dramatic effect
            setTimeout(() => {
                setIsProcessing(false);
                setCurrentStep("VALIDATE");
            }, 3000);

        } catch (err) {
            console.error(err);
            setIsProcessing(false);
            setCurrentStep("UPLOAD");
        }
    };

    const handleDownload = async (fileType: "orders" | "lines" | "sizes") => {
        try {
            console.log(`Initializing download for ${fileType}...`);
            if (!uploadData?.files?.[fileType]) {
                console.error("No file data found in uploadData");
                return;
            }

            const base64 = uploadData.files[fileType];
            const dataUrl = `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${base64}`;

            const res = await fetch(dataUrl);
            const blob = await res.blob();

            const url = window.URL.createObjectURL(blob);
            const link = document.createElement("a");
            link.href = url;
            link.download = `${fileType.toUpperCase()}_NG_TEMPLATE_${new Date().toISOString().split('T')[0]}.xlsx`;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            window.URL.revokeObjectURL(url);
            console.log(`${fileType} download triggered successfully`);
        } catch (err) {
            console.error("Download failed:", err);
            alert("Failed to generate download. Please try re-running the workflow.");
        }
    };

    const handleFileDownload = async (filename: string, fileType: "orders" | "lines" | "sizes") => {
        try {
            if (!uploadData?.fileOutputs?.[filename]?.[fileType]) {
                console.error("No per-file output data found", filename, fileType);
                return;
            }
            const base64 = uploadData.fileOutputs[filename][fileType];
            const dataUrl = `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${base64}`;
            const res = await fetch(dataUrl);
            const blob = await res.blob();
            const url = window.URL.createObjectURL(blob);
            const link = document.createElement("a");
            link.href = url;
            const safeName = filename.replace(/[^a-zA-Z0-9_-]/g, "_");
            link.download = `${safeName}_${fileType.toUpperCase()}.xlsx`;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            window.URL.revokeObjectURL(url);
        } catch (err) {
            console.error("Per-file download failed:", err);
            alert("Failed to generate per-file export. Please try again.");
        }
    };

    const blockerConditions = {
        statusConflict: errors.some((err) => /status.*(unconfirmed|confirmed)/i.test(err.message)),
        customerMappingGap: errors.some((err) => /9999996|customer.*mapping|customer.*code/i.test(err.message)),
        blankCriticalFields: errors.some((err) => /(blank|missing).*(price|delivery|payment)/i.test(err.message)),
        validationFailed: uploadData?.canProceed === false || errors.some((err) => err.severity === "CRITICAL"),
    };

    const blockerMessages = [
        blockerConditions.statusConflict && "Status conflict detected (Unconfirmed vs Confirmed).",
        blockerConditions.customerMappingGap && "Customer mapping gap detected (9999996 / Arcteryx mismatch).",
        blockerConditions.blankCriticalFields && "Blank pricing/delivery/payment fields found in export files.",
        blockerConditions.validationFailed && "Validation failure blocking progression. Resolve all critical errors first.",
    ].filter(Boolean);

    return (
        <div className="w-full max-w-7xl mx-auto space-y-20 px-4 transition-colors duration-300 text-[hsl(var(--foreground))]">
            {/* Progress Stepper - Redesigned */}
            <div className="relative pt-12">
                <div className="absolute top-[calc(3rem+28px)] left-[10%] w-[80%] h-px bg-white/10 progress-line" />
                <div
                    className="absolute top-[calc(3rem+28px)] left-[10%] h-0.5 bg-gradient-to-r from-blue-600 to-indigo-500 shadow-[0_0_20px_rgba(37,99,235,0.4)] transition-all duration-1000 ease-in-out"
                    style={{ width: `${(currentStepIndex / (steps.length - 1)) * 80}%` }}
                />

                <div className="relative flex justify-between max-w-5xl mx-auto">
                    {steps.map((step, index) => {
                        const Icon = step.icon;
                        const isActive = currentStep === step.key;
                        const isCompleted = currentStepIndex > index;

                        return (
                            <button
                                key={step.key}
                                onClick={() => {
                                    if (isCompleted || isActive) setCurrentStep(step.key);
                                }}
                                disabled={!isCompleted && !isActive}
                                className={cn(
                                    "flex flex-col items-center group relative z-10",
                                    (isCompleted || isActive) ? "cursor-pointer" : "cursor-default"
                                )}
                            >
                                <div
                                    className={cn(
                                        "w-16 h-16 rounded-2xl flex items-center justify-center transition-all duration-700 ease-out",
                                        isActive ? "bg-blue-600 shadow-[0_0_40px_rgba(37,99,235,0.4)] scale-110 ring-4 ring-blue-500/20" :
                                            isCompleted ? "bg-emerald-600 shadow-[0_0_30px_rgba(16,185,129,0.2)] scale-100" : "bg-slate-900 border border-white/5 scale-90"
                                    )}
                                >
                                    {isActive && isProcessing ? (
                                        <Loader2 className="w-7 h-7 animate-[spin_2s_linear_infinite] text-white" />
                                    ) : (
                                        <Icon className={cn("w-7 h-7 transition-colors duration-500", (isActive || isCompleted) ? "text-white" : "text-slate-600 translate-y-0")} />
                                    )}

                                    {isCompleted && (
                                        <motion.div
                                            initial={{ scale: 0 }}
                                            animate={{ scale: 1 }}
                                            className="absolute -top-1 -right-1 bg-emerald-400 text-emerald-950 rounded-full p-0.5 border-2 border-slate-950"
                                        >
                                            <CheckCircle2 className="w-3.5 h-3.5 fill-current" />
                                        </motion.div>
                                    )}
                                </div>
                                <span className={cn(
                                    "step-label mt-6 text-[10px] font-black uppercase tracking-[0.3em] transition-all duration-700",
                                    isActive ? "text-blue-400 opacity-100 translate-y-0" : isCompleted ? "text-emerald-400 opacity-80" : "text-slate-700 opacity-50"
                                )}>
                                    {step.label}
                                </span>
                            </button>
                        );
                    })}
                </div>
            </div>

            {/* Content Stage */}
            <motion.div
                layout
                className="glass-panel rounded-[48px] relative overflow-hidden p-10 md:p-20 min-h-[600px] flex items-center justify-center border-white/[0.08]"
            >
                <div className="scanline" />

                {/* Stage-specific subtle glow */}
                <div className={cn(
                    "absolute top-0 left-0 w-full h-1 bg-gradient-to-r transition-all duration-1000",
                    currentStep === "UPLOAD" ? "from-blue-500/0 via-blue-500/50 to-blue-500/0" :
                        currentStep === "RUN" ? "from-indigo-500/0 via-indigo-500/50 to-indigo-500/0" :
                            currentStep === "VALIDATE" ? "from-amber-500/0 via-amber-500/50 to-amber-500/0" :
                                "from-emerald-500/0 via-emerald-500/50 to-emerald-500/0"
                )} />

                <AnimatePresence mode="wait">
                    {currentStep === "UPLOAD" && (
                        <motion.div
                            key="upload"
                            initial={{ opacity: 0, y: 30 }}
                            animate={{ opacity: 1, y: 0 }}
                            exit={{ opacity: 0, y: -30 }}
                            className="text-center max-w-3xl mx-auto space-y-12"
                        >
                            <div className="relative inline-block group">
                                <div className="absolute inset-0 bg-blue-500/20 blur-[80px] rounded-full group-hover:bg-blue-500/30 transition-all duration-500" />
                                <div className="relative w-40 h-40 bg-[hsl(var(--panel))] border border-[hsl(var(--border))] rounded-[38%] flex items-center justify-center mx-auto transition-all duration-500 group-hover:scale-105 group-hover:border-blue-500/50 group-hover:shadow-[0_0_50px_rgba(59,130,246,0.3)] shadow-2xl">
                                    <CloudUpload className="w-16 h-16 text-blue-500 group-hover:text-blue-400 group-hover:-translate-y-1 transition-all" />
                                </div>
                            </div>

                            <div className="space-y-6">
                                <h2 className="text-5xl font-black tracking-tight text-[hsl(var(--foreground))] leading-tight">
                                    INITIALIZE <br /> <span className="bg-clip-text text-transparent bg-gradient-to-r from-blue-500 via-sky-400 to-indigo-500">SYSTEM ACQUISITION</span>
                                </h2>
                                <p className="text-[hsl(var(--muted))] text-xl font-medium max-w-xl mx-auto">
                                    Upload your buy file and optional product sheet. Use both for PLM-enriched outputs.
                                </p>
                            </div>

                            <div className="grid grid-cols-1 gap-4 max-w-3xl mx-auto text-left">
                                <div className="space-y-2">
                                    <label className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-500">Buy File (.xlsx) — Required</label>
                                    <input
                                        type="file"
                                        accept=".xlsx"
                                        multiple
                                        onChange={(e) => setBuyFiles(e.target.files)}
                                        className="w-full rounded-xl bg-white/5 border border-white/10 px-4 py-3 text-sm text-white file-input"
                                    />
                                    {buyFiles && buyFiles.length > 0 && (
                                        <div className="text-[10px] text-slate-400">
                                            {Array.from(buyFiles).map((f) => f.name).join(", ")}
                                        </div>
                                    )}
                                </div>
                                <div className="space-y-2">
                                    <label className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-500">Product Sheet / PLM (.xlsx) — Optional</label>
                                    <input
                                        type="file"
                                        accept=".xlsx"
                                        onChange={(e) => setProductSheetFile(e.target.files?.[0] || null)}
                                        className="w-full rounded-xl bg-white/5 border border-white/10 px-4 py-3 text-sm text-white file-input"
                                    />
                                    {productSheetFile && (
                                        <div className="text-[10px] text-slate-400">{productSheetFile.name}</div>
                                    )}
                                </div>
                                <div className="space-y-2">
                                    <label className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-500">Manual PO</label>
                                    <input
                                        value={manualPo}
                                        onChange={(e) => setManualPo(e.target.value)}
                                        placeholder="PO002954"
                                        className="w-full rounded-xl bg-white/5 border border-white/10 px-4 py-3 text-sm text-white placeholder:text-slate-600 focus:outline-none focus:ring-2 focus:ring-blue-500/40"
                                    />
                                </div>
                                <div className="space-y-2">
                                    <label className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-500">Template</label>
                                    <input
                                        value={manualTemplate}
                                        onChange={(e) => setManualTemplate(e.target.value)}
                                        placeholder="FOB Bulk EDI PO (New) or SMS EDI"
                                        className="w-full rounded-xl bg-white/5 border border-white/10 px-4 py-3 text-sm text-white placeholder:text-slate-600 focus:outline-none focus:ring-2 focus:ring-blue-500/40"
                                    />
                                </div>
                                <div className="space-y-2">
                                    <label className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-500">Comments</label>
                                    <input
                                        value={manualComments}
                                        onChange={(e) => setManualComments(e.target.value)}
                                        placeholder="[TNF] FW25 Nov Buy 15-NOV Bulk"
                                        className="w-full rounded-xl bg-white/5 border border-white/10 px-4 py-3 text-sm text-white placeholder:text-slate-600 focus:outline-none focus:ring-2 focus:ring-blue-500/40"
                                    />
                                </div>
                                <div className="space-y-2">
                                    <label className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-500">Orders KeyDate</label>
                                    <input
                                        value={manualKeyDate}
                                        onChange={(e) => setManualKeyDate(e.target.value)}
                                        placeholder="6/12/2026"
                                        className="w-full rounded-xl bg-white/5 border border-white/10 px-4 py-3 text-sm text-white placeholder:text-slate-600 focus:outline-none focus:ring-2 focus:ring-blue-500/40"
                                    />
                                </div>
                                <div className="space-y-2">
                                    <label className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-500">KeyUser1 (Manual)</label>
                                    <input
                                        value={manualKeyUser1}
                                        onChange={(e) => setManualKeyUser1(e.target.value)}
                                        placeholder="Planning"
                                        className="w-full rounded-xl bg-white/5 border border-white/10 px-4 py-3 text-sm text-white placeholder:text-slate-600 focus:outline-none focus:ring-2 focus:ring-blue-500/40"
                                    />
                                </div>
                                <div className="space-y-2">
                                    <label className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-500">KeyUser2 (Manual)</label>
                                    <input
                                        value={manualKeyUser2}
                                        onChange={(e) => setManualKeyUser2(e.target.value)}
                                        placeholder="Purchasing"
                                        className="w-full rounded-xl bg-white/5 border border-white/10 px-4 py-3 text-sm text-white placeholder:text-slate-600 focus:outline-none focus:ring-2 focus:ring-blue-500/40"
                                    />
                                </div>
                                <div className="space-y-2">
                                    <label className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-500">KeyUser3 (Manual)</label>
                                    <input
                                        value={manualKeyUser3}
                                        onChange={(e) => setManualKeyUser3(e.target.value)}
                                        placeholder="(Optional)"
                                        className="w-full rounded-xl bg-white/5 border border-white/10 px-4 py-3 text-sm text-white placeholder:text-slate-600 focus:outline-none focus:ring-2 focus:ring-blue-500/40"
                                    />
                                </div>
                                <div className="space-y-2">
                                    <label className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-500">KeyUser4 (Manual)</label>
                                    <input
                                        value={manualKeyUser4}
                                        onChange={(e) => setManualKeyUser4(e.target.value)}
                                        placeholder="Production"
                                        className="w-full rounded-xl bg-white/5 border border-white/10 px-4 py-3 text-sm text-white placeholder:text-slate-600 focus:outline-none focus:ring-2 focus:ring-blue-500/40"
                                    />
                                </div>
                                <div className="space-y-2">
                                    <label className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-500">KeyUser5 (Manual)</label>
                                    <input
                                        value={manualKeyUser5}
                                        onChange={(e) => setManualKeyUser5(e.target.value)}
                                        placeholder="Logistics"
                                        className="w-full rounded-xl bg-white/5 border border-white/10 px-4 py-3 text-sm text-white placeholder:text-slate-600 focus:outline-none focus:ring-2 focus:ring-blue-500/40"
                                    />
                                </div>
                            </div>

                            <div className="flex flex-col items-center gap-6">
                                <button
                                    onClick={handleStartUpload}
                                    disabled={!buyFiles || buyFiles.length === 0}
                                    className="primary-button inline-flex items-center gap-4 bg-blue-600 text-white disabled:opacity-50 disabled:cursor-not-allowed"
                                    style={{ background: "linear-gradient(90deg, #2563eb, #1d4ed8)" }}
                                >
                                    <span>START UPLOAD</span>
                                    <ArrowRight className="w-4 h-4" />
                                </button>

                                <div className="flex items-center gap-10 opacity-70">
                                    {['exceljs', 'validation', 'supabase'].map((tech, i) => (
                                        <div key={tech} className="flex items-center gap-2.5 text-[10px] font-black tracking-[0.3em] uppercase text-[hsl(var(--muted))]" style={{ opacity: i === 0 ? 1 : 0.85 }}>
                                            <div className="w-2 h-2 rounded-full bg-blue-500" /> {tech}
                                        </div>
                                    ))}
                                </div>
                            </div>
                        </motion.div>
                    )}

                    {currentStep === "RUN" && (
                        <motion.div
                            key="run"
                            initial={{ opacity: 0, scale: 0.95 }}
                            animate={{ opacity: 1, scale: 1 }}
                            className="text-center space-y-12"
                        >
                            <div className="relative w-32 h-32 mx-auto">
                                <div className="absolute inset-0 bg-blue-500/20 blur-3xl animate-pulse" />
                                <Loader2 className="w-full h-full text-blue-500 animate-[spin_4s_linear_infinite]" />
                                <div className="absolute inset-0 flex items-center justify-center">
                                    <div className="w-2 h-2 bg-white rounded-full shadow-[0_0_15px_#fff]" />
                                </div>
                            </div>

                            <div className="space-y-6">
                                <h3 className="text-3xl font-black tracking-[0.2em] uppercase text-white">Executing Engine</h3>
                                <div className="flex flex-col items-center gap-3">
                                    <div className="px-6 py-2 bg-white/5 border border-white/5 rounded-full">
                                        <p className="text-slate-500 text-[10px] font-black uppercase tracking-widest">Status: <span className="text-blue-400">Transforming Data Nodes</span></p>
                                    </div>
                                    <div className="w-64 h-1 bg-white/5 rounded-full overflow-hidden">
                                        <motion.div
                                            initial={{ x: "-100%" }}
                                            animate={{ x: "100%" }}
                                            transition={{ repeat: Infinity, duration: 1.5, ease: "linear" }}
                                            className="w-1/2 h-full bg-blue-500"
                                        />
                                    </div>
                                    <p className="text-slate-600 text-[9px] font-mono tracking-tighter opacity-70">NODE_RESOLVER :: BUSS_LOGIC_V4 :: OK</p>
                                </div>
                            </div>
                        </motion.div>
                    )}

                    {currentStep === "VALIDATE" && (
                        <motion.div
                            key="validate"
                            initial={{ opacity: 0, y: 20 }}
                            animate={{ opacity: 1, y: 0 }}
                            className="w-full h-full flex flex-col"
                        >
                            <div className="flex flex-col md:flex-row items-center justify-between mb-12 gap-8">
                                <div className="flex items-center gap-6">
                                    <div className="p-4 bg-amber-500/10 rounded-3xl border border-amber-500/20 shadow-[0_0_30px_rgba(245,158,11,0.1)]">
                                        <ShieldCheck className="w-8 h-8 text-amber-500" />
                                    </div>
                                    <div>
                                        <h2 className="text-3xl font-black tracking-tight text-white uppercase">Data Audit Node</h2>
                                        <p className="text-[10px] text-slate-500 font-black uppercase tracking-[0.3em] mt-1">Integrity verification protocol complete</p>
                                    </div>
                                </div>
                                <div className="flex flex-row items-center gap-4">
                                    <button
                                        onClick={() => setCurrentStep("UPLOAD")}
                                        className="inline-flex items-center gap-2 text-slate-500 hover:text-white transition-colors text-[10px] font-black uppercase tracking-widest mr-4 group"
                                    >
                                        <History className="w-4 h-4 group-hover:-rotate-90 transition-transform" /> RE-ACQUIRE SOURCE
                                    </button>
                                    <div className="flex flex-col items-end gap-2">
                                        <button
                                            onClick={() => setCurrentStep("REVIEW")}
                                            disabled={errors.some(e => e.severity === "CRITICAL")}
                                            className={cn(
                                                "primary-button flex items-center gap-4 group",
                                                errors.some(e => e.severity === "CRITICAL") && "opacity-50 cursor-not-allowed grayscale"
                                            )}
                                        >
                                            <span>COMMIT TO REVIEW</span> <ChevronRight className="w-5 h-5 group-hover:translate-x-1 transition-transform" />
                                        </button>
                                        {errors.some(e => e.severity === "CRITICAL") && (
                                            <p className="text-[9px] font-black text-red-500 uppercase tracking-widest mt-2 animate-pulse">
                                                [CRITICAL ERRORS DETECTED] RESOLVE SOURCE FILE DATA
                                            </p>
                                        )}
                                    </div>
                                </div>
                            </div>

                            <div className="glass-card rounded-[32px] overflow-hidden flex-1 border-white/5 bg-slate-950/40 relative">
                                <div className="overflow-x-auto">
                                    <table className="w-full text-left border-collapse">
                                        <thead>
                                            <tr className="border-b border-white/10 bg-white/[0.02]">
                                                <th className="px-10 py-7 text-[10px] font-black uppercase tracking-[0.3em] text-slate-500">Attribute</th>
                                                <th className="px-10 py-7 text-[10px] font-black uppercase tracking-[0.3em] text-slate-500 text-center">Reference</th>
                                                <th className="px-10 py-7 text-[10px] font-black uppercase tracking-[0.3em] text-slate-500 text-center">Intensity</th>
                                                <th className="px-10 py-7 text-[10px] font-black uppercase tracking-[0.3em] text-slate-500">Diagnostic Message</th>
                                            </tr>
                                        </thead>
                                        <tbody className="divide-y divide-white/5">
                                            {errors.length > 0 ? errors.map((err, i) => (
                                                <tr key={i} className="hover:bg-white/[0.03] transition-all group">
                                                    <td className="px-10 py-8">
                                                        <div className="flex items-center gap-3">
                                                            <div className="w-2 h-2 rounded-full bg-blue-500/40" />
                                                            <span className="font-black text-xs tracking-tight text-slate-200 uppercase">{err.field}</span>
                                                        </div>
                                                    </td>
                                                    <td className="px-10 py-8">
                                                        <div className="flex items-center justify-center">
                                                            <span className="text-[10px] font-black bg-slate-900 border border-white/5 px-3 py-1.5 rounded-lg text-slate-400 font-mono tracking-tighter uppercase">ROW_{err.row}</span>
                                                        </div>
                                                    </td>
                                                    <td className="px-10 py-8">
                                                        <div className="flex items-center justify-center">
                                                            <span className={cn(
                                                                "px-4 py-1.5 rounded-xl text-[9px] font-black uppercase tracking-widest border",
                                                                err.severity === "CRITICAL" ? "bg-red-500/10 text-red-500 border-red-500/20 shadow-[0_0_20px_rgba(239,68,68,0.1)]" : "bg-amber-500/10 text-amber-500 border-amber-500/20 shadow-[0_0_20px_rgba(245,158,11,0.1)]"
                                                            )}>
                                                                {err.severity}
                                                            </span>
                                                        </div>
                                                    </td>
                                                    <td className="px-10 py-8 text-slate-400 text-xs font-medium leading-relaxed italic">{err.message}</td>
                                                </tr>
                                            )) : (
                                                <tr>
                                                    <td colSpan={4} className="px-10 py-32 text-center">
                                                        <motion.div
                                                            initial={{ opacity: 0, scale: 0.9 }}
                                                            animate={{ opacity: 1, scale: 1 }}
                                                            className="flex flex-col items-center gap-6"
                                                        >
                                                            <div className="w-20 h-20 rounded-3xl bg-emerald-500/10 border border-emerald-500/20 flex items-center justify-center">
                                                                <CheckCircle2 className="w-10 h-10 text-emerald-500" />
                                                            </div>
                                                            <div className="space-y-2">
                                                                <p className="text-xl font-black tracking-[0.2em] uppercase text-white">Dataset Sanitized</p>
                                                                <p className="text-[10px] font-black tracking-widest uppercase text-slate-600">No anomalous patterns detected in source</p>
                                                            </div>
                                                        </motion.div>
                                                    </td>
                                                </tr>
                                            )}
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                        </motion.div>
                    )}

                    {currentStep === "REVIEW" && (
                        <motion.div
                            key="review"
                            initial={{ opacity: 0, scale: 0.95 }}
                            animate={{ opacity: 1, scale: 1 }}
                            className="text-center max-w-2xl mx-auto space-y-16"
                        >
                            <div className="relative mx-auto w-40 h-40">
                                <div className="absolute inset-0 bg-emerald-500/20 blur-[60px] animate-pulse rounded-full" />
                                <div className="relative w-full h-full bg-slate-900 border border-emerald-500/30 rounded-[35%] flex items-center justify-center shadow-2xl">
                                    <FileCheck className="w-20 h-20 text-emerald-400 transition-transform group-hover:scale-110" />
                                </div>
                            </div>

                            <div className="space-y-6">
                                <h2 className="text-5xl font-black tracking-tighter uppercase text-white">Transformation <br /> <span className="text-emerald-500">Confirmed</span></h2>
                                <div className="flex items-center justify-center gap-6">
                                    <div className="bg-white/5 border border-white/10 rounded-2xl px-6 py-4 backdrop-blur-md">
                                        <p className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-500 mb-1">Entity Count</p>
                                        <p className="text-2xl font-black text-blue-400">{uploadData?.mergedSummary?.orders || uploadData?.dataCount || '---'} <span className="text-[10px] text-slate-600">HEADERS</span></p>
                                    </div>
                                    <div className="bg-white/5 border border-white/10 rounded-2xl px-6 py-4 backdrop-blur-md">
                                        <p className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-500 mb-1">Standard</p>
                                        <p className="text-2xl font-black text-emerald-500">NG <span className="text-[10px] text-slate-600">COMPLIANT</span></p>
                                    </div>
                                </div>

                                {uploadData?.fileSummary?.length > 0 && (
                                    <div className="bg-slate-900/60 border border-white/10 rounded-2xl p-4 text-left text-xs text-slate-300 file-summary-panel">
                                        <div className="font-black text-white uppercase tracking-[0.2em] text-[10px] mb-2 file-summary-title">File-level Summary</div>
                                        <div className="grid grid-cols-2 md:grid-cols-6 gap-2 text-[10px]">
                                            <div className="font-semibold text-slate-200 file-summary-head">File</div>
                                            <div className="font-semibold text-slate-200 file-summary-head">Orders</div>
                                            <div className="font-semibold text-slate-200 file-summary-head">Lines</div>
                                            <div className="font-semibold text-slate-200 file-summary-head">Sizes</div>
                                            <div className="font-semibold text-slate-200 file-summary-head">Errors</div>
                                            <div className="font-semibold text-slate-200 file-summary-head">Warnings</div>
                                                {uploadData.fileSummary.map((f:any, idx:number) => (
                                                <div key={`file-summary-${idx}`} className="contents">
                                                    <div className="truncate">
                                                        <div className="font-black file-summary-file">{f.filename}</div>
                                                        {Array.isArray(f.brands) && f.brands.length > 0 && (
                                                            <div className="mt-1 flex flex-wrap gap-1">
                                                                {f.brands.map((b: string) => (
                                                                    <span key={`${f.filename}-${b}`} className="file-summary-badge">
                                                                        {b}
                                                                    </span>
                                                                ))}
                                                            </div>
                                                        )}
                                                    </div>
                                                    <div className="font-black text-emerald-300">{f.orders}</div>
                                                    <div>{f.lines}</div>
                                                    <div>{f.sizes}</div>
                                                    <div className="text-red-400">{f.errors}</div>
                                                    <div className="text-amber-300">{f.warnings}</div>
                                                </div>
                                            ))}
                                        </div>
                                    </div>
                                )}

                                <div className="quick-qa-panel bg-white/5 border border-white/10 rounded-2xl px-6 py-4 backdrop-blur-md">
                                    <div className="font-black text-[10px] uppercase tracking-[0.3em] text-slate-500 mb-2 quick-qa-label">Quick QA Summary</div>
                                    <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 text-[11px]">
                                        <div>
                                            <div className="text-slate-400 quick-qa-label">PO Count</div>
                                            <div className="text-emerald-300 font-black quick-qa-value">{uploadData?.mergedSummary?.orders || 0}</div>
                                        </div>
                                        <div>
                                            <div className="text-slate-400 quick-qa-label">Line Count</div>
                                            <div className="text-blue-300 font-black quick-qa-value">{uploadData?.mergedSummary?.lines || 0}</div>
                                        </div>
                                        <div>
                                            <div className="text-slate-400 quick-qa-label">Size Total</div>
                                            <div className="text-amber-300 font-black quick-qa-value">{uploadData?.mergedSummary?.sizes || 0}</div>
                                        </div>
                                    </div>
                                </div>

                                {blockerMessages.length > 0 && (
                                    <div className="bg-rose-950/70 border border-rose-400/30 rounded-2xl p-4 text-left text-xs text-rose-200">
                                        <div className="font-black text-rose-200 uppercase tracking-[0.2em] text-[10px] mb-2">Critical Blockers (resolve before cutting)</div>
                                        <ul className="list-disc list-inside space-y-1">
                                            {blockerMessages.map((msg, i) => (
                                                <li key={`blocker-${i}`} className="text-[11px]">{msg}</li>
                                            ))}
                                        </ul>
                                        <p className="mt-2 text-[10px] text-rose-300">Confirm authoritative status source and customer code mappings before exporting.</p>
                                    </div>
                                )}
                            </div>

                            <div className="flex flex-col items-center gap-8">
                                <button
                                    onClick={() => setCurrentStep("DOWNLOAD")}
                                    className="primary-button w-full flex items-center justify-center gap-4 text-sm py-6 group"
                                >
                                    <span>INITIALIZE TEMPLATE GENERATION</span> <Download className="w-6 h-6 group-hover:-translate-y-1 transition-transform" />
                                </button>
                                <button
                                    onClick={() => setCurrentStep("VALIDATE")}
                                    className="flex items-center gap-3 text-slate-600 hover:text-white transition-all text-[11px] font-black tracking-[0.4em] uppercase group"
                                >
                                    <ChevronLeft className="w-5 h-5 group-hover:-translate-x-1 transition-transform" /> BACK TO AUDIT DATA
                                </button>
                            </div>
                        </motion.div>
                    )}

                    {currentStep === "DOWNLOAD" && (
                        <motion.div
                            key="download"
                            initial={{ opacity: 0, y: 30 }}
                            animate={{ opacity: 1, y: 0 }}
                            className="w-full space-y-16"
                        >
                            <div className="text-center space-y-4">
                                <h2 className="text-4xl font-black uppercase tracking-tighter text-white">Repository Synchronization</h2>
                                <p className="text-slate-500 text-lg font-medium opacity-80">Download automated templates for NextGen high-speed ingestion.</p>
                            </div>

                            <div className="grid grid-cols-1 md:grid-cols-3 gap-8">
                                {[
                                    { key: 'orders', label: 'ORDERS' },
                                    { key: 'lines', label: 'LINES' },
                                    { key: 'sizes', label: 'ORDER_SIZES' }
                                ].map((file, i) => (
                                    <motion.div
                                        key={file.key}
                                        initial={{ opacity: 0, scale: 0.9 }}
                                        animate={{ opacity: 1, scale: 1 }}
                                        transition={{ delay: i * 0.1 }}
                                        className="glass-card rounded-[40px] p-10 flex flex-col items-center group relative overflow-hidden border-white/[0.08]"
                                    >
                                        <div className="absolute top-0 right-0 w-40 h-40 bg-blue-500/5 blur-[50px] -z-10 group-hover:bg-blue-500/10 transition-colors" />
                                        <div className="w-20 h-20 bg-blue-500/10 rounded-[30%] flex items-center justify-center mb-10 border border-blue-500/20 group-hover:rotate-6 transition-transform shadow-lg group-hover:shadow-blue-500/20 group-hover:border-blue-500/50">
                                            <FileText className="w-10 h-10 text-blue-400 group-hover:text-blue-300" />
                                        </div>
                                        <h4 className="font-black text-[11px] tracking-[0.3em] mb-2 text-white">{file.label}</h4>
                                        <p className="text-[9px] text-slate-500 mb-10 uppercase font-black tracking-widest">NG_UPLOAD_BLOB .XLSX</p>
                                        <button
                                            onClick={() => handleDownload(file.key as any)}
                                            className="secondary-button w-full py-4 text-[10px] tracking-[0.3em] group/btn overflow-hidden relative download-package-btn"
                                        >
                                            <span className="relative z-10">DOWNLOAD PACKAGE</span>
                                            <div className="absolute inset-0 bg-blue-500 translate-y-full group-hover/btn:translate-y-0 transition-transform duration-300" />
                                        </button>
                                    </motion.div>
                                ))}
                            </div>

                            {uploadData?.fileOutputs && (
                                <div className="bg-slate-900/60 border border-white/10 rounded-2xl p-4 text-left text-xs text-slate-300 file-export-panel">
                                    <div className="font-black text-white uppercase tracking-[0.2em] text-[10px] mb-2 file-export-title">Per-file Template Export</div>
                                    <div className="space-y-3">
                                        {Object.entries(uploadData.fileOutputs).map(([fname, payload]: any) => (
                                            <div key={fname} className="flex flex-wrap items-center gap-3 bg-slate-800/40 border border-white/10 rounded-lg px-3 py-2 file-export-row">
                                                <div className="text-[11px] font-black text-emerald-300 truncate max-w-[240px] file-export-name">{fname}</div>
                                                {['orders','lines','sizes'].map((type) => (
                                                    <button
                                                        key={`${fname}-${type}`}
                                                        onClick={() => handleFileDownload(fname, type as any)}
                                                        className="px-2 py-1 text-[10px] uppercase tracking-[0.2em] font-black border border-slate-500 rounded-md bg-blue-500/15 hover:bg-blue-500/35 file-export-btn"
                                                    >
                                                        {type.toUpperCase()}
                                                    </button>
                                                ))}
                                            </div>
                                        ))}
                                    </div>
                                </div>
                            )}
                            <div className="flex flex-col items-center gap-6">
                                <div className="flex justify-center gap-10">
                                    <button
                                        onClick={() => setCurrentStep("REVIEW")}
                                        className="flex items-center gap-3 text-slate-600 hover:text-white transition-all text-[11px] font-black tracking-[0.4em] uppercase group"
                                    >
                                        <ChevronLeft className="w-5 h-5 group-hover:-translate-x-1 transition-transform" /> BACK TO REVIEW
                                    </button>
                                    <button
                                        onClick={() => setCurrentStep("UPLOAD")}
                                        className="flex items-center gap-3 text-slate-600 hover:text-white transition-all text-[11px] font-black tracking-[0.4em] uppercase group"
                                    >
                                        <History className="w-5 h-5 group-hover:-rotate-90 transition-transform duration-500" /> RESET ENGINE WORKFLOW
                                    </button>
                                </div>
                            </div>
                        </motion.div>
                    )}
                </AnimatePresence>
            </motion.div>
        </div>
    );
}
