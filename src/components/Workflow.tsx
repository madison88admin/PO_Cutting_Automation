"use client";

// Predefined comment options for dropdown
const COMMENT_OPTIONS = [
    "[TNF] FW25 Nov Buy 15-NOV Bulk",
    "[Vans] SS26 Bulk Order",
    "[Columbia] FW26 Special",
    "[Vuori] 2026 Main Buy",
    "[Marmot] Fall 2026",
    "[Other]"
];

import { useEffect, useRef, useState } from "react";
import { Upload, FileCheck, AlertCircle, Download, ChevronRight, ChevronLeft, Settings, History, Loader2, CheckCircle2, CloudUpload, ArrowRight, ShieldCheck, FileText, X, Copy, Eye, PackageCheck } from "lucide-react";
import { saveTemplate } from "@/lib/templates/template-store";
import { clsx, type ClassValue } from "clsx";
import { twMerge } from "tailwind-merge";
import { motion, AnimatePresence } from "framer-motion";

function cn(...inputs: ClassValue[]) {
    return twMerge(clsx(inputs));
}

type Step = "UPLOAD" | "RUN" | "VALIDATE" | "REVIEW" | "DOWNLOAD";

const TEMPLATE_OPTIONS = [
    "Major Brand Bulk",
    "Major Brand Bulk (COPY)",
    "SMS PO Header",
    "Specialty Bulk PO",
];

const LINE_TEMPLATE_OPTIONS = [
    "FOB Bulk EDI PO (New)",
    "FOB Bulk EDI PO (VF & UA) (New)",
    "FOB Bulk Non EDI PO (New)",
    "Forecast",
    "Forecast - CRD",
    "Forecast - OP",
    "SMS EDI (New)",
    "SMS Keep Samples",
    "SMS Non EDI",
    "SMS Non EDI (New)",
];

const SEASON_OPTIONS = [
    "FH:2026",
    "FH:2027",
    "SH:2026",
    "SH:2027",
    "FW26",
    "FW27",
    "SS26",
    "SS27",
    "AW26",
    "AW27",
];

const BRAND_OPTIONS = [
    "Vans",
    "Prana",
    "Columbia",
    "Rossignol",
    "Peak Performance",
    "Dynafit",
    "LL Bean",
    "Helly Hansen",
    "Jack Wolfskin",
    "Vuori",
    "Marmot",
    "Burton",
    "Cotopaxi",
    "Fox Racing",
    "Haglofs",
    "Mammut",
    "Evo",
    "On AG",
    "66 Degrees North",
    "Hunter",
    "Other",
];

const CUSTOMER_OPTIONS = [
    "Vans",
    "Prana",
    "Columbia",
    "The North Face In-Line",
    "The North Face SMU",
    "The North Face RTO",
    "Rossignol",
    "Peak Performance",
    "Dynafit",
    "LL Bean",
    "Helly Hansen",
    "Jack Wolfskin",
    "Vuori",
    "Marmot",
    "Burton",
    "Cotopaxi",
    "Fox Racing",
    "Haglofs",
    "Mammut",
    "Evo",
    "On AG",
    "66 Degrees North",
    "Hunter",
    "Other",
];

const DESTINATION_OPTIONS = [
    "USA",
    "Canada",
    "Mexico",
    "UK",
    "UAE",
    "France",
    "Germany",
    "Korea",
    "Japan",
    "Indonesia",
    "Australia",
    "Hong Kong",
    "Netherlands",
    "Czech Republic",
    "New Zealand",
    "Sweden",
    "Iceland",
    "EU",
    "Belgium",
    "Other",
];

const HEADER_FIELD_OPTIONS = [
    ["buyer_style_number", "Buyer style / article"],
    ["buyer_style_name", "Style name"],
    ["sku", "SKU / material"],
    ["product_description", "Product description"],
    ["color", "Colour name"],
    ["color_code", "Colour code"],
    ["size", "Size"],
    ["quantity", "Quantity"],
    ["delivery_date", "Delivery / ship date"],
    ["season", "Season"],
    ["customer", "Customer / brand"],
    ["factory", "Factory / supplier"],
    ["currency", "Currency"],
    ["unit_cost", "Unit cost"],
    ["po_number", "PO number"],
    ["buyer_po_number", "Buyer PO number"],
    ["start_date", "Start date"],
    ["cancel_date", "Cancel date"],
    ["transport_method", "Transport method"],
] as const;

function inferSeasonFromFilename(filename: string): string {
    const text = filename.toUpperCase();
    const explicitMatch = text.match(/\b((?:FW|FH|AW|AH|SS|SH|SP|SW)\s*\d{2})\b/);
    if (explicitMatch?.[1]) {
        return explicitMatch[1].replace(/\s+/g, "");
    }
    const shortMatch = text.match(/\b([FS])\s*(\d{2})\b/);
    if (shortMatch) {
        return `${shortMatch[1]}W${shortMatch[2]}`;
    }
    return "";
}

export default function Workflow() {
    const [currentStep, setCurrentStep] = useState<Step>("UPLOAD");
    const [isProcessing, setIsProcessing] = useState(false);
    const [errors, setErrors] = useState<any[]>([]);
    const [uploadData, setUploadData] = useState<any>(null);
    const [nextgenValidation, setNextgenValidation] = useState<any>(null);
    const [isValidatingNextgen, setIsValidatingNextgen] = useState(false);
    const [ocrFile, setOcrFile] = useState<File | null>(null);
    const [ocrResults, setOcrResults] = useState<any[] | null>(null);
    const [ocrNextgenUsed, setOcrNextgenUsed] = useState(false);
    const [isProcessingOcr, setIsProcessingOcr] = useState(false);
    const [buyFiles, setBuyFiles] = useState<FileList | null>(null);
    const [manualPo, setManualPo] = useState("");
    const manualPoRef = useRef(manualPo);
    useEffect(() => { manualPoRef.current = manualPo; }, [manualPo]);
    const [extractedPo, setExtractedPo] = useState("");
    const [manualTemplate, setManualTemplate] = useState("");
    const [manualLinesTemplate, setManualLinesTemplate] = useState("");
    const [manualComments, setManualComments] = useState("");
    const [customComment, setCustomComment] = useState("");
    const [manualKeyDate, setManualKeyDate] = useState("");
    const [manualKeyUser1, setManualKeyUser1] = useState("");
    const [manualKeyUser2, setManualKeyUser2] = useState("");
    const [manualKeyUser3, setManualKeyUser3] = useState("");
    const [manualKeyUser4, setManualKeyUser4] = useState("");
    const [manualKeyUser5, setManualKeyUser5] = useState("");
    const [manualSeason, setManualSeason] = useState("");
    const [manualCustomer, setManualCustomer] = useState("");
    const [manualBrand, setManualBrand] = useState("");
    const [manualDestination, setManualDestination] = useState("");
    const [showAdvanced, setShowAdvanced] = useState(false);
    const [isDragging, setIsDragging] = useState(false);
    const [processingStage, setProcessingStage] = useState(0);
    const [acknowledgedLines, setAcknowledgedLines] = useState<Record<string, boolean>>({});
    const [headerPreviews, setHeaderPreviews] = useState<any[]>([]);
    const [isPreviewingHeaders, setIsPreviewingHeaders] = useState(false);
    const [headerPreviewError, setHeaderPreviewError] = useState("");
    const [nextgenOverrides, setNextgenOverrides] = useState<Record<string, any>>({});
    const fileInputRef = useRef<HTMLInputElement | null>(null);

    const processingStages = [
        "Reading Excel sheets and headers",
        "Extracting orders, lines, and sizes",
        "Searching products and colours in NextGen",
        "Checking matches and required fields",
        "Preparing the three Excel output files",
    ];

    useEffect(() => {
        if (!isProcessing) {
            setProcessingStage(0);
            return;
        }
        const timer = window.setInterval(() => {
            setProcessingStage((stage) => Math.min(stage + 1, processingStages.length - 1));
        }, 2600);
        return () => window.clearInterval(timer);
    }, [isProcessing, processingStages.length]);

    useEffect(() => {
        try {
            const saved = JSON.parse(window.localStorage.getItem("po-cutting-form-preferences") || "{}");
            const preference = saved.last || {};
            if (!manualTemplate && preference.ordersTemplate) setManualTemplate(preference.ordersTemplate);
            if (!manualLinesTemplate && preference.linesTemplate) setManualLinesTemplate(preference.linesTemplate);
            if (!manualComments && preference.comments) setManualComments(preference.comments);
        } catch {
            // A malformed local preference should never block PO processing.
        }
    }, []);

    const rememberSelections = (detectedBrand?: string) => {
        try {
            const key = (detectedBrand || manualBrand || "last").trim().toLowerCase();
            const current = JSON.parse(window.localStorage.getItem("po-cutting-form-preferences") || "{}");
            const preference = {
                ordersTemplate: manualTemplate,
                linesTemplate: manualLinesTemplate,
                comments: manualComments,
            };
            current.last = preference;
            current[key] = preference;
            window.localStorage.setItem("po-cutting-form-preferences", JSON.stringify(current));
        } catch {
            // Local storage can be unavailable in private browsing.
        }
    };

    const loadBrandSelections = (brand: string) => {
        try {
            const saved = JSON.parse(window.localStorage.getItem("po-cutting-form-preferences") || "{}");
            const preference = saved[brand.trim().toLowerCase()];
            if (!preference) return;
            setManualTemplate(preference.ordersTemplate || "");
            setManualLinesTemplate(preference.linesTemplate || "");
            setManualComments(preference.comments || "");
        } catch {
            // Keep the current values when preferences cannot be read.
        }
    };

    const previewHeaderMappings = async (files: FileList) => {
        setIsPreviewingHeaders(true);
        setHeaderPreviewError("");
        try {
            const formData = new FormData();
            Array.from(files).forEach((file) => formData.append("file", file));
            const response = await fetch("/api/processing/header-preview", {
                method: "POST",
                body: formData,
            });
            const result = await response.json();
            if (!response.ok || result.error) throw new Error(result.error || "Could not preview headers");
            setHeaderPreviews(result.previews || []);
        } catch (error) {
            setHeaderPreviews([]);
            setHeaderPreviewError(error instanceof Error ? error.message : "Could not preview headers");
        } finally {
            setIsPreviewingHeaders(false);
        }
    };

    const handleFilesSelected = (files: FileList | null) => {
        if (!files?.length) return;
        setBuyFiles(files);
        setHeaderPreviews([]);
        void previewHeaderMappings(files);
        const filename = files[0].name.toLowerCase();
        const filenameBrand = BRAND_OPTIONS.find((brand) =>
            filename.includes(brand.toLowerCase().replace(/\s+/g, ""))
            || filename.includes(brand.toLowerCase())
        );
        if (filenameBrand) {
            setManualBrand(filenameBrand);
            loadBrandSelections(filenameBrand);
        }
    };

    const updateHeaderMapping = (previewIndex: number, field: string, header: string) => {
        setHeaderPreviews((current) => current.map((preview, index) => {
            if (index !== previewIndex) return preview;
            const mapping = { ...(preview.mapping || {}) };
            if (header) mapping[field] = header;
            else delete mapping[field];
            return {
                ...preview,
                mapping,
                unmappedColumns: (preview.headers || []).filter(
                    (candidate: string) => !Object.values(mapping).includes(candidate)
                ),
                source: "user confirmed",
                confidence: 100,
            };
        }));
    };

    const incrementPoNumber = (po: string): string => {
        if (!po.trim()) return "PO000001";
        // Find the last numeric sequence in the PO number and increment it
        // e.g., VUOUS0925B -> VUOUS0926B, PO002739-Marketing -> PO002740-Marketing
        const match = po.match(/^(.*)(\d+)([^\d]*)$/);
        if (!match) return po;
        const prefix = match[1];
        const num = match[2];
        const suffix = match[3];
        const next = String(Number(num) + 1).padStart(num.length, "0");
        return `${prefix}${next}${suffix}`;
    };

    const parseColorName = (value: string): { code: string | null; name: string | null } => {
        if (!value) return { code: null, name: null };
        const raw = value.trim();
        // Generic brand format: "BRAND CODE Color Name" or "BRAND-CODE-Color Name"
        // Examples: TNF E6Q TNF Black YOTG Patch, TNF-JK3-TNF Black, VANS 0B8 Mineral Magenta
        const match = raw.match(/^([A-Z][A-Z0-9]*)[-\s]([A-Z0-9]{2,4})[-\s](.+)$/i);
        if (match) {
            return { code: match[2], name: match[3] };
        }
        return { code: null, name: null };
    };

    const extractColorCodeFromSKU = (style: string, sku: string): string | null => {
        if (!style || !sku) return null;
        const styleUpper = style.toUpperCase().trim();
        const skuUpper = sku.toUpperCase().trim();
        if (skuUpper.startsWith(styleUpper)) {
            return skuUpper.slice(styleUpper.length);
        }
        return null;
    };

    const applyColorNamesFromResponse = (data: any, colorNames: Record<string, string | null>): any => {
        if (!data?.output?.length || !colorNames) return data;
        return {
            ...data,
            output: data.output.map((po: any) => ({
                ...po,
                lines: (po.lines || []).map((line: any) => {
                    const sku = (line.productExternalRef || line.sku || line.rawColour || line.colour || line.color || '').trim();
                    const key = Object.keys(colorNames).find((k) => k.toLowerCase() === sku.toLowerCase());
                    const colorName = key ? colorNames[key] : null;
                    if (!colorName) return line;
                    return {
                        ...line,
                        colourName: colorName,
                        colorName: colorName,
                    };
                }),
            })),
        };
    };

    const fillColorNamesFromNextGen = async (data: any): Promise<any> => {
        if (!data?.output?.length) return data;

        return {
            ...data,
            output: data.output.map((po: any, poIdx: number) => ({
                ...po,
                lines: (po.lines || []).map((line: any, lineIdx: number) => {
                    const style = (line.styleNumber || line.style || '').trim();
                    const sku = line.colour || line.color || line.styleColor || line.colourName || line.colourDisplay || '';
                    console.log(`[workflow] color parse [po ${poIdx} line ${lineIdx}] style:`, style, 'sku:', sku, 'line keys:', Object.keys(line));

                    // 1. Try to parse a color name string (e.g., "TNF E6Q TNF Black YOTG Patch")
                    const nameCandidates = [line.colourName, line.colourDisplay, line.colour, line.color, line.styleColor].filter(Boolean) as string[];
                    let bestCode: string | null = null;
                    let bestName: string | null = null;
                    let bestSource: string | null = null;
                    for (const source of nameCandidates) {
                        const { code, name } = parseColorName(source);
                        if (code) {
                            bestCode = code;
                            bestName = name;
                            bestSource = source;
                            break;
                        }
                    }

                    // 2. Fallback: extract color code from SKU by stripping style prefix
                    if (!bestCode && style && sku) {
                        const extractedCode = extractColorCodeFromSKU(style, sku);
                        if (extractedCode) {
                            bestCode = extractedCode;
                            bestSource = sku;
                        }
                    }

                    console.log(`[workflow] color parse result [po ${poIdx} line ${lineIdx}] bestCode:`, bestCode, 'bestName:', bestName, 'bestSource:', bestSource);
                    if (!bestCode || !bestSource) return line;
                    return {
                        ...line,
                        rawColour: bestSource,
                        colour: bestCode,
                        color: bestCode,
                        colourName: bestName || line.colourName,
                        colourDisplay: bestName || line.colourDisplay,
                    };
                }),
            })),
        };
    };

    const buildUploadDataFromGeminiBuyItems = (items: any[]): any => {
        if (!items?.length) return null;

        const get = (item: any, oldKey: string, newKey: string) => {
            const val = item[oldKey] !== undefined ? item[oldKey] : item[newKey];
            return val === null || val === undefined ? '' : String(val).trim();
        };

        const poGroups = new Map<string, any[]>();
        for (const item of items) {
            const po = get(item, 'po_number', 'poNumber') || 'UNKNOWN';
            if (!poGroups.has(po)) poGroups.set(po, []);
            poGroups.get(po)!.push(item);
        }

        const output = Array.from(poGroups.entries()).map(([poNumber, lines]) => ({
            id: `gemini-buy-${poNumber}`,
            poNumber,
            header: {
                purchaseOrder: poNumber,
                customer: lines[0]?.customer || '',
                currency: lines[0]?.currency || 'USD',
            },
            lines: lines.map((item) => ({
                styleNumber: get(item, 'buyer_style_number', 'style'),
                productExternalRef: get(item, 'sku', 'sku'),
                productCustomerRef: get(item, 'buyer_style_number', 'style'),
                colour: get(item, 'color', 'color') || get(item, 'color_code', 'colorCode'),
                color: get(item, 'color', 'color') || get(item, 'color_code', 'colorCode'),
                colourName: get(item, 'color', 'color') || null,
                styleColor: get(item, 'color_code', 'colorCode') || null,
                rawColour: get(item, 'color', 'color') || get(item, 'color_code', 'colorCode'),
                season: get(item, 'season', 'season'),
                exFtyDate: get(item, 'delivery_date', 'deliveryDate'),
                factory: get(item, 'factory', 'factory'),
                customer: get(item, 'customer', 'customer'),
                cost: item.unit_cost != null ? String(item.unit_cost) : item.unitCost != null ? String(item.unitCost) : '',
                sourceSheet: get(item, 'source_sheet', 'sourceSheet'),
                sourceRow: get(item, 'source_row', 'sourceRow'),
            })),
            sizes: lines.map((item) => [
                {
                    productSize: String(item.size || 'One Size'),
                    sizeName: String(item.size || 'One Size'),
                    quantity: Number(item.quantity) || 0,
                },
            ]),
        }));

        return {
            mergedSummary: { orders: output.length, lines: items.length, sizes: items.length },
            output,
            errors: [],
            formatDetection: {
                'gemini-buy': {
                    detectedCustomer: items[0]?.customer || 'Unknown',
                    detectedFormat: 'Qwen Excel Extractor',
                    unmappedColumns: [],
                },
            },
            fileSummary: [
                {
                    filename: 'gemini-buy',
                    orderCount: output.length,
                    lineCount: items.length,
                    sizeCount: items.length,
                },
            ],
        };
    };

    const fetchLatestPOFromNextGen = async (): Promise<string | null> => {
        try {
            const res = await fetch("/api/nextgen-latest-po");
            const result = await res.json();
            if (!res.ok || result.error || !result.poNumber) return null;
            return incrementPoNumber(result.poNumber);
        } catch (err) {
            console.error("Failed to fetch latest PO from NextGen:", err);
            return null;
        }
    };

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

    useEffect(() => {
        const today = new Date().toISOString().split("T")[0];
        setManualKeyDate(today);
    }, []);

    useEffect(() => {
        if (manualSeason.trim()) return;
        const firstFileName = buyFiles?.[0]?.name || "";
        const inferredSeason = inferSeasonFromFilename(firstFileName);
        if (inferredSeason) setManualSeason(inferredSeason);
    }, [buyFiles, manualSeason]);

    const steps: { key: Step; label: string; icon: any }[] = [
        { key: "UPLOAD", label: "Add Buy File", icon: CloudUpload },
        { key: "RUN", label: "Reading Buy File", icon: Loader2 },
        { key: "VALIDATE", label: "Check Products", icon: ShieldCheck },
        { key: "REVIEW", label: "Review Results", icon: FileCheck },
        { key: "DOWNLOAD", label: "Download Files", icon: Download },
    ];

    const currentStepIndex = steps.findIndex(s => s.key === currentStep);
    const fieldBase =
        "w-full rounded-2xl bg-white/5 border border-white/10 px-4 py-4 text-base md:text-lg text-white placeholder:text-slate-500 focus:outline-none focus:ring-2 focus:ring-blue-500/40 min-h-[56px]";
    const selectBase =
        "w-full appearance-none rounded-2xl bg-white/5 border border-white/10 px-4 py-4 text-base md:text-lg text-white focus:outline-none focus:ring-2 focus:ring-blue-500/40 min-h-[56px]";
    const optionalFieldCount = [
        manualKeyDate,
        manualKeyUser1,
        manualKeyUser2,
        manualKeyUser3,
        manualKeyUser4,
        manualKeyUser5,
        manualSeason,
        manualCustomer,
        manualBrand,
        manualDestination,
    ].filter((value) => value.trim()).length;
    const requiredFieldsComplete = Boolean(
        (buyFiles?.length || ocrResults?.length)
        && manualPo.trim()
        && manualTemplate.trim()
        && manualLinesTemplate.trim()
        && manualComments.trim()
        && (manualComments !== "[Other]" || customComment.trim())
        && !isPreviewingHeaders
    );

    const handleProcessOcr = async () => {
        if (!ocrFile) return;

        // Route Excel files through the local Qwen-assisted Buy File Extractor
        if (ocrFile.name.toLowerCase().endsWith('.xlsx')) {
            setIsProcessingOcr(true);
            try {
                const formData = new FormData();
                formData.append("file", ocrFile);
                const res = await fetch("/api/extract-buy-file", {
                    method: "POST",
                    body: formData,
                });
                const result = await res.json();
                if (!res.ok || result.error) {
                    setErrors([{
                        field: "Qwen Excel Extractor",
                        row: 0,
                        message: result.error || "Buy file extraction failed",
                        severity: "CRITICAL"
                    }]);
                    setOcrResults(null);
                    return;
                }

                const extractedItems = result.result?.items || [];
                const productData = result.result?.productData || [];
                const files = result.result?.files || {};
                const unmappedColumns = result.result?.unmappedColumns || [];
                const templateUsed = result.result?.templateUsed || false;
                const headers = result.result?.headers || [];
                const mapping = result.result?.mapping || {};
                const matchIssues = result.result?.matchIssues || [];

                // Save template client-side so future uploads can skip AI
                if (headers.length && Object.keys(mapping).length && !templateUsed) {
                    saveTemplate(headers, mapping, extractedItems[0]?.customer || null);
                }

                if (!files?.orders || !files?.lines || !files?.sizes) {
                    console.error("[workflow] AI extract files missing from response", files);
                    setErrors([{
                        field: "AI Extract",
                        row: 0,
                        message: "Extracted data but failed to generate downloadable Excel files. Please try again.",
                        severity: "CRITICAL"
                    }]);
                    setIsProcessingOcr(false);
                    return;
                }

                const uploadData = buildUploadDataFromGeminiBuyItems(extractedItems);
                const colorNames = result.result?.colorNames || {};
                const dataWithColorNames = applyColorNamesFromResponse(uploadData, colorNames);
                const dataWithFiles = {
                    ...dataWithColorNames,
                    files,
                    formatDetection: {
                        'ai-buy': {
                            detectedCustomer: extractedItems[0]?.customer || 'Unknown',
                            detectedFormat: templateUsed ? 'Template (learned)' : 'AI header mapping',
                            unmappedColumns,
                        },
                    },
                };
                setUploadData(dataWithFiles);

                if (result.result?.warning || matchIssues.length) {
                    const extractionWarnings = result.result?.warning ? [{
                        field: "AI Extract",
                        row: 0,
                        message: result.result.warning,
                        severity: "WARNING"
                    }] : [];
                    setErrors([...matchIssues, ...extractionWarnings]);
                } else {
                    setErrors([]);
                }

                setOcrResults(extractedItems.map((item: any) => ({
                    poNumber: item.po_number || item.poNumber || '',
                    style: item.buyer_style_number || item.style || '',
                    color: item.color || '',
                    size: item.size || '',
                    quantity: Number(item.quantity) || 0,
                    factory: item.factory || '',
                    customer: item.customer || '',
                    season: item.season || '',
                    exFtyDate: item.delivery_date || item.deliveryDate || '',
                    transportMethod: '',
                    plant: '',
                })));
                setOcrNextgenUsed(false);

                const filePo = extractedItems.length > 0 ? (extractedItems[0].po_number || extractedItems[0].poNumber || '') : '';
                const nextPoFromNextGen = result.result?.latestPO?.poNumber || null;
                let newPoNumber = '';
                if (nextPoFromNextGen) {
                    newPoNumber = nextPoFromNextGen;
                    setManualPo(newPoNumber);
                    setExtractedPo(filePo || `NextGen latest → ${newPoNumber}`);
                } else if (filePo) {
                    newPoNumber = incrementPoNumber(filePo);
                    setExtractedPo(filePo);
                    setManualPo(newPoNumber);
                }

                setTimeout(() => {
                    setIsProcessingOcr(false);
                    setCurrentStep("VALIDATE");
                    handleValidateNextgen(dataWithFiles, newPoNumber || manualPoRef.current);
                }, 1500);
                return;
            } catch (err) {
                console.error("Qwen buy extraction failed:", err);
                setErrors([{
                    field: "Qwen Excel Extractor",
                    row: 0,
                    message: "Buy file extraction request failed",
                    severity: "CRITICAL"
                }]);
                setIsProcessingOcr(false);
                return;
            }
        }

        setIsProcessingOcr(true);
        try {
            const formData = new FormData();
            formData.append("file", ocrFile);
            formData.append("fillFromNextgen", "true");
            const res = await fetch("/api/ocr-gemini", {
                method: "POST",
                body: formData,
            });
            const result = await res.json();
            if (!res.ok || result.error) {
                setErrors([{
                    field: "Gemini OCR",
                    row: 0,
                    message: result.error || "OCR failed",
                    severity: "CRITICAL"
                }]);
                setOcrResults(null);
            } else {
                const merged = result.mergedResults || result.ocrResults || [];
                setOcrResults(merged);
                setOcrNextgenUsed(!!result.nextgenUsed);

                const filePo = merged.length > 0 && merged[0].poNumber ? merged[0].poNumber : '';
                const nextPoFromNextGen = await fetchLatestPOFromNextGen();
                let newPoNumber = '';
                if (nextPoFromNextGen) {
                    newPoNumber = nextPoFromNextGen;
                    setManualPo(newPoNumber);
                    setExtractedPo(filePo || `NextGen latest → ${newPoNumber}`);
                } else if (filePo) {
                    newPoNumber = incrementPoNumber(filePo);
                    setExtractedPo(filePo);
                    setManualPo(newPoNumber);
                }
                setErrors([]);
            }
        } catch (err) {
            console.error("OCR failed:", err);
            setErrors([{
                field: "Gemini OCR",
                row: 0,
                message: "OCR request failed",
                severity: "CRITICAL"
            }]);
        } finally {
            setIsProcessingOcr(false);
        }
    };

    const waitForProcessingJob = async (jobId: string) => {
        const deadline = Date.now() + 15 * 60 * 1000;
        while (Date.now() < deadline) {
            await new Promise((resolve) => window.setTimeout(resolve, 1800));
            const response = await fetch(`/api/processing/jobs?id=${encodeURIComponent(jobId)}`, {
                cache: "no-store",
            });
            const job = await response.json();
            if (!response.ok) throw new Error(job.error || "Could not check processing status");
            if (job.status === "completed") return job.result;
            if (job.status === "failed") throw new Error(job.error || job.result?.error || "PO processing failed");
        }
        throw new Error("Processing exceeded 15 minutes. Please try again.");
    };

    const handleStartUpload = async (files: FileList | null = buyFiles) => {
        // If no Excel files but OCR results exist, build uploadData from OCR and proceed
        if ((!files || files.length === 0) && ocrResults && ocrResults.length > 0) {
            setIsProcessing(true);
            setCurrentStep("RUN");

            const orders = ocrResults.map((line, idx) => ({
                id: `ocr-${idx}`,
                poNumber: manualPo,
                style: line.style,
                color: line.color,
                size: line.size,
                quantity: line.quantity,
                factory: line.factory,
                plant: line.plant,
                customer: line.customer,
                season: line.season,
                exFtyDate: line.exFtyDate,
                transportMethod: line.transportMethod,
            }));

            const result = {
                mergedSummary: { orders: orders.length, lines: orders.length, sizes: orders.length },
                output: orders.map((line, idx) => ({
                    id: `ocr-${idx}`,
                    poNumber: manualPo,
                    lines: [{
                        style: line.style,
                        colour: line.color,
                        color: line.color,
                        factory: line.factory,
                        plant: line.plant,
                        customer: line.customer,
                        season: line.season,
                        exFtyDate: line.exFtyDate,
                        transportMethod: line.transportMethod,
                    }],
                    sizes: [{
                        productSize: line.size,
                        sizeName: line.size,
                        quantity: line.quantity,
                    }],
                })),
                errors: [],
                formatDetection: { 'ocr-gemini': { detectedCustomer: 'OCR', detectedFormat: 'Gemini OCR', unmappedColumns: [] } },
                fileSummary: [{ filename: 'ocr-gemini', orderCount: orders.length, lineCount: orders.length, sizeCount: orders.length }],
            };

            const dataWithColorNames = await fillColorNamesFromNextGen(result);
            setUploadData(dataWithColorNames);
            setErrors([]);

            setTimeout(() => {
                setIsProcessing(false);
                setCurrentStep("VALIDATE");
                handleValidateNextgen(dataWithColorNames, manualPoRef.current);
            }, 1500);
            return;
        }

        if (!files || files.length === 0) return;

        setIsProcessing(true);
        setCurrentStep("RUN");

        const formData = new FormData();
        for (let i = 0; i < files.length; i++) {
            formData.append("file", files[i]);
        }
        if (manualPo.trim()) formData.append("manualPo", manualPo.trim());
        if (manualTemplate.trim()) formData.append("manualTemplate", manualTemplate.trim());
        if (manualLinesTemplate.trim()) formData.append("manualLinesTemplate", manualLinesTemplate.trim());
        if (manualComments.trim()) {
            if (manualComments === '[Other]') {
                formData.append("manualComments", customComment.trim());
            } else {
                formData.append("manualComments", manualComments.trim());
            }
        }
        if (manualKeyDate.trim()) formData.append("manualKeyDate", manualKeyDate.trim());
        if (manualKeyUser1.trim()) formData.append("manualKeyUser1", manualKeyUser1.trim());
        if (manualKeyUser2.trim()) formData.append("manualKeyUser2", manualKeyUser2.trim());
        if (manualKeyUser3.trim()) formData.append("manualKeyUser3", manualKeyUser3.trim());
        if (manualKeyUser4.trim()) formData.append("manualKeyUser4", manualKeyUser4.trim());
        if (manualKeyUser5.trim()) formData.append("manualKeyUser5", manualKeyUser5.trim());
        if (manualSeason.trim()) formData.append("manualSeason", manualSeason.trim());
        if (manualCustomer.trim()) formData.append("manualCustomer", manualCustomer.trim());
        if (manualBrand.trim()) formData.append("manualBrand", manualBrand.trim());
        if (manualDestination.trim()) formData.append("manualDestination", manualDestination.trim());
        if (headerPreviews.length) {
            const confirmedMappings = Object.fromEntries(
                headerPreviews
                    .filter((preview) => !preview.error)
                    .map((preview) => [preview.filename, {
                        headers: preview.headers,
                        mapping: preview.mapping,
                    }])
            );
            formData.append("headerMappings", JSON.stringify(confirmedMappings));
        }
        if (Object.keys(nextgenOverrides).length) {
            formData.append("nextgenOverrides", JSON.stringify(nextgenOverrides));
        }

        try {
            const startResponse = await fetch("/api/processing/jobs", {
                method: "POST",
                body: formData,
            });
            const started = await startResponse.json();
            if (!startResponse.ok || !started.jobId) {
                throw new Error(started.error || "Could not start PO processing");
            }
            const result = await waitForProcessingJob(started.jobId);

            if (result.error) {
                console.error("Upload failed:", result.error);
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

            const dataWithColorNames = await fillColorNamesFromNextGen(result);
            const variantSummary = result.nexgenVariantSummary || { requested: 0, resolved: 0 };
            const detectedBrand = result?.fileSummary?.flatMap((file: any) => file.brands || [])?.[0];
            if (detectedBrand) setManualBrand(detectedBrand);
            rememberSelections(detectedBrand);
            setUploadData(dataWithColorNames);
            setErrors(result.errors || []);
            setNextgenValidation({
                exists: true,
                mode: "direct-product-validation",
                matched: Array.from({ length: variantSummary.resolved }, (_, index) => ({ index })),
                missing: Array.from({ length: Math.max(0, variantSummary.requested - variantSummary.resolved) }, (_, index) => ({ index })),
                extra: [],
                matchSummary: variantSummary,
            });

            const extractedPo =
                result?.output?.[0]?.header?.purchaseOrder ||
                result?.output?.[0]?.header?.poNumber ||
                result?.output?.[0]?.purchaseOrder ||
                result?.output?.[0]?.poNumber ||
                result?.mergedSummary?.orders?.[0]?.poNumber ||
                result?.mergedSummary?.orders?.[0]?.purchaseOrder ||
                '';

            // Always get latest PO from NextGen and increment for the new PO number
            const nextPoFromNextGen = await fetchLatestPOFromNextGen();
            let newPoNumber = '';
            if (nextPoFromNextGen) {
                newPoNumber = nextPoFromNextGen;
                setExtractedPo(extractedPo || `NextGen latest → ${newPoNumber}`);
            } else if (extractedPo) {
                // Fallback: increment from file PO if NextGen fails
                newPoNumber = incrementPoNumber(extractedPo);
                setExtractedPo(extractedPo);
            } else {
                setErrors([{
                    field: "Auto PO",
                    row: 0,
                    message: "Could not get latest PO from NextGen or file.",
                    severity: "CRITICAL"
                }]);
            }

            setTimeout(() => {
                setIsProcessing(false);
                setCurrentStep("VALIDATE");
            }, 450);

        } catch (err) {
            console.error(err);
            setErrors([{
                field: "Processing",
                row: 0,
                message: err instanceof Error ? err.message : "PO processing failed",
                severity: "CRITICAL",
            }]);
            setIsProcessing(false);
            setCurrentStep("VALIDATE");
        }
    };

    const handleValidateNextgen = async (data: any = uploadData, explicitPoNumber?: string) => {
        if (!data?.mergedSummary?.orders) return;
        setIsValidatingNextgen(true);
        try {
            console.log('[workflow] validate debug first PO lines:', data?.output?.[0]?.lines?.slice(0, 3));
            const lines = data?.output?.flatMap((po: any) => {
                const sizeEntries = Array.isArray(po.sizes)
                    ? po.sizes.map((sz: any, idx: number) => ({ lineIdx: idx, size: sz }))
                    : Object.entries(po.sizes || {}).map(([lineIdx, sizes]: [string, any]) =>
                        (sizes as any[]).map((sz: any) => ({ lineIdx: Number(lineIdx), size: sz }))
                    ).flat();
                return sizeEntries.map(({ lineIdx, size }: { lineIdx: number; size: any }) => ({
                    style: po.lines?.[lineIdx]?.styleNumber || po.lines?.[lineIdx]?.style || po.lines?.[lineIdx]?.product || po.lines?.[lineIdx]?.productExternalRef || po.lines?.[0]?.styleNumber || po.lines?.[0]?.style || po.lines?.[0]?.product || po.lines?.[0]?.productExternalRef || '',
                    color: po.lines?.[lineIdx]?.colour || po.lines?.[lineIdx]?.color || po.lines?.[lineIdx]?.styleColor || po.lines?.[0]?.colour || po.lines?.[0]?.color || po.lines?.[0]?.styleColor || '',
                    size: size.productSize || size.sizeName || '',
                    quantity: Number(size.quantity) || 0,
                }));
            }) || [];
            console.log('[workflow] validate debug lines:', lines.slice(0, 5));
            const res = await fetch("/api/validate-nextgen", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ poNumber: explicitPoNumber || manualPoRef.current, lines }),
            });
            const result = await res.json();
            setNextgenValidation(result);

            // Merge missing fields from NextGen matched lines into upload data
            if (result.exists && result.matched?.length > 0 && data?.output) {
                const mergedOutput = data.output.map((po: any) => ({
                    ...po,
                    lines: (po.lines || []).map((line: any) => {
                        const ngMatch = result.matched.find((ng: any) =>
                            (line.style || line.product || '').toLowerCase().trim() === (ng.style || '').toLowerCase().trim() &&
                            (line.colour || line.color || '').toLowerCase().trim() === (ng.color || '').toLowerCase().trim()
                        );
                        if (!ngMatch) return line;
                        return {
                            ...line,
                            factory: line.factory || ngMatch.factory || '',
                            plant: line.plant || ngMatch.plant || '',
                            customer: line.customer || ngMatch.customer || '',
                            season: line.season || ngMatch.season || '',
                            exFtyDate: line.exFtyDate || ngMatch.exFtyDate || '',
                            transportMethod: line.transportMethod || ngMatch.transportMethod || '',
                        };
                    }),
                }));
                setUploadData({ ...data, output: mergedOutput, files: data?.files });
            } else {
                setUploadData({ ...data, files: data?.files });
            }
        } catch (err) {
            console.error("NextGen validation failed:", err);
            setNextgenValidation({ error: "Failed to validate with NextGen" });
        } finally {
            setIsValidatingNextgen(false);
        }
    };

    const handleDownload = async (fileType: "orders" | "lines" | "sizes") => {
        try {
            console.log(`[workflow] Initializing download for ${fileType}...`);
            console.log("[workflow] uploadData keys:", uploadData ? Object.keys(uploadData) : null);
            console.log("[workflow] uploadData.files keys:", uploadData?.files ? Object.keys(uploadData.files) : null);
            if (!uploadData?.files?.[fileType]) {
                console.error("[workflow] No file data found in uploadData", fileType, uploadData?.files);
                alert("Walang generated file payload. Paki-run ulit at i-check ang validation result.");
                return;
            }

            const base64 = uploadData.files[fileType];
            const blob = base64ToXlsxBlob(base64);

            const url = window.URL.createObjectURL(blob);
            const link = document.createElement("a");
            link.href = url;
            const generatedAt = new Date().toISOString().replace(/[:.]/g, "-");
            link.download = `${fileType.toUpperCase()}_NEXGEN_VALIDATED_${generatedAt}.xlsx`;
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

    const handleDownloadAll = () => {
        (["orders", "lines", "sizes"] as const).forEach((fileType) => {
            void handleDownload(fileType);
        });
    };

    const confirmExcludedLine = (item: any, index: number) => {
        const key = `${item.purchaseOrder}-${item.lineItem}-${index}`;
        setAcknowledgedLines((current) => ({ ...current, [key]: true }));
        setErrors((current) => {
            const remaining = current.filter((error) => !(
                error.severity === "CRITICAL"
                && String(error.message || "").includes(`PO ${item.purchaseOrder} line ${item.lineItem}`)
            ));
            if (!remaining.some((error) => error.severity === "CRITICAL")) {
                setUploadData((data: any) => data ? { ...data, canProceed: true } : data);
            }
            return remaining;
        });
    };

    const startAnotherFile = () => {
        setBuyFiles(null);
        setUploadData(null);
        setErrors([]);
        setNextgenValidation(null);
        setAcknowledgedLines({});
        setHeaderPreviews([]);
        setHeaderPreviewError("");
        setNextgenOverrides({});
        setExtractedPo("");
        if (fileInputRef.current) fileInputRef.current.value = "";
        setCurrentStep("UPLOAD");
    };

    const base64ToXlsxBlob = (base64: string) => {
        const binary = window.atob(base64);
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) {
            bytes[i] = binary.charCodeAt(i);
        }
        return new Blob(
            [bytes],
            { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
        );
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
    const detectedBrands = Array.from(new Set(
        (uploadData?.fileSummary || []).flatMap((file: any) => file.brands || [])
    )) as string[];
    const variantSummary = uploadData?.nexgenVariantSummary || nextgenValidation?.matchSummary || {};
    const transportMethods = Array.from(new Set(
        (uploadData?.output || []).flatMap((po: any) =>
            (po.lines || []).map((line: any) => line.transportMethod).filter(Boolean)
        )
    )) as string[];
    const attentionItems = (uploadData?.needsAttention || [])
        .map((item: any, index: number) => ({ ...item, _attentionIndex: index }))
        .filter((item: any) =>
            !acknowledgedLines[`${item.purchaseOrder}-${item.lineItem}-${item._attentionIndex}`]
        );

    return (
        <div className="w-full max-w-7xl mx-auto space-y-10 px-4 transition-colors duration-300 text-[hsl(var(--foreground))]">
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
                                        "w-11 h-11 md:w-16 md:h-16 rounded-xl md:rounded-2xl flex items-center justify-center transition-all duration-700 ease-out",
                                        isActive ? "bg-blue-600 shadow-[0_0_40px_rgba(37,99,235,0.4)] scale-110 ring-4 ring-blue-500/20" :
                                            isCompleted ? "bg-emerald-600 shadow-[0_0_30px_rgba(16,185,129,0.2)] scale-100" : "bg-slate-900 border border-white/5 scale-90"
                                    )}
                                >
                                    {isActive && isProcessing ? (
                                        <Loader2 className="w-7 h-7 animate-[spin_2s_linear_infinite] text-white" />
                                    ) : (
                                        <Icon className={cn("w-5 h-5 md:w-7 md:h-7 transition-colors duration-500", (isActive || isCompleted) ? "text-white" : "text-slate-600 translate-y-0")} />
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
                                    "step-label mt-3 md:mt-6 max-w-[64px] md:max-w-none text-center text-[8px] md:text-[10px] font-black uppercase tracking-[0.12em] md:tracking-[0.3em] transition-all duration-700",
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
                className="glass-panel rounded-[32px] relative overflow-hidden p-6 md:p-12 min-h-[600px] flex items-center justify-center border-white/[0.08]"
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
                            className="text-center max-w-4xl mx-auto space-y-8"
                        >
                            <div className="relative inline-block group">
                                <div className="absolute inset-0 bg-blue-500/20 blur-[80px] rounded-full group-hover:bg-blue-500/30 transition-all duration-500" />
                                <div className="relative w-24 h-24 bg-[hsl(var(--panel))] border border-[hsl(var(--border))] rounded-3xl flex items-center justify-center mx-auto transition-all duration-500 group-hover:scale-105 group-hover:border-blue-500/50 shadow-2xl">
                                    <CloudUpload className="w-10 h-10 text-cyan-400 group-hover:-translate-y-1 transition-all" />
                                </div>
                            </div>

                            <div className="space-y-6">
                                <h2 className="text-4xl md:text-5xl font-black tracking-tight text-[hsl(var(--foreground))] leading-tight">
                                    Create PO Cutting Files
                                </h2>
                                <p className="text-[hsl(var(--muted))] text-base md:text-lg font-medium max-w-xl mx-auto">
                                    Upload the buyer Excel, confirm the required settings, and the system will check NextGen and prepare all three files.
                                </p>
                            </div>

                            <div className="grid grid-cols-1 gap-4 max-w-3xl mx-auto text-left">
                                <div className="space-y-2">
                                    <label className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-300">Buy File <span className="text-rose-400">Required</span></label>
                                    <input
                                        ref={fileInputRef}
                                        type="file"
                                        accept=".xlsx"
                                        multiple
                                        onChange={(e) => handleFilesSelected(e.target.files)}
                                        className="sr-only"
                                    />
                                    <div
                                        role="button"
                                        tabIndex={0}
                                        onClick={() => fileInputRef.current?.click()}
                                        onKeyDown={(event) => {
                                            if (event.key === "Enter" || event.key === " ") fileInputRef.current?.click();
                                        }}
                                        onDragEnter={(event) => { event.preventDefault(); setIsDragging(true); }}
                                        onDragOver={(event) => { event.preventDefault(); setIsDragging(true); }}
                                        onDragLeave={(event) => { event.preventDefault(); setIsDragging(false); }}
                                        onDrop={(event) => {
                                            event.preventDefault();
                                            setIsDragging(false);
                                            handleFilesSelected(event.dataTransfer.files);
                                        }}
                                        className={cn(
                                            "min-h-[150px] rounded-3xl border-2 border-dashed p-5 flex items-center justify-center cursor-pointer transition-all",
                                            isDragging ? "border-blue-400 bg-blue-500/15 scale-[1.01]" : "border-white/15 bg-white/[0.03] hover:border-blue-500/50 hover:bg-blue-500/[0.06]"
                                        )}
                                    >
                                        {buyFiles?.length ? (
                                            <div className="w-full flex items-center gap-4 text-left">
                                                <div className="rounded-2xl bg-emerald-500/10 border border-emerald-500/20 p-3">
                                                    <FileCheck className="w-7 h-7 text-emerald-400" />
                                                </div>
                                                <div className="min-w-0 flex-1">
                                                    <p className="font-black text-sm text-[hsl(var(--foreground))] truncate">{buyFiles[0].name}</p>
                                                    <p className="text-xs text-[hsl(var(--muted))] mt-1">
                                                        {(buyFiles[0].size / 1024).toFixed(1)} KB{buyFiles.length > 1 ? ` · ${buyFiles.length} files selected` : ""}
                                                    </p>
                                                    <p className="text-[10px] text-blue-400 font-bold uppercase tracking-widest mt-2">Click to replace</p>
                                                </div>
                                                <button
                                                    type="button"
                                                    aria-label="Remove selected file"
                                                    onClick={(event) => {
                                                        event.stopPropagation();
                                                        setBuyFiles(null);
                                                        if (fileInputRef.current) fileInputRef.current.value = "";
                                                    }}
                                                    className="rounded-full border border-white/10 p-2 text-slate-400 hover:text-white hover:bg-white/10"
                                                >
                                                    <X className="w-4 h-4" />
                                                </button>
                                            </div>
                                        ) : (
                                            <div className="text-center">
                                                <Upload className="w-8 h-8 text-blue-400 mx-auto mb-3" />
                                                <p className="font-black text-sm text-[hsl(var(--foreground))]">Drop the buyer Excel here</p>
                                                <p className="text-xs text-[hsl(var(--muted))] mt-1">or click to choose an .xlsx file</p>
                                            </div>
                                        )}
                                    </div>
                                    {buyFiles && buyFiles.length > 1 && (
                                        <div className="text-[10px] text-slate-400 break-words">
                                            {Array.from(buyFiles).slice(1).map((file) => file.name).join(", ")}
                                        </div>
                                    )}
                                </div>
                                {(isPreviewingHeaders || headerPreviewError || headerPreviews.length > 0) && (
                                    <div className="rounded-3xl border border-blue-500/20 bg-blue-500/[0.05] p-4 md:p-5 text-left">
                                        {isPreviewingHeaders ? (
                                            <div className="flex items-center gap-3 text-sm text-blue-200">
                                                <Loader2 className="w-5 h-5 animate-spin" />
                                                Detecting and mapping buyer headers…
                                            </div>
                                        ) : headerPreviewError ? (
                                            <div className="flex flex-wrap items-center justify-between gap-3">
                                                <p className="text-xs text-amber-300">{headerPreviewError}</p>
                                                <button
                                                    type="button"
                                                    onClick={() => buyFiles && void previewHeaderMappings(buyFiles)}
                                                    className="rounded-xl border border-amber-500/20 px-3 py-2 text-[10px] font-black uppercase tracking-widest text-amber-200"
                                                >
                                                    Try preview again
                                                </button>
                                            </div>
                                        ) : (
                                            <div className="space-y-4">
                                                <div>
                                                    <h3 className="font-black text-sm text-white">Confirm Header Mapping</h3>
                                                    <p className="text-xs text-slate-400 mt-1">
                                                        Check what each buyer column means. Confirmed corrections are learned automatically for future files.
                                                    </p>
                                                </div>
                                                {headerPreviews.map((preview, previewIndex) => (
                                                    <details key={`${preview.filename}-${previewIndex}`} open className="rounded-2xl border border-white/10 bg-black/15 p-4">
                                                        <summary className="cursor-pointer list-none flex flex-wrap items-center justify-between gap-3">
                                                            <div>
                                                                <p className="text-xs font-black text-white">{preview.filename}</p>
                                                                <p className="text-[10px] text-slate-500 mt-1">
                                                                    Sheet: {preview.worksheet} · Header row {preview.headerRow}
                                                                </p>
                                                            </div>
                                                            <div className="flex gap-2">
                                                                <span className="rounded-full border border-blue-500/20 bg-blue-500/10 px-3 py-1 text-[9px] font-black uppercase tracking-widest text-blue-200">
                                                                    {preview.source}
                                                                </span>
                                                                <span className="rounded-full border border-emerald-500/20 bg-emerald-500/10 px-3 py-1 text-[9px] font-black uppercase tracking-widest text-emerald-200">
                                                                    {preview.confidence}% confidence
                                                                </span>
                                                            </div>
                                                        </summary>
                                                        <div className="mt-5 grid grid-cols-1 md:grid-cols-2 gap-3">
                                                            {HEADER_FIELD_OPTIONS.map(([field, label]) => (
                                                                <label key={`${preview.filename}-${field}`} className="space-y-1.5">
                                                                    <span className="text-[9px] font-black uppercase tracking-widest text-slate-400">{label}</span>
                                                                    <select
                                                                        value={preview.mapping?.[field] || ""}
                                                                        onChange={(event) => updateHeaderMapping(previewIndex, field, event.target.value)}
                                                                        className="w-full rounded-xl border border-white/10 bg-slate-900 px-3 py-2.5 text-xs text-white focus:border-blue-500/50 focus:outline-none"
                                                                    >
                                                                        <option value="">Not mapped</option>
                                                                        {(preview.headers || []).map((header: string, headerIndex: number) => (
                                                                            <option key={`${header}-${headerIndex}`} value={header}>{header}</option>
                                                                        ))}
                                                                    </select>
                                                                </label>
                                                            ))}
                                                        </div>
                                                        {preview.unmappedColumns?.length > 0 && (
                                                            <p className="mt-4 text-[10px] text-slate-500">
                                                                Ignored columns: {preview.unmappedColumns.join(", ")}
                                                            </p>
                                                        )}
                                                    </details>
                                                ))}
                                            </div>
                                        )}
                                    </div>
                                )}
                                <div className="space-y-2">
                                    <div className="flex items-center justify-between">
                                        <label className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-500">PO Number <span className="text-rose-400">*</span></label>
                                        <button
                                            type="button"
                                            onClick={() => setManualPo(incrementPoNumber(manualPo))}
                                            className="text-[10px] font-bold uppercase tracking-widest text-blue-400 hover:text-blue-300 transition-colors"
                                        >
                                            Auto +1
                                        </button>
                                    </div>
                                    <input
                                        value={manualPo}
                                        onChange={(e) => setManualPo(e.target.value.toUpperCase())}
                                        placeholder="PO002954"
                                        className="w-full rounded-xl bg-white/5 border border-white/10 px-4 py-3 text-sm text-white placeholder:text-slate-600 focus:outline-none focus:ring-2 focus:ring-blue-500/40"
                                    />
                                    {extractedPo && (
                                        <div className="flex items-center justify-between text-[10px]">
                                            <span className="text-slate-500">Extracted from file:</span>
                                            <span className="text-slate-400 font-mono">{extractedPo}</span>
                                        </div>
                                    )}
                                </div>
                                <div className="space-y-2">
                                    <label className="text-[10px] font-black uppercase tracking-[0.3em] text-cyan-300">Orders Template <span className="text-rose-400">Required</span></label>
                                    <select
                                        value={manualTemplate}
                                        onChange={e => setManualTemplate(e.target.value)}
                                        className="w-full rounded-xl bg-slate-900 border border-blue-700 px-4 py-4 text-base md:text-lg text-white focus:outline-none focus:ring-2 focus:ring-blue-400/80 min-h-[56px] transition-colors duration-200 shadow-lg hover:border-blue-400"
                                        style={{ backgroundColor: '#181e29', color: '#fff' }}
                                    >
                                        <option value="" disabled className="text-slate-400 bg-slate-800">Choose a template</option>
                                        {TEMPLATE_OPTIONS.map(option => (
                                            <option key={option} value={option} className="text-slate-200 bg-slate-900 hover:bg-blue-900">{option}</option>
                                        ))}
                                    </select>
                                </div>
                                <div className="space-y-2">
                                    <label className="text-[10px] font-black uppercase tracking-[0.3em] text-cyan-300">Lines Template <span className="text-rose-400">Required</span></label>
                                    <select
                                        value={manualLinesTemplate}
                                        onChange={e => setManualLinesTemplate(e.target.value)}
                                        className="w-full rounded-xl bg-slate-900 border border-blue-700 px-4 py-4 text-base md:text-lg text-white focus:outline-none focus:ring-2 focus:ring-blue-400/80 min-h-[56px] transition-colors duration-200 shadow-lg hover:border-blue-400"
                                        style={{ backgroundColor: '#181e29', color: '#fff' }}
                                    >
                                        <option value="" disabled className="text-slate-400 bg-slate-800">Choose lines template</option>
                                        {LINE_TEMPLATE_OPTIONS.map(option => (
                                            <option key={option} value={option} className="text-slate-200 bg-slate-900 hover:bg-blue-900">{option}</option>
                                        ))}
                                    </select>
                                </div>
                                <div className="space-y-2">
                                    <label className="text-[10px] font-black uppercase tracking-[0.3em] text-cyan-300">Order Comments <span className="text-rose-400">Required</span></label>
                                    <select
                                        value={manualComments}
                                        onChange={e => {
                                            setManualComments(e.target.value);
                                            if (e.target.value !== '[Other]') setCustomComment("");
                                        }}
                                        className="w-full rounded-2xl bg-slate-900 border border-white/10 px-4 py-4 text-base md:text-lg text-white focus:outline-none focus:ring-2 focus:ring-blue-400/80 min-h-[56px] transition-colors duration-200 shadow-lg hover:border-blue-400"
                                        style={{ backgroundColor: '#181e29', color: '#fff' }}
                                    >
                                        <option value="" disabled>Choose a comment</option>
                                        {COMMENT_OPTIONS.map(option => (
                                            <option key={option} value={option}>{option}</option>
                                        ))}
                                    </select>
                                    {manualComments === '[Other]' && (
                                        <textarea
                                            value={customComment}
                                            onChange={e => setCustomComment(e.target.value)}
                                            placeholder="Enter your comment here"
                                            rows={2}
                                            className="w-full rounded-2xl bg-slate-900 border border-white/10 px-4 py-4 text-base md:text-lg text-white placeholder:text-slate-400 focus:outline-none focus:ring-2 focus:ring-blue-400/80 min-h-[56px] mt-2 transition-colors duration-200 shadow-lg hover:border-blue-400"
                                            style={{ backgroundColor: '#181e29', color: '#fff' }}
                                        />
                                    )}
                                </div>
                                <div className="flex flex-wrap items-center justify-between gap-3 rounded-2xl border border-white/10 bg-white/5 px-4 py-3">
                                    <div>
                                        <div className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-500">Advanced overrides</div>
                                        <div className="text-xs text-slate-400 mt-1">
                                            Optional fields for files that are missing columns. {optionalFieldCount > 0 ? `${optionalFieldCount} filled.` : "Keep this collapsed unless you need a fallback."}
                                        </div>
                                    </div>
                                    <button
                                        type="button"
                                        onClick={() => setShowAdvanced((v) => !v)}
                                        className="rounded-full border border-blue-500/30 bg-blue-500/10 px-4 py-2 text-xs font-bold uppercase tracking-[0.2em] text-blue-200 hover:bg-blue-500/20 transition-colors"
                                    >
                                        {showAdvanced ? "Hide advanced" : "Show advanced"}
                                    </button>
                                </div>
                                {showAdvanced && (
                                    <div className="space-y-4 rounded-3xl border border-white/10 bg-black/20 p-4 md:p-6">
                                        <div className="space-y-2">
                                            <label className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-500">Order Key Date <span className="normal-case tracking-normal">(auto: today)</span></label>
                                            <input
                                                value={manualKeyDate}
                                                onChange={(e) => setManualKeyDate(e.target.value)}
                                                placeholder="Leave blank to use today's date"
                                                className="w-full rounded-xl bg-white/5 border border-white/10 px-4 py-3 text-sm text-white placeholder:text-slate-600 focus:outline-none focus:ring-2 focus:ring-blue-500/40"
                                            />
                                        </div>
                                        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                                            <div className="space-y-2">
                                                <label className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-500">KeyUser1</label>
                                                <input
                                                    value={manualKeyUser1}
                                                    onChange={(e) => setManualKeyUser1(e.target.value)}
                                                    placeholder="Planning"
                                                    className="w-full rounded-xl bg-white/5 border border-white/10 px-4 py-3 text-sm text-white placeholder:text-slate-600 focus:outline-none focus:ring-2 focus:ring-blue-500/40"
                                                />
                                            </div>
                                            <div className="space-y-2">
                                                <label className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-500">KeyUser2</label>
                                                <input
                                                    value={manualKeyUser2}
                                                    onChange={(e) => setManualKeyUser2(e.target.value)}
                                                    placeholder="Purchasing"
                                                    className="w-full rounded-xl bg-white/5 border border-white/10 px-4 py-3 text-sm text-white placeholder:text-slate-600 focus:outline-none focus:ring-2 focus:ring-blue-500/40"
                                                />
                                            </div>
                                            <div className="space-y-2">
                                                <label className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-500">KeyUser3</label>
                                                <input
                                                    value={manualKeyUser3}
                                                    onChange={(e) => setManualKeyUser3(e.target.value)}
                                                    placeholder="Optional"
                                                    className="w-full rounded-xl bg-white/5 border border-white/10 px-4 py-3 text-sm text-white placeholder:text-slate-600 focus:outline-none focus:ring-2 focus:ring-blue-500/40"
                                                />
                                            </div>
                                            <div className="space-y-2">
                                                <label className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-500">KeyUser4</label>
                                                <input
                                                    value={manualKeyUser4}
                                                    onChange={(e) => setManualKeyUser4(e.target.value)}
                                                    placeholder="Production"
                                                    className="w-full rounded-xl bg-white/5 border border-white/10 px-4 py-3 text-sm text-white placeholder:text-slate-600 focus:outline-none focus:ring-2 focus:ring-blue-500/40"
                                                />
                                            </div>
                                            <div className="space-y-2 md:col-span-2">
                                                <label className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-500">KeyUser5</label>
                                                <input
                                                    value={manualKeyUser5}
                                                    onChange={(e) => setManualKeyUser5(e.target.value)}
                                                    placeholder="Logistics"
                                                    className="w-full rounded-xl bg-white/5 border border-white/10 px-4 py-3 text-sm text-white placeholder:text-slate-600 focus:outline-none focus:ring-2 focus:ring-blue-500/40"
                                                />
                                            </div>
                                        </div>
                                        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                                            <div className="space-y-2">
                                                <label className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-500">Season</label>
                                                <input
                                                    value={manualSeason}
                                                    onChange={(e) => setManualSeason(e.target.value)}
                                                    placeholder="Choose a season"
                                                    className="w-full rounded-xl bg-white/5 border border-white/10 px-4 py-4 text-base md:text-lg text-white placeholder:text-slate-500 focus:outline-none focus:ring-2 focus:ring-blue-500/40"
                                                />
                                            </div>
                                            <div className="space-y-2">
                                                <label className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-500">Customer</label>
                                                <input
                                                    value={manualCustomer}
                                                    onChange={(e) => setManualCustomer(e.target.value)}
                                                    placeholder="511 Tactical"
                                                    list="customer-presets"
                                                    className="w-full rounded-xl bg-white/5 border border-white/10 px-4 py-3 text-sm text-white placeholder:text-slate-600 focus:outline-none focus:ring-2 focus:ring-blue-500/40"
                                                />
                                            </div>
                                            <div className="space-y-2">
                                                <label className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-500">Brand</label>
                                                <input
                                                    value={manualBrand}
                                                    onChange={(e) => setManualBrand(e.target.value)}
                                                    placeholder="haglofs"
                                                    list="brand-presets"
                                                    className="w-full rounded-xl bg-white/5 border border-white/10 px-4 py-3 text-sm text-white placeholder:text-slate-600 focus:outline-none focus:ring-2 focus:ring-blue-500/40"
                                                />
                                            </div>
                                            <div className="space-y-2">
                                                <label className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-500">Destination</label>
                                                <input
                                                    value={manualDestination}
                                                    onChange={(e) => setManualDestination(e.target.value)}
                                                    placeholder="Choose a destination"
                                                    list="destination-presets"
                                                    className="w-full rounded-xl bg-white/5 border border-white/10 px-4 py-4 text-base md:text-lg text-white placeholder:text-slate-500 focus:outline-none focus:ring-2 focus:ring-blue-500/40"
                                                />
                                            </div>
                                        </div>
                                        <div className="flex justify-end">
                                            <button
                                                type="button"
                                                onClick={() => {
                                                    setManualKeyUser1("");
                                                    setManualKeyUser2("");
                                                    setManualKeyUser3("");
                                                    setManualKeyUser4("");
                                                    setManualKeyUser5("");
                                                    setManualSeason("");
                                                    setManualCustomer("");
                                                    setManualBrand("");
                                                    setManualDestination("");
                                                }}
                                                className="text-xs font-bold uppercase tracking-[0.2em] text-slate-400 hover:text-white transition-colors"
                                            >
                                                Clear advanced fields
                                            </button>
                                        </div>
                                    </div>
                                )}
                            </div>

                            <datalist id="template-presets">
                                {TEMPLATE_OPTIONS.map(option => <option key={option} value={option} />)}
                            </datalist>
                            <datalist id="lines-template-presets">
                                {LINE_TEMPLATE_OPTIONS.map(option => <option key={option} value={option} />)}
                            </datalist>
                            <datalist id="brand-presets">
                                {BRAND_OPTIONS.map(option => <option key={option} value={option} />)}
                            </datalist>
                            <datalist id="customer-presets">
                                {CUSTOMER_OPTIONS.map(option => <option key={option} value={option} />)}
                            </datalist>
                            <datalist id="destination-presets">
                                {DESTINATION_OPTIONS.map(option => <option key={option} value={option} />)}
                            </datalist>

                            <div className="sticky bottom-3 z-30 md:static flex flex-col items-center gap-4 rounded-2xl md:rounded-none border border-white/10 md:border-0 bg-[hsl(var(--panel)/0.96)] md:bg-transparent p-3 md:p-0 backdrop-blur-xl md:backdrop-blur-none">
                                <button
                                    onClick={() => handleStartUpload()}
                                    disabled={!requiredFieldsComplete}
                                    className="primary-button w-full md:w-auto inline-flex justify-center items-center gap-4 bg-blue-600 text-white disabled:opacity-50 disabled:cursor-not-allowed"
                                    style={{ background: "linear-gradient(90deg, #2563eb, #1d4ed8)" }}
                                >
                                    <span>Process Buy File</span>
                                    <ArrowRight className="w-4 h-4" />
                                </button>
                                {!requiredFieldsComplete && (
                                    <p className="text-[10px] text-center text-amber-400 font-bold">
                                        Add the file and complete the required fields to continue.
                                    </p>
                                )}
                            </div>
                        </motion.div>
                    )}

                    {currentStep === "RUN" && (
                        <motion.div
                            key="run"
                            initial={{ opacity: 0, scale: 0.95 }}
                            animate={{ opacity: 1, scale: 1 }}
                            className="max-w-2xl mx-auto space-y-8"
                        >
                            <div className="relative w-24 h-24 mx-auto">
                                <div className="absolute inset-0 bg-blue-500/20 blur-3xl animate-pulse" />
                                <Loader2 className="w-full h-full text-blue-500 animate-[spin_4s_linear_infinite]" />
                            </div>

                            <div className="text-center space-y-3">
                                <h3 className="text-3xl font-black text-white">Processing your buy file</h3>
                                <p className="text-sm text-slate-400">Keep this page open. The system is extracting and validating the file.</p>
                            </div>

                            <div className="rounded-3xl border border-white/10 bg-white/[0.03] p-5 md:p-7 space-y-3">
                                {processingStages.map((label, index) => {
                                    const complete = index < processingStage;
                                    const active = index === processingStage;
                                    return (
                                        <div key={label} className={cn(
                                            "flex items-center gap-4 rounded-2xl border px-4 py-3 text-left transition-all",
                                            active ? "border-blue-500/40 bg-blue-500/10" : complete ? "border-emerald-500/20 bg-emerald-500/[0.06]" : "border-white/5 opacity-50"
                                        )}>
                                            {complete ? (
                                                <CheckCircle2 className="w-5 h-5 text-emerald-400 shrink-0" />
                                            ) : active ? (
                                                <Loader2 className="w-5 h-5 text-blue-400 animate-spin shrink-0" />
                                            ) : (
                                                <div className="w-5 h-5 rounded-full border border-white/20 shrink-0" />
                                            )}
                                            <span className={cn("text-sm font-semibold", active ? "text-blue-200" : complete ? "text-emerald-200" : "text-slate-500")}>
                                                {label}
                                            </span>
                                        </div>
                                    );
                                })}
                                <div className="pt-3">
                                    <div className="h-2 bg-white/5 rounded-full overflow-hidden">
                                        <motion.div
                                            animate={{ width: `${((processingStage + 1) / processingStages.length) * 100}%` }}
                                            className="h-full bg-gradient-to-r from-blue-600 to-cyan-400 rounded-full"
                                        />
                                    </div>
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
                                        <h2 className="text-3xl font-black tracking-tight text-white">Validation Results</h2>
                                        <p className="text-sm text-slate-400 mt-1">Products, colours, and required data have been checked.</p>
                                    </div>
                                </div>
                                <div className="flex flex-row items-center gap-4">
                                    <button
                                        onClick={startAnotherFile}
                                        className="inline-flex items-center gap-2 text-slate-500 hover:text-white transition-colors text-[10px] font-black uppercase tracking-widest mr-4 group"
                                    >
                                        <History className="w-4 h-4 group-hover:-rotate-90 transition-transform" /> Upload Another File
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
                                            <span>Continue to Review</span> <ChevronRight className="w-5 h-5 group-hover:translate-x-1 transition-transform" />
                                        </button>
                                        {errors.some(e => e.severity === "CRITICAL") && (
                                            <p className="text-[9px] font-black text-red-500 uppercase tracking-widest mt-2 animate-pulse">
                                                Resolve or exclude the highlighted lines to continue
                                            </p>
                                        )}
                                    </div>
                                </div>
                            </div>

                            {/* NextGen Validation Panel */}
                            <div className="mb-8 rounded-3xl border border-white/10 bg-slate-900/60 p-6">
                                <div className="flex flex-col md:flex-row items-center justify-between gap-4">
                                    <div className="flex items-center gap-3">
                                        <div className="p-3 bg-blue-500/10 rounded-2xl border border-blue-500/20">
                                            <ShieldCheck className="w-5 h-5 text-blue-400" />
                                        </div>
                                        <div>
                                            <p className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-500">NextGen Product &amp; Colour Check</p>
                                            <p className="text-sm font-medium text-slate-300">The buyer references were matched against NextGen.</p>
                                        </div>
                                    </div>
                                    <button
                                        onClick={handleValidateNextgen}
                                        disabled={isValidatingNextgen || !uploadData?.mergedSummary?.orders}
                                        className="inline-flex items-center gap-2 rounded-xl border border-blue-500/30 bg-blue-500/10 px-5 py-3 text-xs font-bold uppercase tracking-widest text-blue-200 hover:bg-blue-500/20 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                                    >
                                        {isValidatingNextgen ? <Loader2 className="w-4 h-4 animate-spin" /> : <ShieldCheck className="w-4 h-4" />}
                                        {isValidatingNextgen ? "Checking..." : "Check Again"}
                                    </button>
                                </div>

                                {nextgenValidation && (
                                    <div className="mt-6 grid grid-cols-1 md:grid-cols-3 gap-4">
                                        <div className="rounded-2xl border border-emerald-500/20 bg-emerald-500/10 p-4">
                                            <p className="text-[10px] font-black uppercase tracking-widest text-emerald-500 mb-1">Matched</p>
                                            <p className="text-2xl font-black text-emerald-300">{nextgenValidation.matched?.length || 0}</p>
                                        </div>
                                        <div className="rounded-2xl border border-amber-500/20 bg-amber-500/10 p-4">
                                            <p className="text-[10px] font-black uppercase tracking-widest text-amber-500 mb-1">Missing in NextGen</p>
                                            <p className="text-2xl font-black text-amber-300">{nextgenValidation.missing?.length || 0}</p>
                                        </div>
                                        <div className="rounded-2xl border border-rose-500/20 bg-rose-500/10 p-4">
                                            <p className="text-[10px] font-black uppercase tracking-widest text-rose-500 mb-1">Extra in Upload</p>
                                            <p className="text-2xl font-black text-rose-300">{nextgenValidation.extra?.length || 0}</p>
                                        </div>
                                    </div>
                                )}

                                {nextgenValidation?.error && (
                                    <p className="mt-4 text-xs font-black text-rose-500 uppercase tracking-widest">{nextgenValidation.error}</p>
                                )}

                                {nextgenValidation && !nextgenValidation.exists && !nextgenValidation.error && (
                                    <p className="mt-4 text-xs font-black text-rose-500 uppercase tracking-widest">PO not found in NextGen</p>
                                )}

                                {nextgenValidation?.exists && nextgenValidation.foundBy === 'styleColor' && (
                                    <p className="mt-4 text-xs font-black text-blue-400 uppercase tracking-widest">PO matched by style + colorway</p>
                                )}
                            </div>

                            {attentionItems.length > 0 && (
                                <div className="rounded-[28px] border border-red-500/25 bg-red-500/[0.06] p-7 space-y-5">
                                    <div className="flex items-start gap-4">
                                        <AlertCircle className="w-6 h-6 text-red-400 mt-0.5" />
                                        <div>
                                            <h3 className="text-sm font-black uppercase tracking-[0.22em] text-red-300">Unresolved Buyer References</h3>
                                            <p className="text-xs text-slate-400 mt-2">
                                                These unresolved lines were excluded from the final Excel files to prevent an incorrect Nexgen import.
                                            </p>
                                        </div>
                                    </div>
                                    <div className="grid gap-3">
                                        {attentionItems.map((item: any, index: number) => (
                                            <div key={`${item.purchaseOrder}-${item.lineItem}-${index}`} className="rounded-2xl border border-white/5 bg-slate-950/50 px-5 py-4">
                                                <div className="flex flex-wrap items-center gap-3">
                                                    <span className="text-[9px] font-black uppercase tracking-widest text-red-400">{item.code}</span>
                                                    <span className="text-[10px] font-mono text-slate-300">PO {item.purchaseOrder}</span>
                                                    <span className="text-[10px] font-mono text-slate-500">LINE {item.lineItem}</span>
                                                </div>
                                                <p className="text-xs text-slate-400 mt-2">{item.message}</p>
                                                <p className="text-[10px] font-mono text-slate-500 mt-2">
                                                    Buyer style: {item.style || "(blank)"} · Buyer colour: {item.colour || "(blank)"}
                                                </p>
                                                {item.candidates?.length > 0 && (
                                                    <label className="mt-4 block space-y-2">
                                                        <span className="text-[9px] font-black uppercase tracking-widest text-blue-300">Choose the correct NextGen match</span>
                                                        <select
                                                            value={nextgenOverrides[`${String(item.style).toLowerCase()}|${String(item.colour).toLowerCase()}`]
                                                                ? JSON.stringify(nextgenOverrides[`${String(item.style).toLowerCase()}|${String(item.colour).toLowerCase()}`])
                                                                : ""}
                                                            onChange={(event) => {
                                                                const key = `${String(item.style).toLowerCase()}|${String(item.colour).toLowerCase()}`;
                                                                setNextgenOverrides((current) => {
                                                                    const updated = { ...current };
                                                                    if (event.target.value) updated[key] = JSON.parse(event.target.value);
                                                                    else delete updated[key];
                                                                    return updated;
                                                                });
                                                            }}
                                                            className="w-full rounded-xl border border-blue-500/25 bg-slate-900 px-3 py-3 text-xs text-white"
                                                        >
                                                            <option value="">Select a NextGen product and colour</option>
                                                            {item.candidates.map((candidate: any, candidateIndex: number) => (
                                                                <option key={`${candidate.product}-${candidate.colorName}-${candidateIndex}`} value={JSON.stringify(candidate)}>
                                                                    {candidate.product} · {candidate.colorName} · {candidate.score}% match
                                                                </option>
                                                            ))}
                                                        </select>
                                                    </label>
                                                )}
                                                <div className="mt-4 flex flex-wrap gap-2">
                                                    <button
                                                        type="button"
                                                        onClick={() => setCurrentStep("UPLOAD")}
                                                        className="inline-flex items-center gap-2 rounded-xl border border-blue-500/25 bg-blue-500/10 px-3 py-2 text-[10px] font-black uppercase tracking-widest text-blue-200 hover:bg-blue-500/20"
                                                    >
                                                        <Upload className="w-3.5 h-3.5" /> Correct source &amp; retry
                                                    </button>
                                                    <button
                                                        type="button"
                                                        onClick={() => navigator.clipboard?.writeText(`PO ${item.purchaseOrder}, line ${item.lineItem}, style ${item.style || "(blank)"}, colour ${item.colour || "(blank)"}: ${item.message}`)}
                                                        className="inline-flex items-center gap-2 rounded-xl border border-white/10 bg-white/5 px-3 py-2 text-[10px] font-black uppercase tracking-widest text-slate-300 hover:bg-white/10"
                                                    >
                                                        <Copy className="w-3.5 h-3.5" /> Copy details
                                                    </button>
                                                    <button
                                                        type="button"
                                                        onClick={() => confirmExcludedLine(item, item._attentionIndex)}
                                                        className="inline-flex items-center gap-2 rounded-xl border border-amber-500/25 bg-amber-500/10 px-3 py-2 text-[10px] font-black uppercase tracking-widest text-amber-200 hover:bg-amber-500/20"
                                                    >
                                                        <Eye className="w-3.5 h-3.5" /> Keep line excluded
                                                    </button>
                                                </div>
                                            </div>
                                        ))}
                                    </div>
                                    {Object.keys(nextgenOverrides).length > 0 && (
                                        <div className="sticky bottom-3 rounded-2xl border border-blue-500/25 bg-slate-950/95 p-3 backdrop-blur-xl">
                                            <button
                                                type="button"
                                                onClick={() => handleStartUpload()}
                                                disabled={isProcessing}
                                                className="primary-button w-full flex items-center justify-center gap-3"
                                            >
                                                {isProcessing ? <Loader2 className="w-4 h-4 animate-spin" /> : <ShieldCheck className="w-4 h-4" />}
                                                Apply selected matches and regenerate
                                            </button>
                                        </div>
                                    )}
                                </div>
                            )}

                            <div className="glass-card rounded-[32px] overflow-hidden flex-1 border-white/5 bg-slate-950/40 relative">
                                {/* Format Detection Panel */}
                                {uploadData?.formatDetection && Object.keys(uploadData.formatDetection).length > 0 && (
                                    <div className="px-10 py-6 border-b border-white/5 space-y-3">
                                        {Object.entries(uploadData.formatDetection as Record<string, { detectedCustomer: string; detectedFormat: string; unmappedColumns: string[] }>).map(([filename, fd]) => (
                                            <div key={filename} className="flex flex-col gap-2">
                                                <div className="flex items-center gap-3">
                                                    <span className="text-[10px] font-black uppercase tracking-[0.3em] text-slate-500">{filename}</span>
                                                    <span className="text-[10px] font-black uppercase tracking-[0.3em] text-blue-400 bg-blue-500/10 border border-blue-500/20 px-3 py-1 rounded-lg">{fd.detectedFormat}</span>
                                                </div>
                                                {fd.unmappedColumns.length > 0 ? (
                                                    <div className="flex flex-wrap gap-2 items-center">
                                                        <span className="text-[9px] font-black uppercase tracking-widest text-amber-500">Unmapped columns:</span>
                                                        {fd.unmappedColumns.map((col: string, colIdx: number) => (
                                                            <span key={`${col}-${colIdx}`} className="text-[9px] font-mono bg-amber-500/10 border border-amber-500/20 text-amber-400 px-2 py-0.5 rounded">{col}</span>
                                                        ))}
                                                    </div>
                                                ) : (
                                                    <span className="text-[9px] font-black uppercase tracking-widest text-emerald-500">All columns mapped</span>
                                                )}
                                            </div>
                                        ))}
                                    </div>
                                )}
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
                            className="text-center max-w-4xl mx-auto space-y-10"
                        >
                            <div className="relative mx-auto w-40 h-40">
                                <div className="absolute inset-0 bg-emerald-500/20 blur-[60px] animate-pulse rounded-full" />
                                <div className="relative w-full h-full bg-slate-900 border border-emerald-500/30 rounded-[35%] flex items-center justify-center shadow-2xl">
                                    <FileCheck className="w-20 h-20 text-emerald-400 transition-transform group-hover:scale-110" />
                                </div>
                            </div>

                            <div className="space-y-6">
                                <h2 className="text-4xl md:text-5xl font-black tracking-tight text-white">Review Results</h2>
                                <p className="text-slate-400">Confirm the summary below before downloading the NextGen Excel files.</p>
                                <div className="grid grid-cols-2 md:grid-cols-3 gap-3 text-left">
                                    {[
                                        { label: "Brand detected", value: detectedBrands.join(", ") || manualBrand || "Not detected" },
                                        { label: "Buyer file", value: buyFiles?.[0]?.name || uploadData?.fileSummary?.[0]?.filename || "Uploaded file" },
                                        { label: "Purchase orders", value: uploadData?.mergedSummary?.orders || 0 },
                                        { label: "Products matched", value: `${variantSummary.resolved || 0} / ${variantSummary.requested || 0}` },
                                        { label: "Unresolved items", value: attentionItems.length },
                                        { label: "Transport methods", value: transportMethods.join(", ") || "From buy file" },
                                    ].map((item) => (
                                        <div key={item.label} className="rounded-2xl border border-white/10 bg-white/5 p-4 min-w-0">
                                            <p className="text-[9px] font-black uppercase tracking-[0.2em] text-slate-500 mb-2">{item.label}</p>
                                            <p className="text-sm font-black text-white break-words">{item.value}</p>
                                        </div>
                                    ))}
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
                                    <span>Continue to Downloads</span> <Download className="w-6 h-6 group-hover:-translate-y-1 transition-transform" />
                                </button>
                                <button
                                    onClick={() => setCurrentStep("VALIDATE")}
                                    className="flex items-center gap-3 text-slate-600 hover:text-white transition-all text-[11px] font-black tracking-[0.4em] uppercase group"
                                >
                                    <ChevronLeft className="w-5 h-5 group-hover:-translate-x-1 transition-transform" /> Back to Validation
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
                                <h2 className="text-4xl font-black tracking-tight text-white">Excel Files Ready</h2>
                                <p className="text-slate-400 text-lg font-medium">Download all three files together or choose an individual file.</p>
                            </div>

                            <div className="max-w-xl mx-auto">
                                <button
                                    onClick={handleDownloadAll}
                                    className="primary-button w-full py-6 flex items-center justify-center gap-4 text-sm"
                                >
                                    <PackageCheck className="w-6 h-6" />
                                    Download All Excel Files
                                </button>
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
                                            <span className="relative z-10">Download {file.label}</span>
                                            <div className="absolute inset-0 bg-blue-500 translate-y-full group-hover/btn:translate-y-0 transition-transform duration-300" />
                                        </button>
                                    </motion.div>
                                ))}
                            </div>

                            <div className="flex flex-col items-center gap-6">
                                <div className="flex justify-center gap-10">
                                    <button
                                        onClick={() => setCurrentStep("REVIEW")}
                                        className="flex items-center gap-3 text-slate-600 hover:text-white transition-all text-[11px] font-black tracking-[0.4em] uppercase group"
                                    >
                                        <ChevronLeft className="w-5 h-5 group-hover:-translate-x-1 transition-transform" /> Back to Review
                                    </button>
                                    <button
                                        onClick={startAnotherFile}
                                        className="flex items-center gap-3 text-slate-600 hover:text-white transition-all text-[11px] font-black tracking-[0.4em] uppercase group"
                                    >
                                        <History className="w-5 h-5 group-hover:-rotate-90 transition-transform duration-500" /> Process Another File
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
