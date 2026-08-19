/**
 * Brand configuration loader.
 *
 * Centralizes all brand-specific settings (customer name, supplier, templates,
 * key users, plant mappings, transport, country codes) into a single JSON file
 * so new brands can be added without code changes.
 *
 * Usage:
 *   import { getBrandConfig, lookupBrand, getCustomerName, getSupplier } from '@/lib/brand-config';
 *
 * The config is loaded once from src/config/brand-config.json and cached.
 */

import brandConfigData from '@/config/brand-config.json';

export interface BrandKeyUsers {
    k1?: string; k2?: string; k3?: string;
    k4?: string; k5?: string; k6?: string;
    k7?: string; k8?: string;
}

export interface BrandEntry {
    aliases: string[];
    customerName: string;
    customerSubtypes?: Record<string, string>;
    defaultSupplier: string;
    ordersTemplate: string;
    linesTemplate: string;
    keyUsers: BrandKeyUsers;
    stylePrefixStrip?: string[];
    plantAliases?: Record<string, string>;
    rules?: import('@/lib/brand-rules').BrandRules;
}

interface BrandConfig {
    brands: Record<string, BrandEntry>;
    factoryCodes: Record<string, string>;
    plantCountryMap: Record<string, string>;
    transportMap: Record<string, string>;
    validTransportValues: string[];
    countryNameMap: Record<string, string>;
}

const config = brandConfigData as unknown as BrandConfig;

// Build a reverse lookup: alias -> brandKey
const aliasIndex: Map<string, string> = new Map();
for (const [brandKey, entry] of Object.entries(config.brands)) {
    aliasIndex.set(brandKey.toLowerCase(), brandKey);
    for (const alias of entry.aliases) {
        aliasIndex.set(alias.toLowerCase(), brandKey);
    }
}

/**
 * Get the full brand config.
 */
export function getBrandConfig(): BrandConfig {
    return config;
}

/**
 * Look up a brand by any of its aliases (case-insensitive).
 * Returns the brand key (e.g. "tnf") or null if not found.
 */
export function lookupBrand(name: string): string | null {
    const key = String(name || '').toLowerCase().trim();
    return aliasIndex.get(key) || null;
}

/**
 * Get a brand entry by name or alias.
 */
export function getBrand(name: string): BrandEntry | null {
    const brandKey = lookupBrand(name);
    if (!brandKey) return null;
    return config.brands[brandKey] || null;
}

/**
 * Get the customer display name for a brand.
 * Falls back to the raw name if brand is not configured.
 */
export function getCustomerName(brandName: string): string {
    const brand = getBrand(brandName);
    return brand?.customerName || brandName;
}

/**
 * Get the default supplier/factory for a brand.
 */
export function getSupplier(brandName: string): string | null {
    const brand = getBrand(brandName);
    return brand?.defaultSupplier || null;
}

/**
 * Get the orders template for a brand.
 */
export function getOrdersTemplate(brandName: string): string | null {
    const brand = getBrand(brandName);
    return brand?.ordersTemplate || null;
}

/**
 * Get the lines template for a brand.
 */
export function getLinesTemplate(brandName: string): string | null {
    const brand = getBrand(brandName);
    return brand?.linesTemplate || null;
}

/**
 * Get the key users for a brand.
 * Returns a full KeyUsers object with all k1-k8 fields.
 */
export function getKeyUsers(brandName: string): BrandKeyUsers {
    const brand = getBrand(brandName);
    return brand?.keyUsers || {};
}

/**
 * Get the customer subtype mapping for a brand (e.g. TNF In-Line vs RTO vs SMU).
 */
export function getCustomerSubtypes(brandName: string): Record<string, string> {
    const brand = getBrand(brandName);
    return brand?.customerSubtypes || {};
}

/**
 * Get the style prefix strings to strip for a brand (e.g. ["NF00", "NF0", "NF"] for TNF).
 */
export function getStylePrefixStrip(brandName: string): string[] {
    const brand = getBrand(brandName);
    return brand?.stylePrefixStrip || [];
}

/**
 * Resolve a factory code to a factory name.
 */
export function resolveFactoryCode(code: string): string | null {
    const key = String(code || '').toLowerCase().trim();
    return config.factoryCodes[key] || null;
}

/**
 * Resolve a plant code or name to a country.
 * Checks brand-specific plant aliases first, then the global plant-country map.
 */
export function resolvePlantToCountry(plant: string): string | null {
    const key = String(plant || '').toLowerCase().trim();
    if (!key) return null;
    return config.plantCountryMap[key] || null;
}

/**
 * Normalize a transport method string to a canonical value (Sea/Air/Courier).
 */
export function normalizeTransport(value: string): string | null {
    const key = String(value || '').toLowerCase().trim();
    if (!key) return null;
    return config.transportMap[key] || null;
}

/**
 * Get the set of valid transport values.
 */
export function getValidTransportValues(): Set<string> {
    return new Set(config.validTransportValues);
}

/**
 * Normalize a country code or name to a canonical country name.
 */
export function normalizeCountryName(value: string): string | null {
    const key = String(value || '').trim();
    if (!key) return null;
    // Try exact match first (case-sensitive for ISO codes)
    if (config.countryNameMap[key]) return config.countryNameMap[key];
    // Try uppercase
    const upper = key.toUpperCase();
    if (config.countryNameMap[upper]) return config.countryNameMap[upper];
    // Try lowercase
    const lower = key.toLowerCase();
    if (config.countryNameMap[lower]) return config.countryNameMap[lower];
    return null;
}

/**
 * Get all configured brand keys.
 */
export function getAllBrandKeys(): string[] {
    return Object.keys(config.brands);
}

/**
 * Get all brand aliases (for detection/matching).
 */
export function getAllBrandAliases(): string[] {
    const aliases: string[] = [];
    for (const entry of Object.values(config.brands)) {
        aliases.push(...entry.aliases);
    }
    return aliases;
}
