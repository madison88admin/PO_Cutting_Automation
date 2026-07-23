import { NextGenStyleInfo } from '@/lib/types/buy-file';

export class NextGenCache {
    private styleCache: Map<string, NextGenStyleInfo> = new Map();

    getStyle(style: string): NextGenStyleInfo | undefined {
        return this.styleCache.get(style.toLowerCase().trim());
    }

    setStyle(style: string, info: NextGenStyleInfo): void {
        this.styleCache.set(style.toLowerCase().trim(), info);
    }

    hasStyle(style: string): boolean {
        return this.styleCache.has(style.toLowerCase().trim());
    }

    getUniqueStyles(styles: string[]): string[] {
        return [...new Set(styles.map((s) => s.toLowerCase().trim()))];
    }
}
