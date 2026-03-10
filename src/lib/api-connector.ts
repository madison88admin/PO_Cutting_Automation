export interface ApiConnectorConfig {
    baseUrl: string;
    apiKey: string;
    mode: "UPLOAD" | "API";
}

export class ApiConnector {
    private config: ApiConnectorConfig;

    constructor(config: ApiConnectorConfig) {
        this.config = config;
    }

    /**
     * Placeholder for future API extraction.
     * This module is currently inactive by default.
     */
    async fetchData() {
        if (this.config.mode !== "API") {
            throw new Error("Connector is in UPLOAD mode. API fetching is disabled.");
        }

        console.log("Fetching from API:", this.config.baseUrl);
        // Future implementation: call external API and map to buy file format
        return [];
    }

    async toggleMode(confirm: boolean) {
        if (!confirm) return;
        this.config.mode = this.config.mode === "API" ? "UPLOAD" : "API";
        return this.config.mode;
    }
}
