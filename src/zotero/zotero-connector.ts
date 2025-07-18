/*
 * Zotero Better BibTeX Connector for PowerPoint Integration
 * Based on obsidian-zotero-integration BBT implementation
 */

import api, { ZoteroApi, RequestOptions, ZoteroItemData } from "zotero-api-client";

export interface ZoteroField {
  id: string;
  citationKey: string;
  formattedText: string;
  shapeId: string;
}

export interface CitationFormat {
  format: string;
  delimiter?: string;
}

/**
 * User configuration for Zotero integration
 */
interface ZoteroConfig {
  apiKey: string;
  userId: number;
  userType?: "user" | "group";
  citationFormats?: Record<string, CitationFormat>;
  selectedCitationFormat?: string;
  searchResultsLimit?: number;
  citationShapeName?: string;
}

const DEFAULT_CONFIG: Required<Omit<ZoteroConfig, "apiKey">> = {
  userId: 0,
  userType: "user",
  citationFormats: {
    apa: {
      format: "{creator} ({year}). {title}. <i>{journal}</i>, {volume}({issue}), {pages}.",
      delimiter: ";  ",
    },
    mla: {
      format:
        '{creator}. "{title}." <i>{journal}</i>, vol. {volume}, no. {issue}, {year}, pp. {pages}.',
      delimiter: ";  ",
    },
    chicago: {
      format: '{creator}. "{title}." <i>{journal}</i> {volume}, no. {issue} ({year}): {pages}.',
      delimiter: ";  ",
    },
    ieee: {
      format:
        '{creator}, "{title}," <i>{journal}</i>, vol. {volume}, no. {issue}, pp. {pages}, {year}.',
      delimiter: ";  ",
    },
    custom: {
      format: "<b>[{#}] {creator}</b>, {year}, <i>{journalAbbreviation}</i>",
      delimiter: ";  ",
    },
  },
  selectedCitationFormat: "ieee",
  searchResultsLimit: 5,
  citationShapeName: "Citation",
};

/**
 * Better BibTeX Connector for Zotero PowerPoint Integration
 * Simplified version based on obsidian-zotero-integration BBT implementation
 */
export class ZoteroLibrary {
  private static client: ZoteroApi | null = null;
  private static instance: ZoteroLibrary;
  private isConnected = false;

  // Persistent state properties
  private config?: ZoteroConfig;
  private constructor() {}

  static getClient(): ZoteroApi {
    if (!ZoteroLibrary.client) {
      const instance = ZoteroLibrary.getInstance();
      ZoteroLibrary.client = api(instance.config?.apiKey).library("user", instance.config?.userId);
      console.log("Zotero API client initialized");
    }
    return ZoteroLibrary.client;
  }

  static getInstance(): ZoteroLibrary {
    if (!ZoteroLibrary.instance) {
      ZoteroLibrary.instance = new ZoteroLibrary();
    }
    return ZoteroLibrary.instance;
  }

  /**
   * Check if Better BibTeX is available and ready
   * NOTE: This will likely fail due to Chrome security headers that BBT rejects
   */
  async checkConnection(): Promise<boolean> {
    try {
      const client = ZoteroLibrary.getClient();
      const collections = await client.collections().get();
      console.log("Checking connection... Collections:", collections);
      this.isConnected = true;
      return true;
    } catch (error) {
      console.error("Error checking Zotero connection:", error);
      this.isConnected = false;
      return false;
    }
  }

  /**
   * Check if connector is ready
   */
  isReady(): boolean {
    return this.isConnected;
  }

  private hasUserIdAndApiKey(
    config: Partial<ZoteroConfig>
  ): config is Pick<ZoteroConfig, "userId" | "apiKey"> &
    Partial<Omit<ZoteroConfig, "userId" | "apiKey">> {
    return !!(config.apiKey && config.userId && config.userId > 0);
  }

  async updateConfig(config: Partial<ZoteroConfig>): Promise<void> {
    // Merge with existing config, only updating apiKey if it's provided
    let newConfig: ZoteroConfig;
    if (this.config !== undefined) {
      newConfig = { ...this.config, ...config };
    } else if (this.hasUserIdAndApiKey(config)) {
      newConfig = { ...config };
    } else {
      throw new Error("ApiKey and User ID are required for initial configuration");
    }
    this.setConfig(newConfig);
  }

  /**
   * Configure Zotero user credentials and save them persistently
   */
  async setConfig(config: ZoteroConfig): Promise<void> {
    try {
      // Validate citation formats if provided
      if (config.citationFormats && !this.validateCitationFormats(config.citationFormats)) {
        throw new Error(
          "Invalid citation formats structure. Each format must have a 'format' property and optional 'delimiter' property."
        );
      }

      this.config = { ...config };

      await this.saveConfig();
      console.log(`Configured Zotero user: ${config.userId} (${config.userType})`);
      ZoteroLibrary.client = null; // Reset client to force re-initialization

      // Test the connection with new credentials
      await this.checkConnection();
    } catch (error) {
      console.error("Error configuring user:", error);
      throw new Error(`Failed to configure Zotero user: ${error}`);
    }
  }

  /**
   * Save settings to PowerPoint document storage
   */
  private async saveConfig(): Promise<void> {
    try {
      const partitionKey = Office.context.partitionKey || "default";
      const configJson = JSON.stringify(this.config);
      console.log("Saving configuration:", configJson);
      localStorage.setItem(`${partitionKey}-zotero-settings`, configJson);
      console.log("Zotero settings saved successfully to localStorage");
    } catch (error) {
      console.error("Error saving Zotero settings:", error);
      // Check if it's a localStorage quota issue
      if (error instanceof Error && error.name === "QuotaExceededError") {
        throw new Error(
          "Configuration too large for storage. Please reduce the size of your citation formats."
        );
      }
      throw new Error(`Failed to save settings: ${error}`);
    }
  }

  /**
   * Load settings from PowerPoint document storage
   */
  public loadConfig(): void {
    try {
      const partitionKey = Office.context.partitionKey || "default";

      const settingsJson = localStorage.getItem(`${partitionKey}-zotero-settings`);
      if (settingsJson) {
        this.config = JSON.parse(settingsJson);
        console.log("Zotero settings loaded:", this.config);
      } else {
        console.log("No Zotero settings found, using defaults");
      }
    } catch (error) {
      console.error("Error loading Zotero settings:", error);
      // Don't throw error - we can work without stored settings
    }
  }

  /**
   * Validate Zotero API key format
   */
  private validateApiKey(apiKey: string): boolean {
    // Zotero API keys are typically 28 characters long and alphanumeric
    return typeof apiKey === "string" && apiKey.length > 0 && /^[A-Za-z0-9]+$/.test(apiKey);
  }

  /**
   * Validate citation format structure
   */
  private validateCitationFormat(format: any): format is CitationFormat {
    return (
      typeof format === "object" &&
      format !== null &&
      !Array.isArray(format) &&
      typeof format.format === "string" &&
      (format.delimiter === undefined || typeof format.delimiter === "string")
    );
  }

  /**
   * Validate citation formats object
   */
  private validateCitationFormats(formats: any): formats is Record<string, CitationFormat> {
    if (typeof formats !== "object" || formats === null || Array.isArray(formats)) {
      return false;
    }

    for (const [key, value] of Object.entries(formats)) {
      if (typeof key !== "string" || !this.validateCitationFormat(value)) {
        return false;
      }
    }

    return true;
  }

  public async searchItems(
    query: string,
    maxResults?: number,
    opts?: RequestOptions
  ): Promise<ZoteroItemData[]> {
    try {
      const limit = maxResults || this.config?.searchResultsLimit || 5;
      const response = await ZoteroLibrary.getClient()
        .items()
        .get({ ...opts, q: query, itemType: "-attachment", limit: limit });
      const itemData = response.getData();
      console.log("Quick search results:", itemData);
      if (!itemData || itemData.length === 0) {
        return [];
      }
      return itemData;
    } catch (error) {
      console.error("Error performing quick search:", error);
      throw new Error(`Failed to perform quick search: ${error}`);
    }
  }

  /**
   * Get current configuration (excluding sensitive data like API key)
   */
  getConfig(): Omit<ZoteroConfig, "apiKey"> & { hasApiKey: boolean } {
    const mergedConfig = {
      ...DEFAULT_CONFIG,
      ...this.config,
    };
    const { apiKey, ...configWithoutApiKey } = mergedConfig;
    return {
      ...configWithoutApiKey,
      hasApiKey: !!apiKey,
    };
  }

  /**
   * Check if configuration is complete
   */
  isConfigured(): boolean {
    return !!(this.config?.apiKey && this.config?.userId && this.config.userId > 0);
  }

  public getCitationFormat(): CitationFormat {
    const defaultFormat = DEFAULT_CONFIG.citationFormats.default;
    if (this.config?.selectedCitationFormat && this.config.citationFormats) {
      return this.config.citationFormats[this.config.selectedCitationFormat] ?? defaultFormat;
    }
    return defaultFormat;
  }

  /**
   * Get the citation shape name
   */
  getCitationShapeName(): string {
    return this.config?.citationShapeName ?? DEFAULT_CONFIG.citationShapeName;
  }

  /**
   * Set the citation shape name
   */
  setCitationShapeName(name: string): void {
    if (this.config) {
      this.config.citationShapeName = name;
      this.saveConfig();
    }
  }

  /**
   * Open configuration dialog
   */
  async openConfigDialog(): Promise<ZoteroConfig | null> {
    return new Promise((resolve, reject) => {
      try {
        const dialogUrl = window.location.origin + "/config-dialog.html";
        console.log("Opening configuration dialog at:", dialogUrl);

        Office.context.ui.displayDialogAsync(dialogUrl, { height: 70, width: 50 }, (result) => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            console.error("Failed to open dialog:", result.error);
            reject(new Error(`Failed to open dialog: ${result.error.message}`));
            return;
          }

          const dialog = result.value;

          // Handle messages from the dialog
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
            try {
              const messageArgs = args as { message: string; origin: string | undefined };
              const message = JSON.parse(messageArgs.message);
              console.log("Dialog message received:", message);

              if (message.type === "config-saved") {
                // Update configuration with new data
                this.updateConfig(message.config)
                  .then(() => {
                    dialog.close();
                    resolve(message.config);
                  })
                  .catch((error) => {
                    dialog.close();
                    reject(error);
                  });
              } else if (message.type === "config-cancelled") {
                dialog.close();
                resolve(null);
              } else if (message.type === "config-error") {
                dialog.close();
                reject(new Error(message.error));
              }
            } catch (error) {
              console.error("Error parsing dialog message:", error);
              dialog.close();
              reject(new Error("Invalid message from dialog"));
            }
          });

          // Handle dialog closed event
          dialog.addEventHandler(Office.EventType.DialogEventReceived, (args) => {
            const eventArgs = args as { error: number };
            console.log("Dialog event received:", eventArgs.error);
            if (eventArgs.error === 12006) {
              // Dialog closed by user
              resolve(null);
            } else {
              reject(new Error(`Dialog error: ${eventArgs.error}`));
            }
          });
        });
      } catch (error) {
        console.error("Error opening config dialog:", error);
        reject(new Error(`Failed to open configuration dialog: ${error}`));
      }
    });
  }

  /**
   * Open configuration dialog and handle the result
   */
  async configureFromDialog(): Promise<boolean> {
    try {
      const result = await this.openConfigDialog();
      return result !== null; // Returns true if config was saved, false if cancelled
    } catch (error) {
      console.error("Configuration dialog failed:", error);
      return false;
    }
  }
}

// Make ZoteroLibrary available globally for debugging
declare global {
  interface Window {
    ZoteroLibrary: typeof ZoteroLibrary;
    zotero: ZoteroLibrary;
  }
}

// Expose to global scope for REPL debugging
if (typeof window !== "undefined") {
  window.ZoteroLibrary = ZoteroLibrary;
  // Also expose the singleton instance for easy access
  window.zotero = ZoteroLibrary.getInstance();
}
