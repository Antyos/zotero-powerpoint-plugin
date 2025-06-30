/*
 * Zotero Better BibTeX Connector for PowerPoint Integration
 * Based on obsidian-zotero-integration BBT implementation
 */

import api, { ZoteroApi } from "zotero-api-client";

interface ZoteroField {
  id: string;
  citationKey: string;
  formattedText: string;
  shapeId: string;
}

/**
 * User configuration for Zotero integration
 */
interface ZoteroConfig {
  apiKey: string;
  userId: number;
  userType?: 'user' | 'group';
}

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
      console.log('Zotero API client initialized');
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
      console.log('Zotero collections:', collections);
      this.isConnected = true;
      return true;
    } catch (error) {
      console.error('Error checking Better BibTeX connection:', error);
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

  /**
   * Configure Zotero user credentials and save them persistently
   */
  async updateConfig(config: ZoteroConfig): Promise<void> {
    try {
      this.config = { ...config };

      await this.saveConfig();
      console.log(`Configured Zotero user: ${config.userId} (${config.userType})`);

      // Test the connection with new credentials
      await this.checkConnection();
    } catch (error) {
      console.error('Error configuring user:', error);
      throw new Error(`Failed to configure Zotero user: ${error}`);
    }
  }

  /**
   * Save settings to PowerPoint document storage
   */
  private async saveConfig(): Promise<void> {
    try {
      const partitionKey = Office.context.partitionKey || "default";
      localStorage.setItem(`${partitionKey}-zotero-settings`, JSON.stringify(this.config));
      console.log('Zotero settings saved successfully');
    } catch (error) {
      console.error('Error saving Zotero settings:', error);
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
        console.log('Zotero settings loaded:', this.config);
      } else {
        console.log('No Zotero settings found, using defaults');
      }
      
      console.log(`Loaded Zotero settings.`);
    } catch (error) {
      console.error('Error loading Zotero settings:', error);
      // Don't throw error - we can work without stored settings
    }
  }

  /**
   * Validate Zotero API key format
   */
  private validateApiKey(apiKey: string): boolean {
    // Zotero API keys are typically 28 characters long and alphanumeric
    return typeof apiKey === 'string' && apiKey.length > 0 && /^[A-Za-z0-9]+$/.test(apiKey);
  }

}
