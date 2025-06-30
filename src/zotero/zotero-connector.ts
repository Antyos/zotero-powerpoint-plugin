/*
 * Zotero Better BibTeX Connector for PowerPoint Integration
 * Based on obsidian-zotero-integration BBT implementation
 */

import api, { SingleReadResponse, ZoteroApi, RequestOptions } from "zotero-api-client";

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
      console.log('Checking connection... Collections:', collections);
      this.isConnected = true;
      return true;
    } catch (error) {
      console.error('Error checking Zotero connection:', error);
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


  public async getItems(opts?: RequestOptions): Promise<ZoteroField[]> {
    try {
      const response = await ZoteroLibrary.getClient().items().get(opts);
      const itemData = response instanceof SingleReadResponse ? [response.getData()] : response.getData();
      console.log('Fetched Zotero items:', itemData);
      return itemData.map(item => ({
        id: item.key,
        citationKey: item.data?.extra || '',
        formattedText: item.data?.title || '',
        shapeId: ''
      }));
    }
    catch (error) {
      console.error('Error getting Zotero items:', error);
      throw new Error(`Failed to get items: ${error}`);
    }
  }

  /**
   * Get current configuration (excluding sensitive data like API key)
   */
  getConfig(): Omit<ZoteroConfig, 'apiKey'> & { hasApiKey: boolean } {
    return {
      userId: this.config?.userId || 0,
      userType: this.config?.userType || 'user',
      hasApiKey: !!(this.config?.apiKey)
    };
  }

  /**
   * Check if configuration is complete
   */
  isConfigured(): boolean {
    return !!(this.config?.apiKey && this.config?.userId);
  }

  /**
   * Open configuration dialog
   */
  async openConfigDialog(): Promise<ZoteroConfig | null> {
    return new Promise((resolve, reject) => {
      try {
        const dialogUrl = window.location.origin + '/config-dialog.html';
        console.log('Opening configuration dialog at:', dialogUrl);
        
        Office.context.ui.displayDialogAsync(
          dialogUrl,
          { height: 70, width: 50 },
          (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
              console.error('Failed to open dialog:', result.error);
              reject(new Error(`Failed to open dialog: ${result.error.message}`));
              return;
            }
            
            const dialog = result.value;
            
            // Handle messages from the dialog
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
              try {
                const messageArgs = args as { message: string; origin: string | undefined; };
                const message = JSON.parse(messageArgs.message);
                console.log('Dialog message received:', message);
                
                if (message.type === 'config-saved') {
                  // Update configuration with new data
                  this.updateConfig(message.config).then(() => {
                    dialog.close();
                    resolve(message.config);
                  }).catch((error) => {
                    dialog.close();
                    reject(error);
                  });
                } else if (message.type === 'config-cancelled') {
                  dialog.close();
                  resolve(null);
                } else if (message.type === 'config-error') {
                  dialog.close();
                  reject(new Error(message.error));
                }
              } catch (error) {
                console.error('Error parsing dialog message:', error);
                dialog.close();
                reject(new Error('Invalid message from dialog'));
              }
            });
            
            // Handle dialog closed event
            dialog.addEventHandler(Office.EventType.DialogEventReceived, (args) => {
              const eventArgs = args as { error: number; };
              console.log('Dialog event received:', eventArgs.error);
              if (eventArgs.error === 12006) { // Dialog closed by user
                resolve(null);
              } else {
                reject(new Error(`Dialog error: ${eventArgs.error}`));
              }
            });
          }
        );
      } catch (error) {
        console.error('Error opening config dialog:', error);
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
      console.error('Configuration dialog failed:', error);
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
if (typeof window !== 'undefined') {
  window.ZoteroLibrary = ZoteroLibrary;
  // Also expose the singleton instance for easy access
  window.zotero = ZoteroLibrary.getInstance();
}
