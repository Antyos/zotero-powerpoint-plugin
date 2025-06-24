/*
 * Zotero Better BibTeX Connector for PowerPoint Integration
 * Based on obsidian-zotero-integration BBT implementation
 */

// Better BibTeX API Types
interface BBTItem {
  key: string;
  version: number;
  itemType: string;
  title: string;
  creators: BBTCreator[];
  date: string;
  DOI?: string;
  ISBN?: string;
  ISSN?: string;
  url?: string;
  abstractNote?: string;
  publicationTitle?: string;
  volume?: string;
  issue?: string;
  pages?: string;
  publisher?: string;
  place?: string;
  edition?: string;
  series?: string;
  seriesNumber?: string;
  numPages?: string;
  language?: string;
  shortTitle?: string;
  archive?: string;
  archiveLocation?: string;
  libraryCatalog?: string;
  callNumber?: string;
  rights?: string;
  extra?: string;
  tags: BBTTag[];
  collections: string[];
  relations: Record<string, string>;
  dateAdded: string;
  dateModified: string;
  uri: string;
  select?: string;
  citationKey?: string;
}

interface BBTCreator {
  creatorType: string;
  firstName?: string;
  lastName?: string;
  name?: string;
}

interface BBTTag {
  tag: string;
  type?: number;
}

interface BBTCollection {
  key: string;
  version: number;
  library: {
    type: string;
    id: number;
    name: string;
    links: Record<string, any>;
  };
  links: Record<string, any>;
  meta: Record<string, any>;
  data: {
    key: string;
    version: number;
    name: string;
    parentCollection?: string;
    relations: Record<string, any>;
  };
}

interface BBTLibrary {
  id: number;
  name: string;
  type: string;
  version: number;
}

interface ZoteroCitation {
  citationKey: string;
  formattedCitation: string;
  item: BBTItem;
}

interface ZoteroField {
  id: string;
  citationKey: string;
  formattedText: string;
  shapeId: string;
}

/**
 * Better BibTeX Connector for Zotero PowerPoint Integration
 * Simplified version based on obsidian-zotero-integration BBT implementation
 */
export class ZoteroBBTConnector {
  private static instance: ZoteroBBTConnector;
  private isConnected = false;
  private bbtPort = 23119;
  private fields: Map<string, ZoteroField> = new Map();

  private constructor() {}

  static getInstance(): ZoteroBBTConnector {
    if (!ZoteroBBTConnector.instance) {
      ZoteroBBTConnector.instance = new ZoteroBBTConnector();
    }
    return ZoteroBBTConnector.instance;
  }
  private get bbtBaseUrl(): string {
    return `http://127.0.0.1:${this.bbtPort}/better-bibtex`;
  }
  /**
   * Check if Better BibTeX is available and ready
   * NOTE: This will likely fail due to Chrome security headers that BBT rejects
   */
  async checkConnection(): Promise<boolean> {
    try {
      console.log('Checking Better BibTeX connection...');
      console.log('NOTE: This will likely fail due to Office.js localhost restrictions');
      
      // Create custom headers that are more likely to work in Office.js
      const customHeaders = {
        'X-Custom-Client': 'ZoteroPowerPointIntegration',
        'X-Client-Version': '1.0',
        'X-Integration-Type': 'PowerPoint-BBT',
        'Accept': 'text/plain',
        'X-Requested-With': 'XMLHttpRequest',
      };
      
      // Log what we're trying to send
      console.log('Attempting to send headers:', customHeaders);
      
      const response = await fetch(`${this.bbtBaseUrl}/cayw?probe=true`, {
        method: 'GET',
        headers: customHeaders,
        // mode: 'cors', // Use cors mode for more accurate error reporting
      });
      
      // Log the actual request headers that were sent (if possible)
      console.log('Response received:', {
        status: response.status,
        statusText: response.statusText,
        ok: response.ok,
        type: response.type,
        headers: {
          'content-type': response.headers.get('content-type'),
          'content-length': response.headers.get('content-length')
        }
      });

      if (response.ok) {
        const result = await response.text();
        this.isConnected = result.trim() === 'ready';
        console.log(`Better BibTeX status: ${this.isConnected ? 'ready' : 'not ready'}`);
        console.log('Response text:', result);
        return this.isConnected;
      } else {
        this.isConnected = false;
        return false;
      }
    } catch (error) {
      console.error('Error checking Better BibTeX connection (EXPECTED in Office.js):', error);
      console.log('This error is expected - Office.js blocks localhost requests for security');
      this.isConnected = false;
      return false;
    }
  }


  /**
   * Get all libraries from Zotero
   */
  async getLibraries(): Promise<BBTLibrary[]> {
    if (!this.isConnected && !(await this.checkConnection())) {
      throw new Error('Better BibTeX not available');
    }

    try {
      const response = await fetch(`${this.bbtBaseUrl}/json-rpc`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          jsonrpc: '2.0',
          method: 'user.groups',
          id: Date.now()
        })
      });

      if (response.ok) {
        const result = await response.json();
        return result.result || [];
      }

      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    } catch (error) {
      console.error('Error fetching libraries:', error);
      throw error;
    }
  }

  /**
   * Search for items in Zotero using Better BibTeX
   */
  async searchItems(query: string, libraryId?: number): Promise<BBTItem[]> {
    if (!this.isConnected && !(await this.checkConnection())) {
      throw new Error('Better BibTeX not available');
    }

    try {
      const searchParams = new URLSearchParams({
        query: query,
        format: 'json'
      });

      if (libraryId !== undefined) {
        searchParams.append('library', libraryId.toString());
      }      const response = await fetch(`${this.bbtBaseUrl}/search?${searchParams}`, {
        method: 'GET',
        headers: {
          'Accept': 'application/json',
          'X-Requested-With': 'ZoteroPowerPointIntegration',
          'X-Client-Version': '1.0'
        }
      });

      if (response.ok) {
        const items = await response.json();
        return Array.isArray(items) ? items : [];
      }

      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    } catch (error) {
      console.error('Error searching items:', error);
      throw error;
    }
  }

  /**
   * Open Zotero's citation picker using CAYW (Cite As You Write)
   * Returns citation keys of selected items
   */
  async openCitationPicker(format: string = 'citekey'): Promise<string[]> {
    if (!this.isConnected && !(await this.checkConnection())) {
      throw new Error('Better BibTeX not available');
    }

    try {
      console.log('Opening Zotero citation picker...');
      
      const params = new URLSearchParams({
        format: format,
        minimize: 'true'
      });

      const response = await fetch(`${this.bbtBaseUrl}/cayw?${params}`, {
        method: 'GET',
        headers: {
          'Accept': 'text/plain',
          'X-Requested-With': 'ZoteroPowerPointIntegration',
          'X-Client-Version': '1.0'
        }
      });

      if (response.ok) {
        const result = await response.text();
        console.log('CAYW response:', result);
        
        if (result && result.trim() && !result.includes('cancelled')) {
          // Parse citation keys from response
          const citationKeys = result.trim().split(',').map(key => key.trim()).filter(key => key);
          console.log('Selected citation keys:', citationKeys);
          return citationKeys;
        }
      }

      console.log('Citation selection cancelled or no items selected');
      return [];
    } catch (error) {
      console.error('Error opening citation picker:', error);
      throw error;
    }
  }

  /**
   * Get formatted citation for citation keys
   */
  async getFormattedCitation(citationKeys: string[], style: string = 'apa'): Promise<string> {
    if (!this.isConnected && !(await this.checkConnection())) {
      throw new Error('Better BibTeX not available');
    }

    try {
      const params = new URLSearchParams({
        format: 'formatted-citation',
        style: style,
        citekeys: citationKeys.join(',')
      });
      const response = await fetch(`${this.bbtBaseUrl}/cayw?${params}`, {
        method: 'GET',
        headers: {
          'Accept': 'text/plain',
          'X-Requested-With': 'ZoteroPowerPointIntegration',
          'X-Client-Version': '1.0'
        }
      });

      if (response.ok) {
        const citation = await response.text();
        return citation.trim();
      }

      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    } catch (error) {
      console.error('Error getting formatted citation:', error);
      throw error;
    }
  }

  /**
   * Get item details by citation key
   */
  async getItemByCitationKey(citationKey: string): Promise<BBTItem | null> {
    if (!this.isConnected && !(await this.checkConnection())) {
      throw new Error('Better BibTeX not available');
    }

    try {
      const response = await fetch(`${this.bbtBaseUrl}/json-rpc`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          jsonrpc: '2.0',
          method: 'item.search',
          params: [{ citationKey: citationKey }],
          id: Date.now()
        })
      });

      if (response.ok) {
        const result = await response.json();
        const items = result.result;
        return Array.isArray(items) && items.length > 0 ? items[0] : null;
      }

      return null;
    } catch (error) {
      console.error('Error getting item by citation key:', error);
      return null;
    }
  }

  /**
   * Insert citation into PowerPoint slide
   */
  async insertCitation(citationKeys: string[], style: string = 'apa'): Promise<string> {
    if (citationKeys.length === 0) {
      throw new Error('No citation keys provided');
    }

    try {
      // Get formatted citation
      const formattedCitation = await this.getFormattedCitation(citationKeys, style);
      
      if (!formattedCitation) {
        throw new Error('Failed to get formatted citation');
      }

      // Insert into PowerPoint
      return PowerPoint.run(async (context) => {
        const slide = context.presentation.getSelectedSlides().getItemAt(0);
        
        // Create field ID
        const fieldId = `zotero-field-${Date.now()}`;
        
        // Insert text box with citation
        const textBox = slide.shapes.addTextBox(formattedCitation, {
          left: 100,
          top: 100,
          width: 400,
          height: 50
        });
        
        // Style the citation
        textBox.textFrame.textRange.font.name = "Calibri";
        textBox.textFrame.textRange.font.size = 11;
        textBox.fill.clear();
        textBox.lineFormat.color = "#E0E0E0";
        textBox.lineFormat.weight = 0.5;
        
        // Store field ID in shape name
        textBox.name = `ZOTERO_FIELD_${fieldId}`;
        
        await context.sync();
        
        // Store field data
        const field: ZoteroField = {
          id: fieldId,
          citationKey: citationKeys.join(','),
          formattedText: formattedCitation,
          shapeId: textBox.id
        };
        
        this.fields.set(fieldId, field);
        await this.saveFieldData();
        
        console.log(`Inserted citation: ${formattedCitation}`);
        return fieldId;
      });
    } catch (error) {
      console.error('Error inserting citation:', error);
      throw error;
    }
  }

  /**
   * Insert citation using the picker workflow
   */
  async insertCitationWithPicker(style: string = 'apa'): Promise<string | null> {
    try {
      // Open citation picker
      const citationKeys = await this.openCitationPicker();
      
      if (citationKeys.length === 0) {
        console.log('No citations selected');
        return null;
      }

      // Insert the selected citations
      const fieldId = await this.insertCitation(citationKeys, style);
      return fieldId;
    } catch (error) {
      console.error('Error in citation picker workflow:', error);
      throw error;
    }
  }

  /**
   * Refresh all citations in the document
   */
  async refreshCitations(style: string = 'apa'): Promise<void> {
    if (this.fields.size === 0) {
      console.log('No citations to refresh');
      return;
    }

    try {
      await PowerPoint.run(async (context) => {
        for (const [fieldId, field] of Array.from(this.fields.entries())) {
          try {
            // Get updated formatted citation
            const citationKeys = field.citationKey.split(',');
            const updatedCitation = await this.getFormattedCitation(citationKeys, style);
            
            // Find and update the text box
            const shapes = context.presentation.slides.getItemAt(0).shapes; // Simplified - would need proper slide detection
            shapes.load('items');
            await context.sync();
            
            const shape = shapes.items.find(s => s.name === `ZOTERO_FIELD_${fieldId}`);
            if (shape && shape.textFrame) {
              shape.textFrame.textRange.text = updatedCitation;
              field.formattedText = updatedCitation;
            }
          } catch (error) {
            console.error(`Error refreshing field ${fieldId}:`, error);
          }
        }
        
        await context.sync();
        await this.saveFieldData();
      });
      
      console.log('Citations refreshed successfully');
    } catch (error) {
      console.error('Error refreshing citations:', error);
      throw error;
    }
  }

  /**
   * Save field data to document settings
   */
  private async saveFieldData(): Promise<void> {
    try {
      const settings = Office.context.document.settings;
      const fieldsArray = Array.from(this.fields.values());
      settings.set('zotero-fields', fieldsArray);
      await settings.saveAsync();
    } catch (error) {
      console.error('Error saving field data:', error);
    }
  }

  /**
   * Load field data from document settings
   */
  async loadFieldData(): Promise<void> {
    try {
      const settings = Office.context.document.settings;
      const fieldsArray = settings.get('zotero-fields') || [];
      
      this.fields.clear();
      fieldsArray.forEach((field: ZoteroField) => {
        this.fields.set(field.id, field);
      });
      
      console.log(`Loaded ${this.fields.size} Zotero fields`);
    } catch (error) {
      console.error('Error loading field data:', error);
    }
  }

  /**
   * Get all fields in the document
   */
  getFields(): ZoteroField[] {
    return Array.from(this.fields.values());
  }

  /**
   * Check if connector is ready
   */
  isReady(): boolean {
    return this.isConnected;
  }

  /**
   * Test connection and return diagnostic information
   */
  async testConnection(): Promise<string[]> {
    const results: string[] = [];
    
    // Test Better BibTeX connection
    const connected = await this.checkConnection();
    results.push(`Better BibTeX Connection: ${connected ? 'SUCCESS' : 'FAILED'}`);

    if (connected) {
      // Test search functionality
      try {
        const libraries = await this.getLibraries();
        results.push(`Libraries found: ${libraries.length}`);
      } catch (error) {
        results.push(`Libraries test: FAILED (${error.message})`);
      }
      
      // Test search
      try {
        const items = await this.searchItems('test', undefined);
        results.push(`Search test: SUCCESS (${items.length} items found)`);
      } catch (error) {
        results.push(`Search test: FAILED (${error.message})`);
      }
    }
    
    results.push(`Loaded fields: ${this.fields.size}`);
    
    return results;
  }



  /**
   * Debug method to test header behavior in Office.js environment
   */
  async debugHeaders(): Promise<{
    requestAttempt: any;
    responseHeaders: any;
    userAgentIssue: string;
  }> {
    try {
      console.log('=== DEBUGGING HEADERS ===');
      
      // Test what happens when we try to set various headers
      const testHeaders = {
        // 'User-Agent': 'ZoteroPowerPointIntegration/1.0',
        'X-Custom-Client': 'ZoteroPowerPointIntegration',
        'X-Client-Version': '1.0',
        'X-Integration-Type': 'PowerPoint-BBT',
        'Accept': 'text/plain',
        'X-Requested-With': 'XMLHttpRequest',
        "Allow-Control-Allow-Origin": "*",
        "Sec-Fetch-Site": "none",
      };
      
      console.log('Attempting to send these headers:', testHeaders);
      
      const response = await fetch(`${this.bbtBaseUrl}/cayw?probe=true`, {
        method: 'GET',
        headers: testHeaders,
        mode: "no-cors",
      });
        // Collect response headers (Office.js compatible way)
      const responseHeaders: any = {};
      const commonHeaders = ['content-type', 'content-length', 'server', 'date', 'cache-control', 'connection'];
      commonHeaders.forEach(header => {
        const value = response.headers.get(header);
        if (value) {
          responseHeaders[header] = value;
        }
      });
      
      const result = {
        requestAttempt: {
          url: `${this.bbtBaseUrl}/cayw?probe=true`,
          method: 'GET',
          intendedHeaders: testHeaders,
          status: response.status,
          statusText: response.statusText
        },
        responseHeaders,
        userAgentIssue: 'Browsers typically override User-Agent header for security. Office.js may further restrict headers. Use X-Custom-* headers instead.'
      };
      
      console.log('Header debug results:', result);
      return result;
      
    } catch (error) {
      console.error('Header debug failed:', error);
      return {
        requestAttempt: { error: error.message },
        responseHeaders: {},
        userAgentIssue: `Debug failed: ${error.message}`
      };
    }
  }
}
