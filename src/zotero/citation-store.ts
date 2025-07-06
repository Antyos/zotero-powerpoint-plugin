/**
 * Citation storage implementation using Office.document.customXmlParts
 *
 * This module implements a robust, persistent, and scalable citation storage model:
 * - All full citation data is stored in a document-level XML part under <ZoteroCitations>
 * - Slides reference citations by key using PowerPoint tags
 * - CitationStore class provides add, get, getAll, and remove methods
 * - Helper functions manage citation keys on slides
 */

import { ZoteroItemData } from "zotero-api-client";
import { XMLParser, XMLBuilder } from "fast-xml-parser";

const CITATION_XML_NAMESPACE = "http://zotero.org/citations";
export const CITATION_TAG_KEY = "ZOTERO_CITATIONS";

/**
 * Schema interface for the citation XML structure stored in customXmlParts
 */
interface CitationStoreXml {
  citations: {
    citation: ZoteroItemData[];
  };
  version: 1;
  "@_xmlns": typeof CITATION_XML_NAMESPACE;
}

async function debugXmlParts(
  context: PowerPoint.RequestContext,
  customXmlParts: PowerPoint.CustomXmlPartCollection
) {
  try {
    console.log("Custom XML Part Debug Info:");
    customXmlParts.load("items");
    await context.sync();
    const numParts = customXmlParts.items.length;
    console.log(`Total XML Parts: ${numParts}`);
    for (const [index, part] of customXmlParts.items.entries()) {
      console.log(`Part ${index + 1}:`);
      try {
        part.load(["id", "namespaceUri", "getXml"]);
        // Ignoring the warning because it's ok for a debug method.
        // eslint-disable-next-line office-addins/no-context-sync-in-loop
        await context.sync();
        console.log(`  Part ID: ${part.id}`);
        console.log(`  Part Namespace: ${part.namespaceUri}`);
        const xmlData = part.getXml();
        // eslint-disable-next-line office-addins/no-context-sync-in-loop
        await context.sync();
        // context.sync() is called.
        // eslint-disable-next-line office-addins/load-object-before-read
        const xmlValue = xmlData.value;
        console.log(`  Part XML: ${xmlValue}`);
      } catch (error) {
        console.error(`  Error loading:`, error);
      }
    }
    console.log("XML Parts debug completed.");
  } catch (error) {
    console.error("Error debugging XML parts:", error);
  }
}

/**
 * CitationStore class for managing citations in Office.document.customXmlParts
 */
export class CitationStore {
  private static instance: CitationStore | null = null;
  private xmlPartId: string = "ZoteroCitations";
  private xmlParser: XMLParser;
  private xmlBuilder: XMLBuilder;

  private constructor() {
    // Configure XML parser options
    this.xmlParser = new XMLParser({
      ignoreAttributes: false,
      attributeNamePrefix: "@_",
      textNodeName: "#text",
      parseAttributeValue: false,
      parseTagValue: true,
      trimValues: true,
    });

    // Configure XML builder options
    this.xmlBuilder = new XMLBuilder({
      ignoreAttributes: false,
      attributeNamePrefix: "@_",
      textNodeName: "#text",
      format: true,
      indentBy: "  ",
      suppressEmptyNode: false,
      suppressBooleanAttributes: false,
    });
  }

  /**
   * Get singleton instance of CitationStore
   */
  public static getInstance(): CitationStore {
    if (!CitationStore.instance) {
      CitationStore.instance = new CitationStore();
    }
    return CitationStore.instance;
  }

  /**
   * Get the custom XML part for citations in `PowerPoint.RequestContext.presentation`. If it
   * doesn't exist, create it.
   */
  public async getOrCreateCustomXmlPart(
    context: PowerPoint.RequestContext
  ): Promise<PowerPoint.CustomXmlPart> {
    const customXmlParts = context.presentation.customXmlParts;
    // await debugXmlParts(context, customXmlParts);
    customXmlParts.load("items");
    await context.sync();

    // Maybe replace with getByNamespace() to check
    for (const part of customXmlParts.items) {
      part.load(["namespaceUri"]);
    }
    // Don't use context.sync() in a loop, it can cause performance issues
    await context.sync();
    for (const part of customXmlParts.items) {
      if (part.namespaceUri === CITATION_XML_NAMESPACE) {
        return part;
      }
    }
    // If no existing part, create a new one
    const xmlString = `<?xml version="1.0" encoding="UTF-8"?><${this.xmlPartId} xmlns="${CITATION_XML_NAMESPACE}"></${this.xmlPartId}>`;
    const newPart = customXmlParts.add(xmlString);
    return newPart;
  }

  /**
   * Helper function to get XML content from a PowerPoint CustomXmlPart
   */
  private async xmlPartToString(
    context: PowerPoint.RequestContext,
    xmlPart: PowerPoint.CustomXmlPart
  ): Promise<string> {
    const xmlData = xmlPart.getXml();
    await context.sync();
    // eslint-disable-next-line office-addins/load-object-before-read
    return xmlData.value;
  }

  /**
   * Get the parsed citation store XML structure
   */
  private async getCitationStoreXml(
    context: PowerPoint.RequestContext,
    xmlPart?: PowerPoint.CustomXmlPart
  ): Promise<CitationStoreXml> {
    xmlPart = xmlPart ?? (await this.getOrCreateCustomXmlPart(context));
    const xmlContent = await this.xmlPartToString(context, xmlPart);
    const parsedXml = this.xmlParser.parse(xmlContent) as any;
    const storeXml = parsedXml[this.xmlPartId] ?? {
      "@_xmlns": CITATION_XML_NAMESPACE,
      citations: { citation: [] },
      version: 1,
    };
    if (!storeXml.citations) {
      storeXml.citations = { citation: [] };
    } else if (!storeXml.citations.citation) {
      storeXml.citations.citation = [];
    } else if (!Array.isArray(storeXml.citations.citation)) {
      storeXml.citations.citation = [storeXml.citations.citation];
    }
    // Ensure creators is always an array
    for (const citation of storeXml.citations.citation) {
      if (citation.key && typeof citation.key !== "string") {
        citation.key = citation.key.toString();
      }
      if (citation.creators && !Array.isArray(citation.creators)) {
        citation.creators = [citation.creators];
      }
    }
    return storeXml as CitationStoreXml;
  }

  /**
   * Save the citation store XML structure back to the document
   */
  private async saveCitationStoreXml(
    context: PowerPoint.RequestContext,
    newXmlContent: CitationStoreXml,
    xmlPart?: PowerPoint.CustomXmlPart
  ): Promise<void> {
    xmlPart = xmlPart ?? (await this.getOrCreateCustomXmlPart(context));
    try {
      const rootObj = { [this.xmlPartId]: newXmlContent };
      const updatedXml = this.xmlBuilder.build(rootObj);
      xmlPart.setXml(updatedXml);
      await context.sync();
    } catch (error) {
      throw new Error(`Failed to save store XML: ${error}`);
    }
  }

  /**
   * Add a citation to the store using fast-xml-parser
   */
  public async add(citation: ZoteroItemData): Promise<void> {
    await PowerPoint.run(async (context) => {
      try {
        const xmlPart = await this.getOrCreateCustomXmlPart(context);
        const storeXml = await this.getCitationStoreXml(context, xmlPart);

        // Remove existing citation with same key
        storeXml.citations.citation = storeXml.citations.citation.filter(
          (existingCitation) => existingCitation.key !== citation.key
        );

        // Add the new citation
        storeXml.citations.citation.push(citation);

        // Save the updated store structure
        await this.saveCitationStoreXml(context, storeXml, xmlPart);

        console.log(`Citation added: ${citation.key}`);
      } catch (error) {
        throw new Error(`Failed to add citation: ${error}`);
      }
    });
  }

  /**
   * Get all citations from the store using the new CitationStoreXml schema
   */
  public async getAll(): Promise<Map<string, ZoteroItemData>> {
    return await PowerPoint.run(async (context) => {
      try {
        const storeXml = await this.getCitationStoreXml(context);

        // Return all citation data from the store as a Map (more concise approach)
        return new Map(storeXml.citations.citation.map((citation) => [citation.key, citation]));
      } catch (error) {
        throw new Error(`Failed to get all citations: ${error}`);
      }
    });
  }

  /**
   * Get a citation by key using the new CitationStoreXml schema
   */
  public async getItem(key: string): Promise<ZoteroItemData | null>;
  public async getItem(key: string[]): Promise<Array<ZoteroItemData | null>>;
  public async getItem(
    key: string | string[]
  ): Promise<ZoteroItemData | null | Array<ZoteroItemData | null>> {
    const items = await this.getAll();
    if (Array.isArray(key)) {
      return key.map((k) => items.get(k) ?? null);
    }
    return items.get(key) ?? null;
  }

  /**
   * Dump all citations to console for debugging
   */
  public async debugXml(): Promise<void> {
    await PowerPoint.run(async (context) => {
      const customXmlParts = context.presentation.customXmlParts;
      await debugXmlParts(context, customXmlParts);
    });
  }

  public async debugCitations(): Promise<void> {
    console.log("Current citations:");
    const citations = await this.getAll();
    console.log(citations);
  }

  /**
   * Remove a citation by key using the new CitationStoreXml schema
   */
  public async remove(key: string): Promise<boolean> {
    return await PowerPoint.run(async (context) => {
      try {
        const xmlPart = await this.getOrCreateCustomXmlPart(context);
        const storeXml = await this.getCitationStoreXml(context, xmlPart);

        // Find and remove the citation with the matching key
        const originalLength = storeXml.citations.citation.length;
        storeXml.citations.citation = storeXml.citations.citation.filter((cit) => cit.key !== key);

        if (storeXml.citations.citation.length < originalLength) {
          // Citation was found and removed, save the updated store
          await this.saveCitationStoreXml(context, storeXml, xmlPart);
          return true;
        } else {
          return false; // Citation not found
        }
      } catch (error) {
        throw new Error(`Failed to remove citation: ${error}`);
      }
    });
  }

  /**
   * Clear all citations from the store
   */
  public async clearStore(): Promise<void> {
    await PowerPoint.run(async (context) => {
      try {
        const xmlPart = await this.getOrCreateCustomXmlPart(context);
        xmlPart.delete();
        await context.sync();
        console.log("All citations cleared.");
      } catch (error) {
        throw new Error(`Failed to clear citations: ${error}`);
      }
    });
  }

  /**
   * Prune citations that are no longer referenced by any slide
   */
  public async prune(): Promise<number> {
    return await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      const usedKeys = new Set<string>();
      slides.load("items/tags");
      await context.sync();
      const citationTags = [];
      for (const slide of context.presentation.slides.items) {
        const citationTag = slide.tags.getItemOrNullObject(CITATION_TAG_KEY);
        citationTag.load("value");
        citationTags.push(citationTag);
      }
      await context.sync();
      for (const citationTag of citationTags) {
        if (citationTag.isNullObject || !citationTag.value) {
          continue; // No citations on this slide
        }
        for (const tag of citationTag.value.split(",")) {
          usedKeys.add(tag);
        }
      }
      await context.sync();
      const xmlPart = await this.getOrCreateCustomXmlPart(context);
      const storeXml = await this.getCitationStoreXml(context, xmlPart);
      const originalCount = storeXml.citations.citation.length;
      storeXml.citations.citation = storeXml.citations.citation.filter((cit) =>
        usedKeys.has(cit.key)
      );
      await this.saveCitationStoreXml(context, storeXml, xmlPart);
      return originalCount - storeXml.citations.citation.length;
    });
  }
}
