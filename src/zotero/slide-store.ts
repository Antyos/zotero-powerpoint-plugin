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
const CITATION_TAG_KEY = "ZOTERO_CITATIONS";

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

async function getCurrentSlide(context: PowerPoint.RequestContext): Promise<PowerPoint.Slide> {
  // Get the selected slide or fallback to first slide
  let currentSlide: PowerPoint.Slide;

  try {
    const selectedSlides = context.presentation.getSelectedSlides();
    selectedSlides.load("items");
    await context.sync();

    if (selectedSlides.items.length > 0) {
      currentSlide = selectedSlides.items[0];
    } else {
      const allSlides = context.presentation.slides;
      allSlides.load("items");
      await context.sync();
      currentSlide = allSlides.items[0];
    }
  } catch {
    const allSlides = context.presentation.slides;
    allSlides.load("items");
    await context.sync();
    currentSlide = allSlides.items[0];
  }

  return currentSlide;
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
          (existingCitation) => existingCitation.key.toString() !== citation.key
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
   * Get a citation by key using the new CitationStoreXml schema
   */
  public async get(key: string): Promise<ZoteroItemData | null> {
    try {
      const citations = await this.getAll();
      // Find the citation with the matching key
      const citation = citations.find((cit) => cit.key.toString() === key);

      return citation ?? null;
    } catch (error) {
      throw new Error(`Failed to get citation: ${error}`);
    }
  }

  /**
   * Get all citations from the store using the new CitationStoreXml schema
   */
  public async getAll(): Promise<ZoteroItemData[]> {
    return await PowerPoint.run(async (context) => {
      try {
        const xmlPart = await this.getOrCreateCustomXmlPart(context);
        const storeXml = await this.getCitationStoreXml(context, xmlPart);

        // Return all citation data from the store
        return storeXml.citations.citation;
      } catch (error) {
        throw new Error(`Failed to get all citations: ${error}`);
      }
    });
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
        storeXml.citations.citation = storeXml.citations.citation.filter(
          (cit) => cit.key.toString() !== key
        );

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
}

/**
 * ===== Slide-level citation key management functions =====
 */

/**
 * Add a citation key to the current slide
 */
export async function addCitationToSlide(citationKey: string): Promise<void> {
  return new Promise((resolve, reject) => {
    PowerPoint.run(async (context) => {
      try {
        const slide = await getCurrentSlide(context);
        slide.tags.add(CITATION_TAG_KEY, citationKey);
        await context.sync();
        resolve();
      } catch (error) {
        reject(new Error(`Failed to add citation to slide: ${error}`));
      }
    });
  });
}

/**
 * Remove a citation key from the current slide
 */
export async function removeCitationFromSlide(citationKey: string): Promise<boolean> {
  return new Promise((resolve) => {
    PowerPoint.run(async (context) => {
      try {
        const slide = await getCurrentSlide(context);

        // Try to get and remove the specific tag
        // Note: PowerPoint Tag API may not support direct deletion
        // This is a workaround - we'll store an empty value to mark as deleted
        try {
          const citationTag = slide.tags.getItem(CITATION_TAG_KEY);
          citationTag.load("value");
          await context.sync();

          const currentTags = citationTag.value.split(",");
          const updatedTags = currentTags.filter((tag) => tag !== citationKey);
          slide.tags.add(CITATION_TAG_KEY, updatedTags.join(","));
          await context.sync();
        } catch {
          resolve(false);
        }

        resolve(true);
      } catch {
        // Any error is fine for removal
        resolve(false);
      }
    });
  });
}

/**
 * Dump slide tags to console for debugging.
 */
export async function debugSlideTags(): Promise<void> {
  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      slide.tags.load("key,value");
      await context.sync();

      console.log("Slide tags:");
      slide.tags.items.forEach((tag) => {
        console.log(`Key: ${tag.key}, Value: ${tag.value}`);
      });
    });
  } catch (error) {
    console.error("Failed to dump slide tags:", error);
  }
}

/**
 * Get all citation keys from the current slide
 */
export async function getCitationKeysOnSlide(
  slide: PowerPoint.Slide | "current"
): Promise<string[]> {
  return await PowerPoint.run(async (context) => {
    if (slide === "current") {
      slide = await getCurrentSlide(context);
    }
    const tags = slide.tags;
    await context.sync();
    // console.log(`Slide Tags: ${JSON.stringify(tags, null, 2)}`);
    const citationTag = tags.getItemOrNullObject(CITATION_TAG_KEY);
    citationTag.load("value");
    await context.sync();

    if (citationTag.isNullObject || !citationTag.value) {
      return []; // No citations on this slide
    }

    const tagValue = citationTag.value;
    // Split the tag value by commas to get individual citation keys
    return tagValue.split(",");
  });
}

/**
 * High-level convenience functions
 */

/**
 * Insert a citation into the document and reference it on the current slide
 */
export async function insertCitationOnSlide(citation: ZoteroItemData): Promise<void> {
  const store = CitationStore.getInstance();

  // Add citation to the document store
  await store.add(citation);

  // Add citation key to the current slide
  await addCitationToSlide(citation.key);
}

/**
 * Get all citations referenced on the current slide
 */
export async function getCitationsOnSlide(): Promise<ZoteroItemData[]> {
  const store = CitationStore.getInstance();
  const citationKeys = await getCitationKeysOnSlide("current");

  const citations: ZoteroItemData[] = [];
  for (const key of citationKeys) {
    const citation = await store.get(key);
    if (citation) {
      citations.push(citation);
    }
  }

  return citations;
}
