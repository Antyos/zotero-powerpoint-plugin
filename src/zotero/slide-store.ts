/**
 * Citation storage implementation using Office.document.customXmlParts
 *
 * This module implements a robust, persistent, and scalable citation storage model:
 * - All full citation data is stored in a document-level XML part under <ZoteroCitations>
 * - Slides reference citations by key using PowerPoint tags
 * - CitationStore class provides add, get, getAll, and remove methods
 * - Helper functions manage citation keys on slides
 */

import { ZoteroItemData, ZoteroCreator, ZoteroTag } from "zotero-api-client";

const CITATION_XML_NAMESPACE = "http://zotero.org/citations";
const CITATION_TAG_KEY = "ZOTERO_CITATIONS";

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

  private constructor() {}

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
  private async getXmlContent(
    context: PowerPoint.RequestContext,
    xmlPart: PowerPoint.CustomXmlPart
  ): Promise<string> {
    // Load the XML part properties we need
    // xmlPart.load(["id", "namespaceUri"]);
    // await context.sync();

    // Now get the XML content
    const xmlData = xmlPart.getXml();
    await context.sync();

    // eslint-disable-next-line office-addins/load-object-before-read
    return xmlData.value;
  }

  // /**
  //  * Helper function to replace XML part content
  //  */
  // private async replaceXmlContent(
  //   context: PowerPoint.RequestContext,
  //   xmlPart: PowerPoint.CustomXmlPart,
  //   newXmlContent: string
  // ): Promise<void> {
  //   xmlPart.setXml(newXmlContent);
  //   await context.sync();
  // }

  /**
   * Add a citation to the store
   */
  public async add(citation: ZoteroItemData): Promise<void> {
    await PowerPoint.run(async (context) => {
      const xmlPart = await this.getOrCreateCustomXmlPart(context);

      let xmlContent = await this.getXmlContent(context, xmlPart);

      // Remove existing citation with same key (simple string replacement)
      const citationRegex = new RegExp(`<citation key="${citation.key}"[^>]*>.*?</citation>`, "g");
      xmlContent = xmlContent.replace(citationRegex, "");

      // Add new citation before closing tag
      const citationXml = this.citationToXml(citation);
      xmlContent = xmlContent.replace(`</${this.xmlPartId}>`, `${citationXml}</${this.xmlPartId}>`);

      // Update the XML part
      xmlPart.setXml(xmlContent);
      await context.sync();

      console.log(`Citation added: ${citation.key}`);
    });
  }

  /**
   * Get a citation by key
   */
  public async get(key: string): Promise<ZoteroItemData | null> {
    return await PowerPoint.run(async (context) => {
      const xmlPart = await this.getOrCreateCustomXmlPart(context);
      const xmlContent = await this.getXmlContent(context, xmlPart);

      try {
        // Use regex to find the citation
        const citationRegex = new RegExp(`<citation key="${key}"[^>]*>(.*?)</citation>`, "g");
        const match = citationRegex.exec(xmlContent);

        if (match) {
          return this.xmlToCitation(match[0]);
        } else {
          return null;
        }
      } catch (error) {
        throw new Error(`Failed to parse XML: ${error}`);
      }
    });
  }

  /**
   * Get all citations from the store
   */
  public async getAll(): Promise<ZoteroItemData[]> {
    return await PowerPoint.run(async (context) => {
      // Get XML content
      const xmlPart = await this.getOrCreateCustomXmlPart(context);
      const xmlData = xmlPart.getXml();
      await context.sync();
      // eslint-disable-next-line office-addins/load-object-before-read
      const xmlContent = xmlData.value;

      try {
        const citations: ZoteroItemData[] = [];
        const citationRegex = /<citation key="[^"]*"[^>]*>[\s\S]*?<\/citation>/g;
        const matches = xmlContent.match(citationRegex);

        if (matches) {
          matches.forEach((match) => {
            try {
              const citation = this.xmlToCitation(match);
              citations.push(citation);
            } catch (parseError) {
              console.warn("Failed to parse citation:", parseError);
            }
          });
        }

        return citations;
      } catch (error) {
        throw new Error(`Failed to parse XML: ${error}`);
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
   * Remove a citation by key
   */
  public async remove(key: string): Promise<boolean> {
    return await PowerPoint.run(async (context) => {
      // Find the XML part
      const customXmlParts = context.presentation.customXmlParts;
      customXmlParts.load("items");
      await context.sync();

      let xmlPart: PowerPoint.CustomXmlPart | null = null;

      for (const part of customXmlParts.items) {
        part.load(["namespaceUri"]);
        // eslint-disable-next-line office-addins/no-context-sync-in-loop
        await context.sync();
        if (part.namespaceUri === CITATION_XML_NAMESPACE) {
          xmlPart = part;
          break;
        }
      }

      if (!xmlPart) {
        return false; // No XML part exists, so citation doesn't exist
      }

      // Get XML content
      const xmlData = xmlPart.getXml();
      await context.sync();
      // eslint-disable-next-line office-addins/load-object-before-read
      let xmlContent = xmlData.value;

      try {
        const citationRegex = new RegExp(`<citation key="${key}"[^>]*>[\\s\\S]*?</citation>`, "g");

        const originalLength = xmlContent.length;
        xmlContent = xmlContent.replace(citationRegex, "");

        if (xmlContent.length < originalLength) {
          // Citation was found and removed
          xmlPart.setXml(xmlContent);
          await context.sync();
          return true;
        } else {
          return false; // Citation not found
        }
      } catch (error) {
        throw new Error(`Failed to process XML: ${error}`);
      }
    });
  }

  /**
   * Helper method to create citation XML string
   */
  private citationToXml(citation: ZoteroItemData): string {
    let xml = `<citation key="${citation.key}">`;
    xml += `<title>${this.escapeXml(citation.title)}</title>`;
    xml += `<authors>`;
    citation.creators.forEach((creator) => {
      const name =
        creator.firstName && creator.lastName
          ? `${creator.firstName} ${creator.lastName}`
          : creator.lastName || creator.firstName || creator.name || "";
      xml += `<author>${this.escapeXml(name)}</author>`;
    });
    xml += `</authors>`;

    // Extract year from date field
    const year = citation.date ? new Date(citation.date).getFullYear() || 0 : 0;
    xml += `<year>${year}</year>`;

    if (citation.publicationTitle) {
      xml += `<journal>${this.escapeXml(citation.publicationTitle)}</journal>`;
    }
    if (citation.DOI) {
      xml += `<doi>${this.escapeXml(citation.DOI)}</doi>`;
    }
    if (citation.url) {
      xml += `<url>${this.escapeXml(citation.url)}</url>`;
    }
    if (citation.abstractNote) {
      xml += `<abstract>${this.escapeXml(citation.abstractNote)}</abstract>`;
    }
    if (citation.tags && citation.tags.length > 0) {
      xml += `<tags>`;
      citation.tags.forEach((tag) => {
        xml += `<tag>${this.escapeXml(tag.tag)}</tag>`;
      });
      xml += `</tags>`;
    }

    xml += `</citation>`;
    return xml;
  }

  /**
   * Helper method to escape XML special characters
   */
  private escapeXml(text: string): string {
    return text
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  /**
   * Parse citation data from XML string
   */
  private xmlToCitation(xmlString: string): ZoteroItemData {
    // Extract key from citation tag
    const keyMatch = xmlString.match(/key="([^"]*)"/);
    const key = keyMatch ? keyMatch[1] : "";

    // Extract title
    const titleMatch = xmlString.match(/<title>([\s\S]*?)<\/title>/);
    const title = titleMatch ? titleMatch[1] : "";

    // Extract year
    const yearMatch = xmlString.match(/<year>(\d+)<\/year>/);
    const year = yearMatch ? parseInt(yearMatch[1]) : 0;

    // Extract authors using a simple approach
    const creators: ZoteroCreator[] = [];
    const authorsMatch = xmlString.match(/<authors>([\s\S]*?)<\/authors>/);
    if (authorsMatch) {
      const authorsContent = authorsMatch[1];
      const authorRegex = /<author>([\s\S]*?)<\/author>/g;
      let authorMatch;
      while ((authorMatch = authorRegex.exec(authorsContent)) !== null) {
        const authorName = authorMatch[1];
        // Simple parsing - try to split first and last name
        const nameParts = authorName.trim().split(" ");
        if (nameParts.length > 1) {
          creators.push({
            creatorType: "author",
            firstName: nameParts.slice(0, -1).join(" "),
            lastName: nameParts[nameParts.length - 1],
          });
        } else {
          creators.push({
            creatorType: "author",
            lastName: authorName,
          });
        }
      }
    }

    const citation: ZoteroItemData = {
      key,
      version: 1,
      itemType: "journal-article", // Default type
      title,
      creators,
      date: year ? `${year}-01-01` : "",
      tags: [],
      collections: [],
      relations: {},
      dateAdded: new Date().toISOString(),
      dateModified: new Date().toISOString(),
    };

    // Extract optional fields
    const journalMatch = xmlString.match(/<journal>([\s\S]*?)<\/journal>/);
    if (journalMatch) citation.publicationTitle = journalMatch[1];

    const doiMatch = xmlString.match(/<doi>([\s\S]*?)<\/doi>/);
    if (doiMatch) citation.DOI = doiMatch[1];

    const urlMatch = xmlString.match(/<url>([\s\S]*?)<\/url>/);
    if (urlMatch) citation.url = urlMatch[1];

    const abstractMatch = xmlString.match(/<abstract>([\s\S]*?)<\/abstract>/);
    if (abstractMatch) citation.abstractNote = abstractMatch[1];

    // Extract tags
    const tagsMatch = xmlString.match(/<tags>([\s\S]*?)<\/tags>/);
    if (tagsMatch) {
      const tags: ZoteroTag[] = [];
      const tagsContent = tagsMatch[1];
      const tagRegex = /<tag>([\s\S]*?)<\/tag>/g;
      let tagMatch;
      while ((tagMatch = tagRegex.exec(tagsContent)) !== null) {
        tags.push({ tag: tagMatch[1] });
      }
      if (tags.length > 0) citation.tags = tags;
    }

    return citation;
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
