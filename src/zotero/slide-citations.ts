import { ZoteroItemData } from "zotero-api-client";

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

function citationToXml(citation: ZoteroItemData): string {
  return `
    <citation xmlns="http://zotero.org/citation">
        <id>${citation.id}</id>
        <author>${citation.creators.map((c) => c.lastName).join(", ")}</author>
        <year>${citation.date?.split("-")[0] ?? ""}</year>
        <text>${citation.title}</text>
        <timestamp>${new Date().toISOString()}</timestamp>
    </citation>
    `;
}

export async function saveSlideCitation(citation: ZoteroItemData) {
  const xmlString = citationToXml(citation);
  try {
    await PowerPoint.run(async (context) => {
      const currentSlide = await getCurrentSlide(context);

      // Add the XML part to the slide
      currentSlide.customXmlParts.add(xmlString);
      await context.sync();

      console.log("Citation XML added to slide successfully");
      console.log("XML content:", xmlString);
    });
  } catch (error) {
    console.error("Error saving citation to slide XML:", error);
    throw error;
  }
}

(window as any).addCitationToSlide = saveSlideCitation;

export async function getSlideCitations(): Promise<ZoteroItemData[]> {
  const citations: ZoteroItemData[] = [];
  try {
    await PowerPoint.run(async (context) => {
      const currentSlide = await getCurrentSlide(context);

      // Load all custom XML parts from the slide
      const customXmlParts = currentSlide.customXmlParts;
      customXmlParts.load("items");
      await context.sync();

      console.log(`Found ${customXmlParts.items.length} XML parts in slide`);

      // Load the XML content for each part
      customXmlParts.items.forEach((xmlPart) => {
        xmlPart.load("xml");
      });
      await context.sync();

      // Parse each XML part to find citations
      for (const xmlPart of customXmlParts.items) {
        try {
          const xmlContent = (xmlPart as any).xml || xmlPart.toString();
          console.log(`Processing XML part: ${xmlContent}`);

          // Parse the XML
          const parser = new DOMParser();
          const xmlDoc = parser.parseFromString(xmlContent, "text/xml");

          // Look for citation elements
          const citationElements = xmlDoc.getElementsByTagName("citation");

          for (let i = 0; i < citationElements.length; i++) {
            const citationEl = citationElements[i];

            // Extract citation data from XML
            const id = citationEl.getElementsByTagName("id")[0]?.textContent || "";
            const author = citationEl.getElementsByTagName("author")[0]?.textContent || "";
            const year = citationEl.getElementsByTagName("year")[0]?.textContent || "";
            const title = citationEl.getElementsByTagName("text")[0]?.textContent || "";
            const timestamp = citationEl.getElementsByTagName("timestamp")[0]?.textContent || "";

            // Create a ZoteroItemData object from the XML data
            const citation: ZoteroItemData = {
              key: id,
              version: 1,
              itemType: "book", // Default type
              title: title,
              creators: author ? [{ creatorType: "author", lastName: author }] : [],
              date: year ? `${year}-01-01` : "",
              collections: [],
              relations: { foo: [] },
              dateAdded: timestamp || new Date().toISOString(),
              dateModified: timestamp || new Date().toISOString(),
              tags: [],
            };

            citations.push(citation);
            console.log("Found citation in XML:", citation);
          }
        } catch (parseError) {
          console.log(`Failed to parse XML part: ${parseError}`);
        }
      }

      console.log(`Loaded ${citations.length} citations from slide XML`);
    });
  } catch (error) {
    console.error("Error retrieving citations from slide XML:", error);
    throw error;
  }
  return citations;
}

export async function removeSlideCitation(citationId: string): Promise<boolean> {
  try {
    let found = false;
    await PowerPoint.run(async (context) => {
      const currentSlide = await getCurrentSlide(context);

      // Load all custom XML parts from the slide
      const customXmlParts = currentSlide.customXmlParts;
      customXmlParts.load("items");
      await context.sync();

      // Load the XML content for each part
      customXmlParts.items.forEach((xmlPart) => {
        xmlPart.load("xml");
      });
      await context.sync();

      // Find and remove the XML part containing the citation
      for (const xmlPart of customXmlParts.items) {
        try {
          const xmlContent = (xmlPart as any).xml || xmlPart.toString();

          // Parse the XML to check if it contains our citation
          const parser = new DOMParser();
          const xmlDoc = parser.parseFromString(xmlContent, "text/xml");
          const citationElements = xmlDoc.getElementsByTagName("citation");

          for (let i = 0; i < citationElements.length; i++) {
            const citationEl = citationElements[i];
            const idElement = citationEl.getElementsByTagName("id")[0];
            if (idElement && idElement.textContent === citationId) {
              xmlPart.delete();
              found = true;
              break;
            }
          }

          if (found) break;
        } catch (parseError) {
          console.log(`Failed to parse XML part for removal: ${parseError}`);
        }
      }

      await context.sync();
      console.log(`Citation removal ${found ? "successful" : "failed"} for ID: ${citationId}`);
    });

    return found;
  } catch (error) {
    console.error("Error removing citation from slide XML:", error);
    throw error;
  }
}

// Alternative approach: Store citations in slide tags instead of custom XML
export async function saveSlideCitationAsTag(citation: ZoteroItemData): Promise<void> {
  try {
    await PowerPoint.run(async (context) => {
      const currentSlide = await getCurrentSlide(context);

      // Create a unique tag key for this citation
      const citationKey = `zotero-citation-${citation.key || Date.now()}`;

      // Serialize citation data as JSON string
      const citationData = JSON.stringify({
        id: citation.key,
        title: citation.title,
        creators: citation.creators,
        date: citation.date,
        itemType: citation.itemType,
        timestamp: new Date().toISOString(),
      });

      // Store citation as slide tag
      currentSlide.tags.add(citationKey, citationData);
      await context.sync();

      console.log("Citation saved as slide tag successfully");
      console.log("Citation key:", citationKey);
      console.log("Citation data:", citationData);
    });
  } catch (error) {
    console.error("Error saving citation as slide tag:", error);
    throw error;
  }
}

export async function getSlideCitationsFromTags(): Promise<ZoteroItemData[]> {
  const citations: ZoteroItemData[] = [];
  try {
    await PowerPoint.run(async (context) => {
      const currentSlide = await getCurrentSlide(context);

      // Load all tags from the slide
      const tags = currentSlide.tags;
      tags.load("items");
      await context.sync();

      console.log(`Found ${tags.items.length} tags in slide`);

      // Load tag details
      tags.items.forEach((tag) => {
        tag.load("key");
        tag.load("value");
      });
      await context.sync();

      // Find citation tags and parse their data
      tags.items.forEach((tag) => {
        if (tag.key && tag.key.startsWith("zotero-citation-")) {
          try {
            const citationData = JSON.parse(tag.value);

            // Convert back to ZoteroItemData format
            const citation: ZoteroItemData = {
              key: citationData.id,
              version: 1,
              itemType: citationData.itemType || "book",
              title: citationData.title,
              creators: citationData.creators || [],
              date: citationData.date || "",
              collections: [],
              relations: { foo: [] },
              dateAdded: citationData.timestamp || new Date().toISOString(),
              dateModified: citationData.timestamp || new Date().toISOString(),
              tags: [],
            };

            citations.push(citation);
            console.log("Found citation in tag:", citation);
          } catch (parseError) {
            console.log(`Failed to parse citation tag: ${tag.key}`, parseError);
          }
        }
      });

      console.log(`Loaded ${citations.length} citations from slide tags`);
    });
  } catch (error) {
    console.error("Error retrieving citations from slide tags:", error);
    throw error;
  }
  return citations;
}

export async function removeSlideCitationFromTags(citationId: string): Promise<boolean> {
  try {
    let found = false;
    await PowerPoint.run(async (context) => {
      const currentSlide = await getCurrentSlide(context);

      // Load all tags from the slide
      const tags = currentSlide.tags;
      tags.load("items");
      await context.sync();

      // Load tag details
      tags.items.forEach((tag) => {
        tag.load("key");
        tag.load("value");
      });
      await context.sync();

      // Collect all citation tags except the one to remove
      const citationTagsToKeep: { key: string; value: string }[] = [];

      for (const tag of tags.items) {
        if (tag.key && tag.key.startsWith("zotero-citation-")) {
          try {
            const citationData = JSON.parse(tag.value);
            if (citationData.id === citationId) {
              // Skip this tag (effectively removing it)
              found = true;
              console.log(`Found citation to remove: ${citationId}`);
            } else {
              // Keep this citation tag
              citationTagsToKeep.push({ key: tag.key, value: tag.value });
            }
          } catch (parseError) {
            console.log(`Failed to parse citation tag for removal: ${tag.key}`, parseError);
            // Keep unparseable citation tags to be safe
            citationTagsToKeep.push({ key: tag.key, value: tag.value });
          }
        }
      }

      if (found) {
        console.log(
          `Removing citation and re-adding ${citationTagsToKeep.length} remaining citation tags`
        );

        // Since we can't delete individual tags, we'll clear all citation tags
        // and re-add only the ones we want to keep
        // First, let's try to add new tags with different keys to override the old ones

        // Add all the citation tags we want to keep with temporary new keys
        const tempTags: string[] = [];
        for (let i = 0; i < citationTagsToKeep.length; i++) {
          const tempKey = `temp-citation-${i}-${Date.now()}`;
          try {
            currentSlide.tags.add(tempKey, citationTagsToKeep[i].value);
            tempTags.push(tempKey);
          } catch (addError) {
            console.log(`Failed to add temporary tag ${tempKey}:`, addError);
          }
        }

        await context.sync();

        // Now add the tags back with their original keys
        for (let i = 0; i < citationTagsToKeep.length; i++) {
          try {
            const originalTag = citationTagsToKeep[i];
            currentSlide.tags.add(originalTag.key, originalTag.value);
          } catch (addError) {
            console.log(`Failed to re-add tag ${citationTagsToKeep[i].key}:`, addError);
          }
        }

        await context.sync();
      }

      console.log(`Citation removal ${found ? "successful" : "failed"} for ID: ${citationId}`);
    });

    return found;
  } catch (error) {
    console.error("Error removing citation from slide tags:", error);
    throw error;
  }
}

export async function debugSlideTags(): Promise<void> {
  try {
    await PowerPoint.run(async (context) => {
      console.log("=== DEBUG: Slide Tags ===");

      const currentSlide = await getCurrentSlide(context);
      console.log("Using current slide for debugging");

      // Load all tags from the slide
      const tags = currentSlide.tags;
      tags.load("items");
      await context.sync();

      console.log(`Found ${tags.items.length} tags in slide`);

      if (tags.items.length === 0) {
        console.log("No tags found in slide");
        return;
      }

      // Load tag details
      tags.items.forEach((tag, index) => {
        tag.load("key");
        tag.load("value");
        console.log(`Loading tag ${index + 1}`);
      });
      await context.sync();

      // Display each tag
      tags.items.forEach((tag, index) => {
        console.log(`Tag ${index + 1}:`);
        console.log(`  Key: ${tag.key}`);
        console.log(`  Value: ${tag.value}`);

        // If it's a citation tag, try to parse it
        if (tag.key && tag.key.startsWith("zotero-citation-")) {
          try {
            const citationData = JSON.parse(tag.value);
            console.log(`  Parsed Citation:`, citationData);
          } catch (parseError) {
            console.log(`  Parse Error: ${parseError}`);
          }
        }
      });
    });
  } catch (error) {
    console.error("Error debugging slide tags:", error);
  }
}

// Export debug function to global scope
(window as any).debugSlideTags = debugSlideTags;
