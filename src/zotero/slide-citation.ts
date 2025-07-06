import { ZoteroItemData } from "zotero-api-client";
import { CITATION_TAG_KEY, CitationStore } from "./citation-store";

export async function getCurrentSlide(
  context: PowerPoint.RequestContext
): Promise<PowerPoint.Slide> {
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
