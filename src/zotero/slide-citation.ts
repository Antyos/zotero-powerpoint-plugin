import { ZoteroItemData } from "zotero-api-client";
import { CITATION_TAG_KEY, CitationStore } from "./citation-store";
import { getJournalAbbreviation } from "./journal-abbreviations";

export async function getCurrentSlide(
  context?: PowerPoint.RequestContext
): Promise<PowerPoint.Slide> {
  if (!context) {
    return await PowerPoint.run(async (ctx) => {
      return await getCurrentSlide(ctx);
    });
  }
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
 * Get or create a text box for citations on the slide
 */
async function getCitationTextBox(
  slide: PowerPoint.Slide,
  citationBoxName: string = "Citations"
): Promise<PowerPoint.Shape> {
  const shapes = slide.shapes;
  shapes.load("items");
  await slide.context.sync();

  for (const shape of shapes.items) {
    shape.load("name");
  }
  await slide.context.sync();
  let citationBox = shapes.items.find((shape) => shape.name === citationBoxName);
  if (citationBox) {
    return citationBox;
  }
  // Check the slide layout master for a hidden citation box
  const slideLayout = slide.layout;
  slideLayout.load("shapes");
  await slide.context.sync();
  for (const shape of slideLayout.shapes.items) {
    shape.load("name");
  }
  await slide.context.sync();
  const layoutShape = slideLayout.shapes.items.find((shape) => shape.name === citationBoxName);
  if (layoutShape) {
    // Reveal the shape on the slide by copying it
    const copiedShape = slide.shapes.addTextBox("", {
      left: layoutShape.left,
      top: layoutShape.top,
      width: layoutShape.width,
      height: layoutShape.height,
    });
    copiedShape.name = citationBoxName;
    const layoutTextFrame = layoutShape.textFrame;
    const layoutFont = layoutShape.textFrame.textRange.font;
    layoutTextFrame.load([
      "autoSizeSetting",
      "verticalAlignment",
      "topMargin",
      "bottomMargin",
      "leftMargin",
      "rightMargin",
      "wordWrap",
    ]);
    layoutFont.load(["name", "size", "color", "allCaps", "smallCaps"]);
    await slide.context.sync();
    copiedShape.textFrame.autoSizeSetting = layoutTextFrame.autoSizeSetting;
    copiedShape.textFrame.verticalAlignment = layoutTextFrame.verticalAlignment;
    copiedShape.textFrame.topMargin = layoutTextFrame.topMargin;
    copiedShape.textFrame.bottomMargin = layoutTextFrame.bottomMargin;
    copiedShape.textFrame.leftMargin = layoutTextFrame.leftMargin;
    copiedShape.textFrame.rightMargin = layoutTextFrame.rightMargin;
    copiedShape.textFrame.wordWrap = layoutTextFrame.wordWrap;
    // Don't need to copy bold/italic since they will be applied to substrings
    copiedShape.textFrame.textRange.font.name = layoutFont.name;
    copiedShape.textFrame.textRange.font.size = layoutFont.size;
    copiedShape.textFrame.textRange.font.color = layoutFont.color;
    copiedShape.textFrame.textRange.font.allCaps = layoutFont.allCaps;
    copiedShape.textFrame.textRange.font.smallCaps = layoutFont.smallCaps;
    // Not copying over paragraph-level formatting for now because I don't feel like it.
    // In theory, there should also be a format painter api at some point to do that.
    await slide.context.sync();
    return copiedShape;
  }
  // Create a new text box if not found
  // There is no reliable API to get slide dimensions, so we are hard-coding it for now.
  const slideWidth = 960;
  const slideHeight = 540;
  const boxHeight = 30;
  const newBox = shapes.addTextBox("", {
    left: 0,
    top: slideHeight - boxHeight,
    width: slideWidth,
    height: boxHeight,
  });
  newBox.name = citationBoxName;
  newBox.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.bottom;
  return newBox;
}

class CitationFormatter {
  private _format: string;
  private _delimeter: string;

  constructor(format: string, delimeter: string) {
    // Initialization if needed
    this._format = format;
    this._delimeter = delimeter;
  }

  private async getJournalAbbreviation(citation: ZoteroItemData): Promise<string | null> {
    if (!citation.publicationTitle) {
      return null;
    }
    return (await getJournalAbbreviation(citation.publicationTitle)) ?? citation.publicationTitle;
  }

  async format(citation: ZoteroItemData): Promise<FormattedText[]> {
    const year =
      citation.date && typeof citation.date === "string"
        ? citation.date.split("-")[0]
        : citation.date || "n.d.";
    const creator = citation.creators[0];
    const etal = citation.creators && citation.creators.length > 1 ? " et al." : "";
    const journalAbbreviation = (await this.getJournalAbbreviation(citation)) ?? "";

    // Replace placeholders with actual values
    let text = this._format
      .replace("{creator.lastName}", creator?.lastName || "Unknown")
      .replace("{creator.firstName}", creator?.firstName || "Unknown")
      .replace("{creator.name}", creator?.name || "Unknown")
      .replace("{title}", citation.title || "No Title")
      .replace("{key}", citation.key)
      .replace("{itemType}", citation.itemType || "Unknown Type")
      .replace("{abstractNote}", citation.abstractNote || "No Abstract")
      .replace("{publicationTitle}", citation.publicationTitle || "No Publication")
      .replace("{volume}", citation.volume || "No Volume")
      .replace("{issue}", citation.issue || "No Issue")
      .replace("{pages}", citation.pages || "No Pages")
      .replace("{publisher}", citation.publisher || "No Publisher")
      .replace("{DOI}", citation.DOI || "No DOI")
      .replace("{ISBN}", citation.ISBN || "No ISBN")
      .replace("{URL}", citation.url || "No URL")
      .replace("{accessDate}", citation.accessDate || "No Access Date")
      .replace("{archive}", citation.archive || "No Archive")
      .replace("{archiveLocation}", citation.archiveLocation || "No Archive Location")
      .replace("{libraryCatalog}", citation.libraryCatalog || "No Library Catalog")
      .replace("{callNumber}", citation.callNumber || "No Call Number")
      .replace("{rights}", citation.rights || "No Rights Info")
      .replace("{date}", citation.date || "No Date")
      .replace("{extra}", citation.extra || "No Extra Info")
      .replace("{series}", citation.series || "No Series")
      .replace("{seriesNumber}", citation.seriesNumber || "No Series Number")
      .replace("{institution}", citation.institution || "No Institution")
      .replace("{department}", citation.department || "No Department")
      .replace("{etal}", etal)
      .replace("{year}", year)
      .replace("{journalAbbreviation}", journalAbbreviation);

    // Parse formatting tags and create formatted text segments
    return CitationFormatter.parseFormattedText(text);
  }

  private static parseFormattedText(text: string): FormattedText[] {
    const segments: FormattedText[] = [];
    let currentIndex = 0;

    // Regular expression to find <b>, <i>, and </b>, </i> tags
    const formatRegex = /<(\/?)([bi])>/g;
    let match;
    let isBold = false;
    let isItalic = false;

    while ((match = formatRegex.exec(text)) !== null) {
      // Add text before the tag
      if (match.index > currentIndex) {
        const textContent = text.substring(currentIndex, match.index);
        if (textContent) {
          segments.push({
            text: textContent,
            bold: isBold,
            italic: isItalic,
          });
        }
      }

      // Update formatting state
      const isClosing = match[1] === "/";
      const tagType = match[2];

      if (tagType === "b") {
        isBold = !isClosing;
      } else if (tagType === "i") {
        isItalic = !isClosing;
      }

      currentIndex = match.index + match[0].length;
    }

    // Add remaining text
    if (currentIndex < text.length) {
      const remainingText = text.substring(currentIndex);
      if (remainingText) {
        segments.push({
          text: remainingText,
          bold: isBold,
          italic: isItalic,
        });
      }
    }

    return segments;
  }

  public get delimeter(): string {
    return this._delimeter;
  }
}

interface FormattedText {
  text: string;
  bold: boolean;
  italic: boolean;
}

export async function showCitationsOnSlide(
  slide?: PowerPoint.Slide,
  format: string = "<b>{creator.lastName}</b>{etal}, {year}, <i>{journalAbbreviation}</i>",
  delimeter: string = ";  "
): Promise<void> {
  const citationFormatter = new CitationFormatter(format, delimeter);
  return await PowerPoint.run(async (context) => {
    slide = slide ?? (await getCurrentSlide(context));
    const citationKeys = await getCitationKeysOnSlide(slide);
    const citations = await CitationStore.getInstance().getItem(citationKeys);
    const citationBox = await getCitationTextBox(slide);

    if (citationBox && citations.length === 0) {
      // The API seems to be missing a hide() method, so we'll just delete it and recreate it later
      // if needed.
      citationBox.delete();
      await context.sync();
      console.log("No citations on current slide.");
      return;
    }

    // Build the complete text first, then apply formatting
    let completeText = "";
    const allSegments: Array<{ segment: FormattedText; startIndex: number; endIndex: number }> = [];

    for (const [index, citation] of citations.entries()) {
      if (!citation) {
        continue;
      }
      const formattedSegments = await citationFormatter.format(citation);

      for (const segment of formattedSegments) {
        const startIndex = completeText.length;
        completeText += segment.text;
        const endIndex = completeText.length;

        allSegments.push({
          segment,
          startIndex,
          endIndex,
        });
      }

      if (index < citations.length - 1) {
        allSegments.push({
          segment: { text: citationFormatter.delimeter, bold: false, italic: false },
          startIndex: completeText.length,
          endIndex: completeText.length + citationFormatter.delimeter.length,
        });
        completeText += citationFormatter.delimeter;
      }
    }

    // Set the complete text
    citationBox.textFrame.textRange.text = completeText;

    // Apply formatting to each segment
    for (const { segment, startIndex, endIndex } of allSegments) {
      const range = citationBox.textFrame.textRange.getSubstring(startIndex, endIndex - startIndex);
      range.font.bold = segment.bold;
      range.font.italic = segment.italic;
    }

    await context.sync();
    console.log(`Citations on current slide: ${citationKeys.join(", ")}`);
  });
}

/**
 * ===== Slide-level citation key management functions =====
 */

/**
 * Add a citation key to the current slide
 */
export async function addCitationToSlide(
  citationKey: string,
  slide?: PowerPoint.Slide
): Promise<void> {
  slide = slide ?? (await getCurrentSlide());
  const context = slide.context;
  const citationKeys = await getCitationKeysOnSlide(slide);
  if (!citationKeys.includes(citationKey)) {
    citationKeys.push(citationKey);
    slide.tags.add(CITATION_TAG_KEY, citationKeys.join(","));
  }
  context.sync();
}

/**
 * Remove a citation key from the current slide
 */
export async function removeCitationFromSlide(
  citationKey: string,
  slide?: PowerPoint.Slide
): Promise<boolean> {
  try {
    slide = slide ?? (await getCurrentSlide());
    const context = slide.context;
    const citationKeys = await getCitationKeysOnSlide(slide);
    if (!citationKeys.includes(citationKey)) {
      return false;
    }
    citationKeys.splice(citationKeys.indexOf(citationKey), 1);
    slide.tags.add(CITATION_TAG_KEY, citationKeys.join(","));
    context.sync();
    CitationStore.getInstance().prune();
    return true;
  } catch (error) {
    console.error("Failed to remove citation from slide:", error);
    return false;
  }
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
export async function getCitationKeysOnSlide(slide?: PowerPoint.Slide): Promise<string[]> {
  if (!slide) {
    slide = await getCurrentSlide();
  }
  // Having a separate variable for context makes TypeScript happy
  const context = slide.context;
  const tags = slide.tags;
  await context.sync();
  // console.log(`Slide Tags: ${JSON.stringify(tags, null, 2)}`);
  const citationTag = tags.getItemOrNullObject(CITATION_TAG_KEY);
  // const citationTag = tags.getItem("foo");
  citationTag.load("value");
  await context.sync();

  console.log("Citation Tag:", citationTag);
  if (citationTag.isNullObject || !citationTag.value) {
    return []; // No citations on this slide
  }

  const tagValue = citationTag.value;
  // Split the tag value by commas to get individual citation keys
  return tagValue.split(",");
}

/**
 * Update the order of citation keys on the current slide
 */
export async function updateCitationKeysOrder(
  orderedKeys: string[],
  slide?: PowerPoint.Slide
): Promise<void> {
  if (!slide) {
    slide = await getCurrentSlide();
  }

  const context = slide.context;
  const tags = slide.tags;
  await context.sync();

  // Update the citation tag with the new order
  tags.add(CITATION_TAG_KEY, orderedKeys.join(","));
  await context.sync();

  console.log(`Updated citation order on slide: ${orderedKeys.join(", ")}`);
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
  const citationKeys = await getCitationKeysOnSlide();

  const citations: ZoteroItemData[] = [];
  for (const key of citationKeys) {
    const citation = await store.getItem(key);
    if (citation) {
      citations.push(citation);
    }
  }

  return citations;
}
