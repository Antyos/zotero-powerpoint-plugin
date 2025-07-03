import { ZoteroItemData } from "zotero-api-client";

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
      const slides = context.presentation.getSelectedSlides();
      const currentSlide: PowerPoint.Slide = slides.getItemAt(0);
      const customXmlParts = currentSlide.customXmlParts.load();
      await context.sync();
      currentSlide.customXmlParts.add(xmlString);
      currentSlide.customXmlParts.load("toJSON");
      await context.sync();
      console.log(currentSlide.customXmlParts.toJSON());
    });
  } catch (error) {
    console.error("Error saving citation to slide XML:", error);
    throw error;
  }
}

(window as any).addCitationToSlide = saveSlideCitation;

export async function getSlideCitations(): Promise<ZoteroItemData[]> {
  const jsonData: ZoteroItemData[] = [];
  try {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.getSelectedSlides();
      slides.load("items");
      await context.sync();
      slides.items[0].customXmlParts.load("id");
      const xmlString = slides.getItemAt(0).customXmlParts.getItemOrNullObject("citation");
      const xmlData = xmlString.toJSON();
      jsonData.push(xmlData as any);
      await context.sync();

      console.log(`Loading citations from XML part with ID: ${jsonData}`);
    });
  } catch (error) {
    console.error("Error retrieving citations from slide XML:", error);
    throw error;
  }
  return jsonData;
}
