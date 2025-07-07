export interface JMedline {
  JrId: number;
  JournalTitle: string;
  MedAbbr: string;
  "ISSN (Print)"?: string;
  "ISSN (Online)"?: string;
  ISOAbbr: string;
  NlmId: string;
}

const J_MEDLINE_URL = "https://ftp.ncbi.nih.gov/pubmed/J_Medline.txt";

// Cache for storing processed data
let cachedJMedlineData: JMedline[] | null = null;
let cachedMedAbbrMap: Map<string, string> | null = null;
let cachedISOAbbrMap: Map<string, string> | null = null;
let cacheTimestamp: number | null = null;
const CACHE_DURATION = 24 * 60 * 60 * 1000; // 24 hours in milliseconds

/**
 * Check if cached data is still valid
 */
function isCacheValid(): boolean {
  if (!cacheTimestamp) return false;
  const now = Date.now();
  return now - cacheTimestamp < CACHE_DURATION;
}

/**
 * Fetch and parse the J_Medline.txt data
 */
async function fetchAndParseJMedlineData(): Promise<JMedline[]> {
  console.log("Fetching and parsing journal abbreviations data from NCBI");
  try {
    const response = await fetch(J_MEDLINE_URL);
    if (!response.ok) {
      throw new Error(`Failed to fetch journal data: ${response.status} ${response.statusText}`);
    }

    const data = await response.text();
    const lines = data.split("\n");
    const abbreviations: JMedline[] = [];
    let currentEntry: Partial<JMedline> = {};

    for (const line of lines) {
      if (line.trim() === "" || line.startsWith("-")) {
        if (Object.keys(currentEntry).length > 0) {
          abbreviations.push(currentEntry as JMedline);
          currentEntry = {};
        }
        continue;
      }
      const [key, value] = line.split(":", 2);
      switch (key) {
        case "JrId":
          currentEntry.JrId = parseInt(value.trim(), 10);
          break;
        case "JournalTitle":
          currentEntry.JournalTitle = value.trim();
          break;
        case "MedAbbr":
          currentEntry.MedAbbr = value;
          break;
        case "ISSN (Print)":
          currentEntry["ISSN (Print)"] = value;
          break;
        case "ISSN (Online)":
          currentEntry["ISSN (Online)"] = value;
          break;
        case "ISOAbbr":
          currentEntry.ISOAbbr = value;
          break;
        case "NlmId":
          currentEntry.NlmId = value;
          break;
      }
    }

    return abbreviations;
  } catch (error) {
    console.error("Error fetching journal abbreviations:", error);
    throw error;
  }
}

/**
 * Clear the cache (useful for testing or forcing a refresh)
 */
export function clearJournalAbbreviationCache(): void {
  cachedJMedlineData = null;
  cachedMedAbbrMap = null;
  cachedISOAbbrMap = null;
  cacheTimestamp = null;
  console.log("Journal abbreviation cache cleared");
}

export async function getPubMedAbbreviations(): Promise<JMedline[]> {
  // Check if we have cached data and it's still valid
  if (cachedJMedlineData && isCacheValid()) {
    console.log("Using cached JMedline data");
    return cachedJMedlineData;
  }

  // Fetch and parse fresh data
  const abbreviations = await fetchAndParseJMedlineData();

  // Update cache
  cachedJMedlineData = abbreviations;
  cacheTimestamp = Date.now();

  return abbreviations;
}

export async function getJournalAbbreviationMap(
  mode: "MedAbbr" | "ISOAbbr"
): Promise<Map<string, string>> {
  // Check if we have the specific cached map and it's still valid
  const cachedMap = mode === "MedAbbr" ? cachedMedAbbrMap : cachedISOAbbrMap;
  if (cachedMap && isCacheValid()) {
    console.log(`Using cached ${mode} abbreviations map`);
    return cachedMap;
  }

  // Get the full dataset (this will use cache if available)
  const jMedlineData = await getPubMedAbbreviations();

  // Build the specific abbreviation map
  const abbreviations = new Map<string, string>();
  for (const entry of jMedlineData) {
    if (entry.JournalTitle && entry[mode]) {
      abbreviations.set(entry.JournalTitle.toLowerCase(), entry[mode]);
    }
  }

  // Cache the specific map
  if (mode === "MedAbbr") {
    cachedMedAbbrMap = abbreviations;
  } else {
    cachedISOAbbrMap = abbreviations;
  }

  return abbreviations;
}

export async function getJournalAbbreviation(
  journalTitle?: string,
  mode: "MedAbbr" | "ISOAbbr" = "MedAbbr"
): Promise<string | null> {
  if (!journalTitle) {
    return null;
  }
  const abbreviations = await getJournalAbbreviationMap(mode);
  return abbreviations.get(journalTitle.toLowerCase()) || null;
}
