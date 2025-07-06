/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { ZoteroItemData } from "zotero-api-client";
import { ZoteroLibrary, TitleCreatorDate } from "../zotero/zotero-connector";
import {
  CitationStore,
  debugSlideTags,
  getCitationsOnSlide,
  insertCitationOnSlide,
  removeCitationFromSlide,
} from "../zotero/citation-store";

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    const sideloadMsg = document.getElementById("sideload-msg");
    if (sideloadMsg) {
      sideloadMsg.style.display = "none";
    }
    const appBody = document.getElementById("app-body");
    if (appBody) {
      appBody.classList.remove("app-body-hidden");
      appBody.style.display = "flex";
    }

    // Initialize Zotero integration
    initializeZoteroUI();
  }
});

function initializeZoteroUI() {
  const zotero = ZoteroLibrary.getInstance();

  // Load existing configuration
  zotero.loadConfig();

  // Set up event listeners
  const configureButton = document.getElementById("configure-zotero");
  const searchButton = document.getElementById("search-zotero");
  const testButton = document.getElementById("test-connection");
  const refreshCitationsButton = document.getElementById("refresh-citations");
  const insertMockCitationButton = document.getElementById("insert-mock-citation");
  const searchInput = document.getElementById("search-query");
  const debugSlideTagsButton = document.getElementById("debug-slide-tags");
  const debugCitationStoreButton = document.getElementById("debug-citation-store");
  const debugCitationsButton = document.getElementById("debug-citations");
  const clearCitationStoreButton = document.getElementById("clear-citation-store");

  if (configureButton) {
    configureButton.onclick = configureZotero;
  }

  if (searchButton) {
    searchButton.onclick = searchZoteroLibrary;
  }

  if (testButton) {
    testButton.onclick = testZoteroConnection;
  }

  if (refreshCitationsButton) {
    refreshCitationsButton.onclick = updateCitationsPanel;
  }

  if (debugSlideTagsButton) {
    debugSlideTagsButton.onclick = () => {
      debugSlideTags().catch((error) => {
        console.error("Error debugging slide tags:", error);
      });
    };
  }

  if (debugCitationsButton) {
    debugCitationsButton.onclick = async () => {
      try {
        console.log("Debugging citations...");
        await CitationStore.getInstance().debugCitations();
      } catch (error) {
        console.error("Error debugging citations:", error);
      }
    };
  }

  if (debugCitationStoreButton) {
    debugCitationStoreButton.onclick = () => {
      console.log("Debugging citation store...");
      CitationStore.getInstance().debugXml();
    };
  }

  if (insertMockCitationButton) {
    insertMockCitationButton.onclick = () => {
      insertCitation({
        key: "123456",
        title: "Test Citation",
        creators: [{ creatorType: "author", lastName: "Smith" }],
        date: "2023-04-19",
        itemType: "book",
        collections: [],
        dateAdded: new Date().toISOString(),
        dateModified: new Date().toISOString(),
        relations: { foo: [] },
        version: 1,
        tags: [],
      });
    };
  }

  if (clearCitationStoreButton) {
    clearCitationStoreButton.onclick = async () => {
      try {
        console.log("Clearing citation store...");
        await CitationStore.getInstance().clearStore();
        console.log("Citation store cleared.");
        // Refresh the current citations list
        setTimeout(updateCitationsPanel, 500);
      } catch (error) {
        console.error("Error clearing citation store:", error);
      }
    };
  }

  if (searchInput) {
    searchInput.addEventListener("keypress", (e) => {
      if (e.key === "Enter") {
        searchZoteroLibrary();
      }
    });
  }

  // Load current citations on initialization
  updateCitationsPanel();
}

async function configureZotero() {
  try {
    console.log("Opening configuration dialog...");
    const zotero = ZoteroLibrary.getInstance();
    const result = await zotero.configureFromDialog();

    if (result) {
      console.log("Configuration saved successfully!");
    } else {
      console.log("Configuration cancelled.");
    }
  } catch (error) {
    console.error("Configuration error:", error);
  }
}

async function searchZoteroLibrary() {
  try {
    const searchInput = document.getElementById("search-query") as HTMLInputElement;
    const query = searchInput?.value?.trim();

    if (!query) {
      return;
    }

    console.log("Searching Zotero library...");
    const zotero = ZoteroLibrary.getInstance();

    if (!zotero.isConfigured()) {
      console.log("Please configure Zotero API settings first.");
      return;
    }

    const results = await zotero.quickSearch(query);
    displaySearchResults(results);
    console.log(`Found ${results.length} items.`);
  } catch (error) {
    console.error("Search error:", error);
  }
}

async function testZoteroConnection() {
  try {
    const zotero = ZoteroLibrary.getInstance();
    if (!zotero.isConfigured()) {
      return;
    }

    const isConnected = await zotero.checkConnection();
    console.log(isConnected ? "Connection successful!" : "Connection failed.");
  } catch (error) {
    console.error("Connection test error:", error);
  }
}

function displaySearchResults(results: TitleCreatorDate[]) {
  const resultsContainer = document.getElementById("search-results");
  if (!resultsContainer) return;

  if (results.length === 0) {
    resultsContainer.innerHTML = '<p class="ms-font-s">No results found.</p>';
    return;
  }

  const resultsList = results
    .map((item) => {
      const title = item.title || "Untitled";
      const author =
        item.creators.length > 1
          ? `${item.creators[0].lastName} et al.`
          : item.creators.length == 1
            ? item.creators[0].lastName
            : "unknown";
      return `
        <div class="ms-ListItem zotero-result-item"
             onclick="insertCitation('${item.id || "unknown"}', '${author}', '(${item.date ?? ""})')">
          <div class="ms-font-m zotero-result-title">${title}</div>
          <div class="ms-font-s zotero-result-meta">${author} (${item.date.split("-")[0]})</div>
        </div>
      `;
    })
    .join("");

  resultsContainer.innerHTML = resultsList;
}

// Global function for citation insertion (called from HTML onclick)
async function insertCitation(citation: ZoteroItemData) {
  try {
    console.log(`Inserting citation: ${citation.creators[0].lastName}, ${citation.date}`);
    // Use simplified storage approach
    await insertCitationOnSlide(citation);
  } catch (error) {
    console.error("Citation insertion error:", error);
  }
  // Refresh the current citations list
  setTimeout(updateCitationsPanel, 500);
}
(window as any).insertCitation = insertCitation;

async function updateCitationsPanel() {
  try {
    console.log("Loading current citations from slide...");
    await PowerPoint.run(async (_context) => {
      const citations = await getCitationsOnSlide();
      console.log(`Found ${citations.length} citations in current slide.`);
      console.log(citations);
      displayCitationsOnTaskpane(citations);
    });
  } catch (error) {
    console.error("Error loading citations:", error);
    displayCitationsOnTaskpane([]);
  }
}

function displayCitationsOnTaskpane(citations: ZoteroItemData[]) {
  const citationsContainer = document.getElementById("current-citations");
  if (!citationsContainer) return;

  if (citations.length === 0) {
    citationsContainer.innerHTML = '<p class="ms-font-s">No citations found in current slide.</p>';
    return;
  }

  const citationsList = citations
    .map((citation) => {
      return `
        <div class="ms-ListItem zotero-result-item citation-item"
             data-citation-id="${citation.key}">
          <div class="ms-font-m zotero-result-title">${citation.title}</div>
          <div class="ms-font-s zotero-result-meta">
            ID: ${citation.key} | Author: ${citation.creators.map((c) => c.lastName).join(", ")} | Year: ${citation.date?.split("-")[0] ?? ""}
          </div>
          <div class="citation-actions">
            <button class="ms-Button ms-Button--small" onclick="removeCitation('${citation.key}')">
              <span class="ms-Button-label">Remove</span>
            </button>
          </div>
        </div>
      `;
    })
    .join("");

  citationsContainer.innerHTML = citationsList;
}

// Global function for citation removal (called from HTML onclick)
async function removeCitation(citationId: string) {
  try {
    console.log(`Removing citation: ${citationId}`);

    const success = await removeCitationFromSlide(citationId);

    if (success) {
      console.log("Citation removed successfully");
      // Refresh the current citations list
      setTimeout(updateCitationsPanel, 500);
    } else {
      console.warn("Citation not found for removal");
    }
  } catch (error) {
    console.error("Citation removal error:", error);
  }
}
(window as any).removeCitation = removeCitation;

export async function run() {
  /**
   * Insert your PowerPoint code here
   */
  const options: Office.SetSelectedDataOptions = { coercionType: Office.CoercionType.Text };

  await Office.context.document.setSelectedDataAsync(" ", options);
  await Office.context.document.setSelectedDataAsync("Hello World!", options);
}
