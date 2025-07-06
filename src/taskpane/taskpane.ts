/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { ZoteroItemData } from "zotero-api-client";
import { ZoteroLibrary } from "../zotero/zotero-connector";
import { CitationStore } from "../zotero/citation-store";
import {
  getCitationsOnSlide,
  insertCitationOnSlide,
  debugSlideTags,
  removeCitationFromSlide,
  showCitationsOnSlide,
} from "../zotero/slide-citation";

const SEARCH_DEBOUNCE_MS = 300;

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
        key: "1234ABCD",
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
    // Debounced search as user types
    let searchTimeout: ReturnType<typeof setTimeout>;
    let selectedIndex = -1;
    let searchResults: ZoteroItemData[] = [];

    searchInput.addEventListener("input", () => {
      clearTimeout(searchTimeout);
      searchTimeout = setTimeout(() => {
        const query = (searchInput as HTMLInputElement).value.trim();
        if (query.length > 0) {
          searchZoteroLibrary();
        } else {
          // Clear results when input is empty
          hideSearchDropdown();
        }
        selectedIndex = -1; // Reset selection on new search
      }, SEARCH_DEBOUNCE_MS);
    });

    // Handle keyboard navigation
    searchInput.addEventListener("keydown", (e) => {
      const dropdown = document.getElementById("search-dropdown");
      if (!dropdown || dropdown.classList.contains("hidden")) {
        return;
      }

      switch (e.key) {
        case "ArrowDown":
          e.preventDefault();
          selectedIndex = Math.min(selectedIndex + 1, searchResults.length - 1);
          updateSearchDropdownSelection();
          break;
        case "ArrowUp":
          e.preventDefault();
          selectedIndex = Math.max(selectedIndex - 1, -1);
          updateSearchDropdownSelection();
          break;
        case "Enter":
          e.preventDefault();
          if (selectedIndex >= 0 && selectedIndex < searchResults.length) {
            selectCitation(searchResults[selectedIndex]);
          }
          break;
        case "Escape":
          e.preventDefault();
          hideSearchDropdown();
          selectedIndex = -1;
          break;
      }
    });

    // Hide dropdown when clicking outside
    document.addEventListener("click", (e) => {
      const searchContainer = document.querySelector(".zotero-search-dropdown");
      if (searchContainer && !searchContainer.contains(e.target as Node)) {
        hideSearchDropdown();
      }
    });

    // Store search results for navigation
    (window as any).setSearchResults = (results: ZoteroItemData[]) => {
      searchResults = results;
      selectedIndex = -1;
    };

    // Update dropdown selection highlighting
    function updateSearchDropdownSelection() {
      const items = document.querySelectorAll(".zotero-dropdown-item");
      items.forEach((item, index) => {
        if (index === selectedIndex) {
          item.classList.add("selected");
          // Scroll the selected item into view
          item.scrollIntoView({
            behavior: "smooth",
            block: "nearest",
            inline: "nearest",
          });
        } else {
          item.classList.remove("selected");
        }
      });
    }

    // Select a citation and close dropdown
    function selectCitation(citation: ZoteroItemData) {
      insertCitation(citation);
      hideSearchDropdown();
      (searchInput as HTMLInputElement).value = "";
      selectedIndex = -1;
    }

    // Make selectCitation available globally for click handlers
    (window as any).selectCitation = selectCitation;
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

function showSearchDropdown() {
  const dropdown = document.getElementById("search-dropdown");
  if (dropdown) {
    dropdown.classList.remove("hidden");
  }
}

function hideSearchDropdown() {
  const dropdown = document.getElementById("search-dropdown");
  if (dropdown) {
    dropdown.classList.add("hidden");
  }
}

function displaySearchResults(results: ZoteroItemData[]) {
  const resultsContainer = document.getElementById("search-results");
  if (!resultsContainer) return;

  // Store results for keyboard navigation
  (window as any).setSearchResults(results);

  if (results.length === 0) {
    resultsContainer.innerHTML = '<div class="zotero-dropdown-empty">No results found.</div>';
    showSearchDropdown();
    return;
  }

  const resultsList = results
    .map((item, index) => {
      const title = item.title || "Untitled";
      const author =
        item.creators.length > 1
          ? `${item.creators[0].lastName} et al.`
          : item.creators.length == 1
            ? item.creators[0].lastName
            : "unknown";
      const year =
        item.date && typeof item.date === "string" ? item.date.split("-")[0] : item.date || "";
      const itemString = item ? JSON.stringify(item).replace(/"/g, "&quot;") : "{}";
      return `
        <div class="zotero-dropdown-item" data-index="${index}"
             onclick="selectCitation(${itemString})">
          <div class="ms-font-m zotero-result-title">${title}</div>
          <div class="ms-font-s zotero-result-meta">${author} (${year})</div>
        </div>
      `;
    })
    .join("");

  resultsContainer.innerHTML = resultsList;
  showSearchDropdown();
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
      await showCitationsOnSlide();
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
      const year =
        citation.date && typeof citation.date === "string"
          ? citation.date.split("-")[0]
          : citation.date || "";
      const creator =
        citation.creators.length > 0
          ? (citation.creators[0].lastName ?? citation.creators[0].name ?? "unknown")
          : "unknown";
      return `
        <div class="ms-ListItem zotero-result-item citation-item"
             data-citation-id="${citation.key}">
          <div class="ms-font-m zotero-result-title">${citation.title}</div>
          <div class="ms-font-s zotero-result-meta">
            Author: ${creator} | Year: ${year}
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
