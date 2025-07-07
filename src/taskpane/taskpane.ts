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
  updateCitationKeysOrder,
  getCurrentSlide,
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
    refreshCitationsButton.onclick = () => updateCitationsPanel(false);
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
          // Show recent citations when input is empty
          showRecentCitations();
        }
        selectedIndex = -1; // Reset selection on new search
      }, SEARCH_DEBOUNCE_MS);
    });

    // Show recent citations when input is focused but empty
    searchInput.addEventListener("focus", () => {
      const query = (searchInput as HTMLInputElement).value.trim();
      if (query.length === 0) {
        showRecentCitations();
      }
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

async function updateCitationsPanel(updateSlide: boolean = true) {
  try {
    console.log("Loading current citations from slide...");
    await PowerPoint.run(async (context) => {
      const citations = await getCitationsOnSlide();
      console.log(`Found ${citations.length} citations in current slide.`);
      console.log(citations);
      displayCitationsOnTaskpane(citations);
      if (updateSlide) {
        const slide = await getCurrentSlide(context);
        await showCitationsOnSlide(slide);
        await context.sync();
      }
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
    citationsContainer.innerHTML =
      '<p class="ms-font-s" style="margin-left: 8px;">No citations found in current slide.</p>';
    return;
  }

  const citationsList = citations
    .map((citation, index) => {
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
             data-citation-id="${citation.key}"
             data-citation-index="${index}"
             draggable="true">
          <div class="citation-drag-handle">
            <i class="ms-Icon ms-Icon--GripperBarHorizontal"></i>
          </div>
          <div class="citation-content">
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
        </div>
      `;
    })
    .join("");

  citationsContainer.innerHTML = citationsList;

  // Add drag and drop event listeners
  setupCitationDragAndDrop(citations);
}

// Drag and drop functionality
function setupCitationDragAndDrop(citations: ZoteroItemData[]) {
  const citationItems = document.querySelectorAll(".citation-item");
  let draggedElement: HTMLElement | null = null;
  let draggedIndex = -1;

  citationItems.forEach((item) => {
    const element = item as HTMLElement;

    element.addEventListener("dragstart", (e) => {
      draggedElement = element;
      draggedIndex = parseInt(element.getAttribute("data-citation-index") || "-1");
      element.classList.add("dragging");

      // Set drag effect
      if (e.dataTransfer) {
        e.dataTransfer.effectAllowed = "move";
        e.dataTransfer.setData("text/html", element.outerHTML);
      }
    });

    element.addEventListener("dragend", () => {
      element.classList.remove("dragging");
      // Remove all drop indicators
      citationItems.forEach((item) => {
        item.classList.remove("drop-above", "drop-below");
      });
    });

    element.addEventListener("dragover", (e) => {
      e.preventDefault();
      if (e.dataTransfer) {
        e.dataTransfer.dropEffect = "move";
      }

      if (draggedElement && draggedElement !== element) {
        const rect = element.getBoundingClientRect();
        const midPoint = rect.top + rect.height / 2;
        const mouseY = e.clientY;

        // Remove previous indicators
        element.classList.remove("drop-above", "drop-below");

        // Add appropriate indicator
        if (mouseY < midPoint) {
          element.classList.add("drop-above");
        } else {
          element.classList.add("drop-below");
        }
      }
    });

    element.addEventListener("dragleave", (e) => {
      // Only remove indicators if we're actually leaving the element
      const rect = element.getBoundingClientRect();
      const x = e.clientX;
      const y = e.clientY;

      if (x < rect.left || x > rect.right || y < rect.top || y > rect.bottom) {
        element.classList.remove("drop-above", "drop-below");
      }
    });

    element.addEventListener("drop", (e) => {
      e.preventDefault();

      if (!draggedElement || draggedElement === element) {
        return;
      }

      const targetIndex = parseInt(element.getAttribute("data-citation-index") || "-1");
      const rect = element.getBoundingClientRect();
      const midPoint = rect.top + rect.height / 2;
      const mouseY = e.clientY;

      let newIndex = targetIndex;
      if (mouseY > midPoint) {
        newIndex = targetIndex + 1;
      }

      // Adjust for items moving up vs down
      if (draggedIndex < newIndex) {
        newIndex--;
      }

      if (draggedIndex !== newIndex && draggedIndex >= 0 && newIndex >= 0) {
        reorderCitations(citations, draggedIndex, newIndex);
      }

      // Clean up
      element.classList.remove("drop-above", "drop-below");
    });
  });
}

async function reorderCitations(citations: ZoteroItemData[], fromIndex: number, toIndex: number) {
  try {
    console.log(`Reordering citation from index ${fromIndex} to ${toIndex}`);

    // Create new array with reordered citations
    const reorderedCitations = [...citations];
    const [movedCitation] = reorderedCitations.splice(fromIndex, 1);
    reorderedCitations.splice(toIndex, 0, movedCitation);

    // Update the slide with the new order
    await updateCitationOrder(reorderedCitations);

    // Refresh the display
    setTimeout(updateCitationsPanel, 500);
  } catch (error) {
    console.error("Error reordering citations:", error);
  }
}

async function updateCitationOrder(reorderedCitations: ZoteroItemData[]) {
  try {
    // Extract the citation keys in the new order
    const orderedKeys = reorderedCitations.map((c) => c.key);

    // Update the citation key order on the slide
    await updateCitationKeysOrder(orderedKeys);

    console.log("Citation order updated successfully:", orderedKeys);

    // Refresh the display to reflect the new order
    await updateCitationsPanel();
  } catch (error) {
    console.error("Failed to update citation order:", error);
    // Optionally show user feedback about the error
    throw error;
  }
}

async function showRecentCitations() {
  try {
    const citationStore = CitationStore.getInstance();

    // Use the new getLast method to get the 5 most recently added citations
    const recentCitations = await citationStore.getRecent(5);

    const resultsContainer = document.getElementById("search-results");
    if (!resultsContainer) return;

    // Store results for keyboard navigation
    (window as any).setSearchResults(recentCitations);

    if (recentCitations.length === 0) {
      resultsContainer.innerHTML = '<div class="zotero-dropdown-empty">No recent citations.</div>';
      showSearchDropdown();
      return;
    }

    const resultsList = recentCitations
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
            <div class="ms-font-s zotero-result-meta">${author} (${year}) • Recent</div>
          </div>
        `;
      })
      .join("");

    resultsContainer.innerHTML = resultsList;
    showSearchDropdown();
  } catch (error) {
    console.error("Error loading recent citations:", error);
    const resultsContainer = document.getElementById("search-results");
    if (resultsContainer) {
      resultsContainer.innerHTML =
        '<div class="zotero-dropdown-empty">Error loading recent citations.</div>';
      showSearchDropdown();
    }
  }
}
