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

  // Settings panel elements
  const configureButton = document.getElementById("configure-zotero");
  if (configureButton) {
    configureButton.onclick = showSettingsPanel;
  }

  const closeSettingsButton = document.getElementById("close-settings");
  if (closeSettingsButton) {
    closeSettingsButton.onclick = hideSettingsPanel;
  }

  const settingsForm = document.getElementById("settings-form");
  if (settingsForm) {
    settingsForm.onsubmit = handleSettingsSubmit;
  }

  const settingsCancelButton = document.getElementById("settings-cancel");
  if (settingsCancelButton) {
    settingsCancelButton.onclick = hideSettingsPanel;
  }

  // Set up live JSON validation for citation formats
  setupCitationFormatsValidation();

  const refreshCitationsButton = document.getElementById("refresh-citations");
  if (refreshCitationsButton) {
    refreshCitationsButton.onclick = () => updateCitationsPanel(true);
  }

  const debugSlideTagsButton = document.getElementById("debug-slide-tags");
  if (debugSlideTagsButton) {
    debugSlideTagsButton.onclick = () => {
      debugSlideTags().catch((error) => {
        console.error("Error debugging slide tags:", error);
      });
    };
  }

  const debugCitationsButton = document.getElementById("debug-citations");
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

  const debugCitationStoreButton = document.getElementById("debug-citation-store");
  if (debugCitationStoreButton) {
    debugCitationStoreButton.onclick = () => {
      console.log("Debugging citation store...");
      CitationStore.getInstance().debugXml();
    };
  }

  const insertMockCitationButton = document.getElementById("insert-mock-citation");
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

  const clearCitationStoreButton = document.getElementById("clear-citation-store");
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

  const searchInput = document.getElementById("search-query");
  if (searchInput) {
    // Debounced search as user types
    let searchTimeout: ReturnType<typeof setTimeout>;
    let selectedIndex = -1;
    let searchResults: ZoteroItemData[] = [];

    searchInput.addEventListener("input", () => {
      clearTimeout(searchTimeout);
      searchTimeout = setTimeout(() => {
        const query = (searchInput as HTMLInputElement).value.trim();
        console.log(`Search input changed: "${query}"`);
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
      if (!dropdown || dropdown.hidden) {
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
        case "Tab":
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
      const target = e.target as Node;
      if (searchContainer && !searchContainer.contains(target)) {
        hideSearchDropdown();
      }
    });

    // Store search results for navigation
    (window as any).setSearchResults = (results: ZoteroItemData[]) => {
      searchResults = results;
      selectedIndex = -1;
    };

    // Update dropdown selection highlighting
    const updateSearchDropdownSelection = () => {
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
    };

    // Select a citation and close dropdown
    const selectCitation = (citation: ZoteroItemData) => {
      insertCitation(citation);
      hideSearchDropdown();
      (searchInput as HTMLInputElement).value = "";
      selectedIndex = -1;
    };

    // Make selectCitation available globally for click handlers
    (window as any).selectCitation = selectCitation;
  }

  // Load current citations on initialization
  updateCitationsPanel(false);
}

function _showSettingsPanel(show: boolean = true) {
  const showOrHide: string = show ? "Showing" : "Hiding";
  try {
    console.debug(`${showOrHide} settings panel...`);
    // Load current configuration into the form
    if (show) {
      loadCurrentSettingsConfig();
    }
    const settingsPanel = document.getElementById("settings-panel");
    if (settingsPanel) {
      settingsPanel.hidden = !show;
    }
    const mainContent = document.getElementById("main-content");
    if (mainContent) {
      mainContent.hidden = show;
    }
  } catch (error) {
    console.error(`Error ${showOrHide.toLowerCase()} settings panel:`, error);
  }
}

function showSettingsPanel() {
  return _showSettingsPanel(true);
}

function hideSettingsPanel() {
  return _showSettingsPanel(false);
}

function loadCurrentSettingsConfig() {
  try {
    // Use the imported default configuration
    const config = ZoteroLibrary.getInstance().getConfig();

    // Populate form fields (except API key for security)
    if (config.userId) {
      (document.getElementById("settings-user-id") as HTMLInputElement).value =
        config.userId.toString();
    }
    if (config.userType) {
      (document.getElementById("settings-user-type") as HTMLSelectElement).value = config.userType;
    }
    if (config.searchResultsLimit) {
      (document.getElementById("settings-search-limit") as HTMLInputElement).value =
        config.searchResultsLimit.toString();
    }
    if (config.citationFormats) {
      (document.getElementById("settings-citation-formats") as HTMLTextAreaElement).value =
        JSON.stringify(config.citationFormats, null, 2);
      populateSettingsCitationFormatOptions(config.citationFormats);
    }
    if (config.selectedCitationFormat) {
      (document.getElementById("settings-selected-format") as HTMLSelectElement).value =
        config.selectedCitationFormat;
    }
    if (config.citationShapeName) {
      (document.getElementById("settings-citation-shape") as HTMLInputElement).value =
        config.citationShapeName;
    }

    console.log("Current config loaded into settings form");
  } catch (error) {
    console.error("Error loading current settings config:", error);
  }
}

function populateSettingsCitationFormatOptions(formats: any) {
  const select = document.getElementById("settings-selected-format") as HTMLSelectElement;
  const currentValue = select.value;

  // Clear existing options except the first one
  select.innerHTML = '<option value="">Select a format...</option>';

  // Add options from formats object
  for (const [key] of Object.entries(formats)) {
    const option = document.createElement("option");
    option.value = key;
    // Display the key as the label
    option.textContent = key;
    select.appendChild(option);
  }

  // Restore selection if it still exists
  if (currentValue && formats[currentValue]) {
    select.value = currentValue;
  }
}

async function handleSettingsSubmit(event: Event) {
  event.preventDefault();

  try {
    console.log("Handling settings form submission...");

    // Get form values
    const apiKeyInput = (document.getElementById("settings-api-key") as HTMLInputElement).value;
    const userId = parseInt(
      (document.getElementById("settings-user-id") as HTMLInputElement).value
    );
    const userType = (document.getElementById("settings-user-type") as HTMLSelectElement).value;
    const searchResultsLimit =
      parseInt((document.getElementById("settings-search-limit") as HTMLInputElement).value) || 5;
    const selectedCitationFormat =
      (document.getElementById("settings-selected-format") as HTMLSelectElement).value || undefined;
    const citationShapeName =
      (document.getElementById("settings-citation-shape") as HTMLInputElement).value || "Citation";

    // Handle API key - use new one if provided, otherwise keep existing
    let apiKey = apiKeyInput;
    if (!apiKey) {
      // Try to get existing API key
      try {
        const partitionKey = (Office as any).context?.partitionKey || "default";
        const settingsJson = localStorage.getItem(`${partitionKey}-zotero-settings`);
        if (settingsJson) {
          const existingConfig = JSON.parse(settingsJson);
          apiKey = existingConfig.apiKey;
        }
      } catch (error) {
        console.error("Error getting existing API key:", error);
      }
    }

    // Parse citation formats if provided
    let citationFormats = undefined;
    const citationFormatsText = (
      document.getElementById("settings-citation-formats") as HTMLTextAreaElement
    ).value.trim();

    if (citationFormatsText) {
      // Validate before saving
      const errorContainer = document.getElementById(
        "settings-citation-formats-error"
      ) as HTMLDivElement;

      const isValid = validateCitationFormats(citationFormatsText, errorContainer);
      if (!isValid) {
        console.error("Invalid citation formats, cannot save");
        return;
      }

      try {
        citationFormats = JSON.parse(citationFormatsText);
        console.log("Parsed citation formats:", citationFormats);
      } catch (error) {
        console.error("Error parsing citation formats:", error);
        return;
      }
    }

    const configToSave = {
      apiKey: apiKey,
      userId: userId,
      userType: userType as "user" | "group",
      searchResultsLimit: searchResultsLimit,
      citationFormats: citationFormats,
      selectedCitationFormat: selectedCitationFormat,
      citationShapeName: citationShapeName,
    };

    console.log("Saving configuration:", configToSave);

    // Save the configuration
    const zotero = ZoteroLibrary.getInstance();
    await zotero.updateConfig(configToSave);

    console.log("Configuration saved successfully!");

    hideSettingsPanel();
  } catch (error) {
    console.error("Error saving settings:", error);
  }
}

async function searchZoteroLibrary() {
  try {
    const searchInput = document.getElementById("search-query") as HTMLInputElement;
    const query = searchInput?.value?.trim();

    const resultsContainer = document.getElementById("search-results");
    if (resultsContainer) {
      resultsContainer.innerHTML = `<div class="zotero-dropdown-loading">Searching...</div>`;
    }
    console.log(`Searching Zotero library for: "${query}"`);

    if (!query) {
      console.log("Empty query, skipping search");
      return;
    }

    console.log("Searching Zotero library...");
    const zotero = ZoteroLibrary.getInstance();

    if (!zotero.isConfigured()) {
      console.log("Zotero not configured, showing configuration message");
      if (resultsContainer) {
        resultsContainer.innerHTML =
          '<div class="zotero-dropdown-empty">Please configure Zotero API settings first.</div>';
        showSearchDropdown();
      }
      return;
    }

    const results = await zotero.searchItems(query);
    displaySearchResults(results);
    console.log(`Search completed. Found ${results.length} items.`);
  } catch (error) {
    console.error("Search error:", error);
    const resultsContainer = document.getElementById("search-results");
    if (resultsContainer) {
      resultsContainer.innerHTML =
        '<div class="zotero-dropdown-empty">Search failed. Please check your configuration.</div>';
      showSearchDropdown();
    }
  }
}

function showSearchDropdown() {
  const dropdown = document.getElementById("search-dropdown");
  if (dropdown) {
    console.debug("Showing search dropdown");
    dropdown.hidden = false;
  } else {
    console.error("Search dropdown element not found");
  }
}

function hideSearchDropdown() {
  const dropdown = document.getElementById("search-dropdown");
  if (dropdown) {
    console.debug("Hiding search dropdown");
    dropdown.hidden = true;
  }
}

function displaySearchResults(results: ZoteroItemData[]) {
  const resultsContainer = document.getElementById("search-results");
  if (!resultsContainer) {
    console.error("Search results container not found");
    return;
  }

  console.debug(`Displaying ${results.length} search results`);

  // Store results for keyboard navigation
  if (typeof (window as any).setSearchResults === "function") {
    (window as any).setSearchResults(results);
  }

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
             onclick="window.selectCitation && window.selectCitation(${itemString})">
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
    const success = await removeCitationFromSlide(citationId);
    if (success) {
      console.log(`Removed citation: ${citationId}`);
      // Refresh the current citations list
      setTimeout(updateCitationsPanel, 500);
    } else {
      console.warn(`Could not find citation ${citationId} for removal.`);
    }
  } catch (error) {
    console.error(`Failed to remove citation ${citationId}:`, error);
  }
}
(window as any).removeCitation = removeCitation;

async function updateCitationsPanel(updateSlide: boolean = true) {
  try {
    console.debug("updateCitationsPanel(): Loading current citations from slide...");
    await PowerPoint.run(async (context) => {
      const citations = await getCitationsOnSlide();
      console.debug(`Found ${citations.length} citations in current slide.`);
      console.debug(citations);
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
  if (!citationsContainer) {
    return;
  }

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
    console.debug(`Reordering citation from index ${fromIndex} to ${toIndex}`);

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
    console.debug("Citation order updated successfully:", orderedKeys);
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
    if (!resultsContainer) {
      return;
    }

    // Store results for keyboard navigation
    if (typeof (window as any).setSearchResults === "function") {
      (window as any).setSearchResults(recentCitations);
    }
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
               onclick="window.selectCitation && window.selectCitation(${itemString})">
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

// Live JSON validation for citation formats
function setupCitationFormatsValidation() {
  const formatsTextArea = document.getElementById(
    "settings-citation-formats"
  ) as HTMLTextAreaElement;
  const errorContainer = document.getElementById(
    "settings-citation-formats-error"
  ) as HTMLDivElement;
  const saveButton = document.getElementById("settings-save") as HTMLButtonElement;

  if (!formatsTextArea || !errorContainer || !saveButton) {
    console.warn("Citation formats validation elements not found");
    return;
  }

  // Initial validation
  validateCitationFormats(formatsTextArea.value, errorContainer, saveButton);

  // Live validation on input with debouncing
  let validationTimeout: ReturnType<typeof setTimeout>;
  formatsTextArea.addEventListener("input", () => {
    clearTimeout(validationTimeout);
    validationTimeout = setTimeout(() => {
      validateCitationFormats(formatsTextArea.value, errorContainer, saveButton);
    }, 300);
  });
}

// Validate citation formats JSON and update dropdown
function validateCitationFormats(
  formatsText: string,
  errorContainer: HTMLDivElement,
  saveButton?: HTMLButtonElement
): boolean {
  try {
    // Clear previous error
    errorContainer.textContent = "";
    errorContainer.style.display = "none";

    // If empty, that's okay - just clear the dropdown
    if (!formatsText.trim()) {
      populateSettingsCitationFormatOptions({});
      if (saveButton) {
        saveButton.disabled = false;
      }
      return true;
    }

    // Parse the JSON
    const formatsJson = JSON.parse(formatsText);

    // Basic structure check
    if (typeof formatsJson !== "object" || formatsJson === null) {
      throw new Error("Citation formats must be a JSON object");
    }

    // Check format structure
    const formatKeys = Object.keys(formatsJson);
    for (const key of formatKeys) {
      const format = formatsJson[key];
      if (typeof format !== "object" || format === null) {
        throw new Error(`Format "${key}" must be an object`);
      }
      if (!format.format || typeof format.format !== "string") {
        throw new Error(`Format "${key}" must have a "format" property with a string value`);
      }
      // delimiter is optional, but if present must be a string
      if (format.delimiter && typeof format.delimiter !== "string") {
        throw new Error(`Format "${key}" delimiter must be a string`);
      }
    }

    // If everything is valid, update the dropdown and enable save button
    populateSettingsCitationFormatOptions(formatsJson);
    if (saveButton) {
      saveButton.disabled = false;
    }
    return true;
  } catch (error) {
    // Show error message
    const errorMessage = error instanceof Error ? error.message : "Invalid JSON format";
    errorContainer.textContent = errorMessage;
    errorContainer.style.display = "block";

    // Clear the dropdown and disable save button on error
    populateSettingsCitationFormatOptions({});
    if (saveButton) {
      saveButton.disabled = true;
    }
    return false;
  }
}
