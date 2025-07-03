/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { ZoteroLibrary, TitleCreatorDate } from "../zotero/zotero-connector";

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
  const searchInput = document.getElementById("search-query") as HTMLInputElement;

  if (configureButton) {
    configureButton.onclick = configureZotero;
  }

  if (searchButton) {
    searchButton.onclick = searchZoteroLibrary;
  }

  if (testButton) {
    testButton.onclick = testZoteroConnection;
  }

  if (searchInput) {
    searchInput.addEventListener("keypress", (e) => {
      if (e.key === "Enter") {
        searchZoteroLibrary();
      }
    });
  }
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
(window as any).insertCitation = async function (
  itemId: string,
  author: string,
  year: string = ""
) {
  try {
    console.log(`Inserting citation: ${author}, ${year}`);

    // Insert the citation into PowerPoint
    const options: Office.SetSelectedDataOptions = { coercionType: Office.CoercionType.Text };
    await Office.context.document.setSelectedDataAsync(`[${author}, ${year}]`, options);

    console.log(`Citation inserted: ${author}, ${year}`);
  } catch (error) {
    console.error("Citation insertion error:", error);
  }
};

export async function run() {
  /**
   * Insert your PowerPoint code here
   */
  const options: Office.SetSelectedDataOptions = { coercionType: Office.CoercionType.Text };

  await Office.context.document.setSelectedDataAsync(" ", options);
  await Office.context.document.setSelectedDataAsync("Hello World!", options);
}
