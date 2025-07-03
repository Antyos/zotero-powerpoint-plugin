/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/// <reference types="office-js" />

import { ZoteroLibrary } from "../zotero/zotero-connector";

Office.onReady(() => {
  const zotero = ZoteroLibrary.getInstance();
  zotero.loadConfig();
  // zotero.updateConfig({apiKey: "ie7sTgek6mfDDq1T8un9tTtf", userId: 11539091});
  // If needed, Office.js is ready to be called.
});

/**
 * Insert Zotero citation into the current slide
 * @param event
 */
function insertZoteroCitation(event: Office.AddinCommands.Event) {
  return PowerPoint.run(async () => {
    try {
      console.log("Starting citation insertion...");
      // const zotero = await ZoteroLibrary.getClient();

      // Initialize if not ready
      // if (!zotero.isReady()) {
      //   console.log('Initializing Zotero integration...');
      //   const connected = await zotero.checkConnection();
      //   if (!connected) {
      //     console.error('Zotero initialization failed');
      //     event.completed();
      //     return;
      //   }
      //   console.log('Zotero integration initialized successfully');
      // }

      // Show loading message
      // await showInfoMessage(context, 'Opening Zotero citation picker...');

      // Send addCitation command
      // console.log('Sending addCitation command...');
      // const citation = await zotero.sendZoteroCommand('addCitation', true);

      // if (citation) {
      //   console.log('Citation received, inserting into slide...');
      //   // Insert citation field
      //   await zotero.insertField(
      //     `ADDIN ZOTERO_ITEM CSL_CITATION {"citationID":"${Date.now()}","properties":{},"citationItems":[]}`,
      //     citation,
      //     0
      //   );

      //   await showSuccessMessage(context, 'Citation inserted successfully');
      //   console.log('Citation insertion completed successfully');
      // } else {
      //   console.log('No citation received from Zotero');
      //   await showInfoMessage(context, 'Citation selection cancelled or no items selected. Please try again.');
      // }

      event.completed();
    } catch (error) {
      console.error("Error inserting citation:", error);

      let errorMsg = "Unknown error occurred";
      if (error instanceof Error) {
        errorMsg = error.message;
      }

      console.error(`Error inserting citation: ${errorMsg}`);
      event.completed();
    }
  });
}

/**
 * Test Zotero connection - for debugging purposes
 * @param event
 */
function testZoteroConnection(event: Office.AddinCommands.Event) {
  return PowerPoint.run(async (context) => {
    try {
      console.log("Testing Zotero connection...");
      const zotero = ZoteroLibrary.getInstance();

      // Get test results
      const testResults = await zotero.checkConnection();

      console.log(`Connected: ${testResults}`);
      // console.log("Library:", (await ZoteroLibrary.getClient()))
      ZoteroLibrary.getClient()
        .items()
        .get()
        .then((items) => {
          console.log("Zotero items:", items);
        });
      // Show results

      await context.sync();
      event.completed();
    } catch (error) {
      console.error("Error testing Zotero connection:", error);
      event.completed();
    }
  });
}

/**
 * Open Zotero configuration dialog
 * @param event
 */
function openZoteroConfig(event: Office.AddinCommands.Event) {
  return PowerPoint.run(async () => {
    try {
      console.log("Opening Zotero configuration dialog...");
      const zotero = ZoteroLibrary.getInstance();

      const configured = await zotero.configureFromDialog();

      if (configured) {
        console.log("Zotero configuration completed successfully");
      } else {
        console.log("Zotero configuration was cancelled");
      }

      event.completed();
    } catch (error) {
      console.error("Error opening Zotero configuration:", error);
      event.completed();
    }
  });
}

// Register the functions with Office.
Office.actions.associate("insertZoteroCitation", insertZoteroCitation);
Office.actions.associate("testZoteroConnection", testZoteroConnection);
Office.actions.associate("openZoteroConfig", openZoteroConfig);
