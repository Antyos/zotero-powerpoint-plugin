/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/// <reference types="office-js" />

import { ZoteroBBTConnector } from '../zotero/zotero-connector';

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});


/**
 * Insert Zotero citation into the current slide
 * @param event
 */
function insertZoteroCitation(event: Office.AddinCommands.Event) {
  return PowerPoint.run(async (context) => {
    try {
      console.log('Starting citation insertion...');
      const zotero = ZoteroBBTConnector.getInstance();
      
      // Initialize if not ready
      if (!zotero.isReady()) {
        console.log('Initializing Zotero integration...');
        const connected = await zotero.checkConnection();
        if (!connected) {
          console.error('Zotero initialization failed');
          event.completed();
          return;
        }
        console.log('Zotero integration initialized successfully');
      }

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
      
      let errorMsg = 'Unknown error occurred';
      if (error.message) {
        errorMsg = error.message;
      }
      
      // Alert.Show(`Error inserting citation: ${errorMsg}`);
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
      console.log('Testing Zotero connection...');
      const zotero = ZoteroBBTConnector.getInstance();
      
      // Get test results
      const testResults = await zotero.testConnection();
      
      console.log('Test results:', testResults);
      // Show results
      const resultText = `Zotero Connection Test Results:\n\n`+testResults.join('\n');
      
      await context.sync();
      console.log(resultText);
      
      event.completed();
      
    } catch (error) {
      console.error("Error testing Zotero connection:", error);
      event.completed();
    }
  });
}

/**
 * Test headers and display results
 */
function testHeaders(event: Office.AddinCommands.Event) {
  return PowerPoint.run(async (context) => {
    try {
      console.log('Testing headers functionality...');
      
      const connector = ZoteroBBTConnector.getInstance();
      const results = await connector.debugHeaders();
      
      const message = `=== HEADER TEST RESULTS ===\n\n` +
        `Request Attempt:\n${JSON.stringify(results.requestAttempt, null, 2)}\n\n` +
        `Response Headers:\n${JSON.stringify(results.responseHeaders, null, 2)}\n\n` +
        `User Agent Issue:\n${results.userAgentIssue}`;
      
      console.log(message);
      event.completed();
      
    } catch (error) {
      console.error('Header test failed:', error);
      event.completed();
    }
  });
}

// Register the functions with Office.
Office.actions.associate("insertZoteroCitation", insertZoteroCitation);
Office.actions.associate("testZoteroConnection", testZoteroConnection);
Office.actions.associate("testHeaders", testHeaders);
