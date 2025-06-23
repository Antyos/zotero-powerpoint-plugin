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
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message.
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    "ActionPerformanceNotification",
    message
  );

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

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
          console.log('Zotero initialization failed');
          await showZoteroNotAvailableError(context);
          event.completed();
          return;
        }
        console.log('Zotero integration initialized successfully');
      }

      // Show loading message
      await showInfoMessage(context, 'Opening Zotero citation picker...');

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
      
      await showErrorMessage(context, `Error inserting citation: ${errorMsg}`);
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
      const integration = ZoteroBBTConnector.getInstance();
      
      // Get test results
      const testResults = await integration.testConnection();
      
      // Show results
      const resultText = 'Zotero Connection Test Results:\n\n' + testResults.join('\n');
      
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const textBox = slide.shapes.addTextBox(resultText, {
        left: 50,
        top: 50,
        width: 600,
        height: 300
      });
      
      textBox.textFrame.textRange.font.name = "Consolas";
      textBox.textFrame.textRange.font.size = 10;
      textBox.fill.setSolidColor("#F0F0F0");
      textBox.lineFormat.color = "#666666";
      
      await context.sync();
      console.log('Test results inserted into slide');
      
      event.completed();
      
    } catch (error) {
      console.error("Error testing Zotero connection:", error);
      await showErrorMessage(context, `Error testing connection: ${error.message}`);
      event.completed();
    }
  });
}

/**
 * Show error when Zotero is not available
 */
async function showZoteroNotAvailableError(context: PowerPoint.RequestContext) {
  try {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const textBox = slide.shapes.addTextBox(
      "Zotero Integration Error\n\nZotero is not running or not properly configured.\n\nPlease:\n1. Start Zotero\n2. Install Better BibTeX plugin\n3. Ensure Zotero is configured for integration\n4. Try again", 
      {
        left: 50,
        top: 50,
        width: 500,
        height: 150
      }
    );
    
    textBox.textFrame.textRange.font.color = "#CC0000";
    textBox.textFrame.textRange.font.size = 12;
    textBox.textFrame.textRange.font.name = "Calibri";
    textBox.fill.setSolidColor("#FFE6E6");
    
    await context.sync();
  } catch (error) {
    console.error("Error showing Zotero not available error:", error);
  }
}

/**
 * Show success message
 */
async function showSuccessMessage(context: PowerPoint.RequestContext, message: string) {
  console.log('Success:', message);
  // Could show a temporary notification or update status
}

/**
 * Show informational message
 */
async function showInfoMessage(context: PowerPoint.RequestContext, message: string) {
  try {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const textBox = slide.shapes.addTextBox(message, {
      left: 100,
      top: 100,
      width: 400,
      height: 50
    });
    
    textBox.textFrame.textRange.font.color = "#888888";
    textBox.textFrame.textRange.font.size = 11;
    textBox.textFrame.textRange.font.name = "Calibri";
    
    await context.sync();
  } catch (error) {
    console.error("Error showing info message:", error);
  }
}

/**
 * Show error message
 */
async function showErrorMessage(context: PowerPoint.RequestContext, errorMessage: string) {
  try {
    const slide = context.presentation.getSelectedSlides().getItemAt(0);
    const textBox = slide.shapes.addTextBox(errorMessage, {
      left: 100,
      top: 100,
      width: 500,
      height: 100
    });
    
    textBox.textFrame.textRange.font.color = "#CC0000";
    textBox.textFrame.textRange.font.size = 11;
    textBox.textFrame.textRange.font.name = "Calibri";
    
    await context.sync();
  } catch (error) {
    console.error("Error showing error message:", error);
  }
}

// Register the functions with Office.
Office.actions.associate("action", action);
Office.actions.associate("insertZoteroCitation", insertZoteroCitation);
Office.actions.associate("testZoteroConnection", testZoteroConnection);
