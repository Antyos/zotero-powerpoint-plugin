# Build PowerPoint add-ins using Office Add-ins Development Kit

PowerPoint add-ins are integrations built by third parties into PowerPoint by using [PowerPoint JavaScript API](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/powerpoint-add-ins-reference-overview) and [Office Platform capabilities](https://learn.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins).

## How to run this project

### Prerequisites

- Node.js (the latest LTS version). Visit the [Node.js site](https://nodejs.org/) to download and install the right version for your operating system. To verify that you've already installed these tools, run the commands `node -v` and `npm -v` in your terminal.
- Office connected to a Microsoft 365 subscription. You might qualify for a Microsoft 365 E5 developer subscription through the [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program), see [FAQ](https://learn.microsoft.com/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-) for details. Alternatively, you can [sign up for a 1-month free trial](https://www.microsoft.com/microsoft-365/try?rtc=1) or [purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/buy/compare-all-microsoft-365-products).

### Run the add-in using Office Add-ins Development Kit extension

1. **Open the Office Add-ins Development Kit**

    In the **Activity Bar**, select the **Office Add-ins Development Kit** icon to open the extension.

1. **Preview Your Office Add-in (F5)**

    Select **Preview Your Office Add-in(F5)** to launch the add-in and debug the code. In the Quick Pick menu, select the option **PowerPoint Desktop (Edge Chromium)**.

    The extension then checks that the prerequisites are met before debugging starts. Check the terminal for detailed information if there are issues with your environment. After this process, the PowerPoint desktop application launches and sideloads the add-in.

1. **Stop Previewing Your Office Add-in**

    Once you are finished testing and debugging the add-in, select **Stop Previewing Your Office Add-in**. This closes the web server and removes the add-in from the registry and cache.

## Use the add-in project

The add-in project that you've created contains code for a basic task pane add-in.

## Explore the add-in code

To explore an Office add-in project, you can start with the key files listed below.

- The `./manifest.xml` file in the root directory of the project defines the settings and capabilities of the add-in.  <br>You can check whether your manifest file is valid by selecting **Validate Manifest File** option from the Office Add-ins Development Kit.
- The `./src/taskpane/taskpane.html` file contains the HTML markup for the task pane.
- The `./src/taskpane/taskpane.css` file contains the CSS that's applied to content in the task pane.
- The `./src/taskpane/taskpane.ts` file contains the Office JavaScript API code that facilitates interaction between the task pane and the PowerPoint application.

## Troubleshooting

If you have problems running the add-in, take these steps.

- Close any open instances of PowerPoint.
- Close the previous web server started for the add-in with the **Stop Previewing Your Office Add-in** Office Add-ins Development Kit extension option.

If you still have problems, see [troubleshoot development errors](https://learn.microsoft.com//office/dev/add-ins/testing/troubleshoot-development-errors) or [create a GitHub issue](https://aka.ms/officedevkitnewissue) and we'll help you.  

For information on running the add-in on PowerPoint on the web, see [Sideload Office Add-ins to Office on the web](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing).

For information on debugging on older versions of Office, see [Debug add-ins using developer tools in Microsoft Edge Legacy](https://learn.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-devtools-edge-legacy).

## Make code changes

All the information about Office Add-ins is found in our [official documentation](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins). You can also explore more samples in the Office Add-ins Development Kit. Select **View Samples** to see more samples of real-world scenarios.

If you edit the manifest as part of your changes, use the **Validate Manifest File** option in the Office Add-ins Development Kit. This shows you errors in the manifest syntax.

## Engage with the team

Did you experience any problems? [Create an issue](https://aka.ms/officedevkitnewissue) and we'll help you out.

Want to learn more about new features and best practices for the Office platform? [Join the Microsoft Office Add-ins community call](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins-community-call).

## Copyright

Copyright (c) 2024 Microsoft Corporation. All rights reserved.

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

# Zotero PowerPoint Integration Add-in

This PowerPoint add-in integrates with Zotero to allow citation insertion directly into PowerPoint slides, modeled after the official Zotero Word integration but adapted for the Office.js/PowerPoint environment.

## Features

- **Citation Insertion**: Insert citations from Zotero into PowerPoint slides
- **Multiple Communication Methods**: Supports both Better BibTeX HTTP API and Zotero integration service protocols
- **Field Management**: Stores citation metadata similar to the Word integration
- **Diagnostic Tools**: Built-in connection testing and troubleshooting

## Prerequisites

- **Zotero**: Install the latest version of [Zotero](https://www.zotero.org/download/)
- **Better BibTeX Plugin**: Install the [Better BibTeX plugin](https://retorque.re/zotero-better-bibtex/) for Zotero
- **Node.js**: Latest LTS version from [nodejs.org](https://nodejs.org/)
- **Office 365**: PowerPoint with Office.js support

## Setup Instructions

### 1. Install Zotero Components

1. **Install Zotero** from [zotero.org](https://www.zotero.org/download/)
2. **Install Better BibTeX plugin**:
   - Download the latest `.xpi` file from [Better BibTeX releases](https://github.com/retorque-re/zotero-better-bibtex/releases)
   - In Zotero: Tools → Add-ons → Install Add-on From File
   - Select the downloaded `.xpi` file
   - Restart Zotero

### 2. Configure Zotero

1. **Enable Better BibTeX**:
   - Go to Zotero Preferences → Better BibTeX
   - Ensure "Enable export by HTTP" is checked
   - Note the port (default: 23119)

2. **Test Zotero Connection**:
   - Open a web browser
   - Visit: `http://127.0.0.1:23119/better-bibtex/cayw?probe=true`
   - Should return "ready" if working correctly

### 3. Install and Run the Add-in

1. Clone or download this repository
2. Open terminal in the project directory
3. Run: `npm install`
4. Run: `npm run build:dev`
5. Run: `npm start` to start the development server

## Using the Add-in

### Insert Citations

1. **Start Zotero** and ensure it's running
2. **Open PowerPoint** with the add-in loaded
3. **Select a slide** where you want to insert a citation
4. **Click "Insert Citation"** button in the Zotero Integration ribbon
5. **Select items** in the Zotero picker dialog that opens
6. **Citation will be inserted** as a formatted text box on your slide

### Test Connection

Use the **"Test Connection"** button to diagnose connectivity issues:

- Tests Better BibTeX availability
- Tests Zotero integration service
- Shows detailed connection status
- Displays test results directly in your slide

## Troubleshooting

### "Citation selection cancelled or no items selected"

This error typically indicates connection issues with Zotero:

1. **Check Zotero is Running**:
   - Ensure Zotero application is open and running
   - Check system tray for Zotero icon

2. **Verify Better BibTeX**:
   - In Zotero: Help → Check for Updates
   - Ensure Better BibTeX is installed and enabled
   - Test the URL: `http://127.0.0.1:23119/better-bibtex/cayw?probe=true`

3. **Use Test Connection**:
   - Click the "Test Connection" button in PowerPoint
   - Review the diagnostic information shown
   - Common issues:
     - Port 23119 blocked by firewall
     - Better BibTeX not installed
     - Zotero not running

4. **Check Browser Console**:
   - Press F12 in the add-in
   - Look for error messages in the console
   - Network errors indicate connectivity issues

5. **Firewall/Antivirus**:
   - Ensure localhost connections on port 23119 are allowed
   - Some security software blocks local HTTP requests

### Common Solutions

- **Restart Zotero**: Close and reopen Zotero completely
- **Reinstall Better BibTeX**: Remove and reinstall the plugin
- **Check Port**: Verify port 23119 is available (netstat -an | findstr 23119)
- **Try Different Port**: Configure Zotero to use a different port

## Technical Implementation

This add-in attempts to replicate the Word integration's approach within Office.js constraints:

- **No Direct DLL Access**: Office.js sandbox prevents loading `libzoteroWinWordIntegration.dll`
- **HTTP Communication**: Uses HTTP APIs instead of OLE Automation
- **Field Simulation**: Stores citation metadata in PowerPoint shapes and document settings
- **Multiple Protocols**: Tries Better BibTeX, integration service, and JSON-RPC methods

### Communication Methods (in order of preference)

1. **Better BibTeX CAYW**: `http://127.0.0.1:23119/better-bibtex/cayw`
2. **Integration Service**: `http://127.0.0.1:23119/integration/`
3. **JSON-RPC**: `http://127.0.0.1:23119/integration/json-rpc`
