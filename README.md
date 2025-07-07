# Zotero PowerPoint Integration Add-in

A PowerPoint Office Add-in that integrates with Zotero to provide seamless
citation management directly within PowerPoint presentations. The add-in
features a task pane interface for searching your Zotero library, inserting
citations, and managing citations on slides with persistent storage and flexible
configuration.

## Features

- Insert citations from your Zotero library directly into PowerPoint slides.
- Customize in-slide citation formats.

## Getting Started

### Installation

These are temporary instructions until the add-in is published to Microsoft AppSource.

For the full instructions see: <https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins>

1. Download the `dist/` folder and extract it to a local directory.
2. Create a local network share for the extracted folder by right clicking on
   the folder in Windows Explorer, clicking **Properties > Sharing > Share** and
   click "Share". (You might need to add "Everyone" with "Read/Write"
   permissions). Copy the shared folder path, e.g.
   `file://C:/Users/YourName/Documents/ZoteroPowerPointAddin/dist/`.
3. Open PowerPoint and go to **File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs**. Paste the shared folder path into the "Catalog Url" box and click "Add catalog". Make sure to click the "Show in Menu" checkbox.
4. Restart PowerPoint.
5. On the ribbon, go to "Home > Add-ins". Click "Advanced" at the bottom left,
   and select the add-in.

### Setup Zotero API Key

1. Obtain your Zotero API key from your Zotero account settings.
2. Open the add-in in PowerPoint and navigate to the settings panel.
3. Enter your API key and User ID and save the settings.

### Insert a Citation

1. Open the add-in **Citation Pane** in PowerPoint.
2. Search for the desired reference using the search bar.
3. Click on the reference to insert it into the slide.
4. Re-order or remove citations on the current slide as needed.

In-slide citations are displayed in a text box named "Citations" (configurable
in the settings). If none is present, it will check the layout master for a text
box named "Citations" and use its properties. Otherwise, a text box is created
in the bottom-left corner of the slide.

### Slide Citation Format

In-slide citations use a format specified in the settings panel in JSON format. Here are some examples:

```json
{
  "apa": {
    "format": "{creator} ({year}). {title}. <i>{journal}</i>, {volume}({issue}), {pages}.",
    "delimiter": ";  ",
  },
  "ieee": {
    "format": "{creator}, \"{title},\" <i>{journal}</i>, vol. {volume}, no. {issue}, pp. {pages}, {year}.",
    "delimiter": ";  ",
  },
  "myCustomFormat1": {
    "format": "<b>[{#}] {creator}</b>, {year}, <i>{journalAbbreviation}</i>",
    "delimiter": "; ",
  }
}
```

- Bold and italics are supported using `<b>` and `<i>` tags, respectively.
- The "delimiter" field specifies how multiple citations are separated.

Supported placeholders include:

- `{creator}`: Author(s) of the work. For one author, displays "Last", for two
authors, "Last1 and Last2", for three or more, "Last1 et al."
- `{#}`: Citation number on slide, starts at 1 on each slide.
- `{title}`: Title of the work.
- `{year}`: Year of publication.
- `{date}`: Date (if applicable)
- `{journalAbbreviation}`: Abbreviated journal name based on PubMed database (if applicable). Defaults to `{publicationTitle}` if not available.
- `{publicationTitle}`: Title of the publication (e.g., journal name).
- `{key}`: Citation key (internal use Zotero Item ID)
- `{volume}`: Volume number (if applicable)
- `{issue}`: Issue number (if applicable)
- `{pages}`: Page range (if applicable)
- `{publisher}`: Publisher (if applicable)
- `{itemType}`: Item type (e.g., book, article)
- `{abstractNote}`: Abstract or summary of the work
- `{DOI}`: DOI (if applicable)
- `{ISBN}`: ISBN (if applicable)
- `{URL}`: URL (if applicable)
- `{accessDate}`: Access date (if applicable)
- `{archive}`: Archive (if applicable)
- `{archiveLocation}`: Archive location (if applicable)
- `{libraryCatalog}`: Library catalog (if applicable)
- `{callNumber}`: Call number (if applicable)
- `{rights}`: Rights information (if applicable)

## Development Setup

### Prerequisites

- **Node.js**: Latest LTS version from [nodejs.org](https://nodejs.org/)
- **Office 365**: PowerPoint with Office.js support
- **Zotero**: Install the latest version of [Zotero](https://www.zotero.org/download/)
- **Visual Studio Code** with the [Office Add-ins Development Kit](https://marketplace.visualstudio.com/items?itemName=OfficeDev.Office-Addin-Dev-Kit) extension

### 1. Install Dependencies

Open the project in VS Code and run:

```bash
npm install
```

### 2. Debug in PowerPoint

Use the Office Add-ins Development Kit extension in VS Code:

1. Open the **Office Add-ins Development Kit** in the Activity Bar
2. Select **Preview Your Office Add-in (F5)**
3. Choose **PowerPoint Desktop (Edge Chromium)**
4. The extension will launch PowerPoint and sideload the add-in

Alternatively, use the VS Code tasks:

- **Debug: PowerPoint Desktop** - Start debugging session
- **Stop Debug** - Stop the debugging session

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Ensure all tests pass
6. Submit a pull request

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
