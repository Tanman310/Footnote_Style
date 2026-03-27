# Footnote_Style

A Google Apps Script that automatically formats every footnote in a Google Doc to match a consistent house style.

---

## Features

| Setting | Default |
|---|---|
| Font family | Times New Roman |
| Font size | 10 pt |
| Bold / Italic / Underline | Removed (plain text) |
| Alignment | Left |
| Trailing period | Added if missing |
| Extra spaces collapsed | Yes |
| Leading / trailing whitespace | Trimmed |

All defaults are configurable via the `CONFIG` object at the top of `Code.gs`.

---

## Setup

### Option A — Copy & paste into the Apps Script editor

1. Open the Google Doc you want to use the script with.
2. Go to **Extensions → Apps Script**.
3. Delete any existing code in `Code.gs` and paste the contents of `Code.gs` from this repository.
4. Click **Save** (💾).
5. Reload the Google Doc — a **Footnotes** menu will appear in the menu bar.

### Option B — Deploy with `clasp` (command-line)

[`clasp`](https://github.com/google/clasp) lets you manage Apps Script projects from the terminal.

```bash
# Install clasp globally
npm install -g @google/clasp

# Log in with your Google account
clasp login

# Clone an existing Apps Script project (get the Script ID from
# Extensions → Apps Script → Project Settings → Script ID)
clasp clone <scriptId>

# — or — create a new standalone script
clasp create --title "Footnote_Style" --type docs

# Push local files to Apps Script
clasp push
```

After pushing, open the bound document and reload to see the **Footnotes** menu.

---

## Usage

1. Open a Google Doc that contains footnotes.
2. Click **Footnotes → Format all footnotes** in the menu bar.
3. Grant the requested permissions on first run (the script only needs access to the current document).
4. A confirmation dialog shows how many footnotes were formatted.

---

## Configuration

Open `Code.gs` and edit the `CONFIG` object near the top of the file:

```javascript
var CONFIG = {
  fontFamily: 'Times New Roman',  // any Google Fonts name
  fontSize: 10,                   // points
  bold: false,
  italic: false,
  underline: false,
  alignment: DocumentApp.HorizontalAlignment.LEFT,
  ensureTrailingPeriod: true,     // set to false to skip punctuation fix
  collapseSpaces: true,           // set to false to keep extra spaces
  trimWhitespace: true            // set to false to keep leading/trailing spaces
};
```

---

## File Structure

```
Footnote_Style/
├── appsscript.json   # Apps Script project manifest
├── Code.gs           # Main script — footnote formatter
├── LICENSE
└── README.md
```

---

## License

[MIT](LICENSE)
