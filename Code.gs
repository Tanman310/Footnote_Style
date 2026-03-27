/**
 * Footnote_Style — Google Apps Script
 *
 * Formats every footnote in the active Google Doc so that all footnotes
 * share a consistent appearance:
 *   • Font family  : Times New Roman
 *   • Font size    : 10 pt
 *   • Bold / Italic: cleared (plain text)
 *   • Alignment    : LEFT
 *   • Punctuation  : trailing period ensured, multiple spaces collapsed,
 *                    leading/trailing whitespace trimmed
 *
 * To add a custom menu entry, the onOpen() trigger is included so users
 * can run the formatter directly from the Docs UI without opening the
 * Apps Script editor.
 */

// ---------------------------------------------------------------------------
// Configuration — change these values to match your house style
// ---------------------------------------------------------------------------
var CONFIG = {
  fontFamily: 'Times New Roman',
  fontSize: 10,          // points
  bold: false,
  italic: false,
  underline: false,
  alignment: DocumentApp.HorizontalAlignment.LEFT,
  ensureTrailingPeriod: true,   // add a period if the footnote doesn't end with one
  collapseSpaces: true,         // replace runs of spaces with a single space
  trimWhitespace: true          // remove leading/trailing whitespace from footnote text
};

// ---------------------------------------------------------------------------
// Menu
// ---------------------------------------------------------------------------

/**
 * Creates a custom "Footnotes" menu in the Google Docs UI when the document
 * is opened.
 */
function onOpen() {
  DocumentApp.getUi()
    .createMenu('Footnotes')
    .addItem('Format all footnotes', 'formatAllFootnotes')
    .addToUi();
}

// ---------------------------------------------------------------------------
// Main entry point
// ---------------------------------------------------------------------------

/**
 * Iterates over every footnote in the active document and applies the
 * formatting defined in CONFIG to each paragraph inside each footnote.
 *
 * A summary toast is displayed when the operation completes.
 */
function formatAllFootnotes() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var footnotes = getFootnotes(body);

  if (footnotes.length === 0) {
    DocumentApp.getUi().alert('No footnotes found in this document.');
    return;
  }

  footnotes.forEach(function (footnote) {
    formatFootnote(footnote);
  });

  DocumentApp.getUi().alert('Done! Formatted ' + footnotes.length + ' footnote(s).');
}

// ---------------------------------------------------------------------------
// Footnote helpers
// ---------------------------------------------------------------------------

/**
 * Returns all Footnote elements found within a Body element.
 *
 * Google Docs stores footnotes as inline Footnote elements inside
 * paragraph text runs. We walk the entire body to collect them.
 *
 * @param {GoogleAppsScript.Document.Body} body
 * @returns {GoogleAppsScript.Document.Footnote[]}
 */
function getFootnotes(body) {
  var footnotes = [];
  collectFootnotes(body, footnotes);
  return footnotes;
}

/**
 * Recursively walks an element tree and pushes any Footnote nodes into
 * the provided array.
 *
 * @param {GoogleAppsScript.Document.Element} element
 * @param {GoogleAppsScript.Document.Footnote[]} result
 */
function collectFootnotes(element, result) {
  var type = element.getType();

  if (type === DocumentApp.ElementType.FOOTNOTE) {
    result.push(element.asFootnote());
    return; // footnotes don't nest further
  }

  // Walk children if the element supports it
  if (typeof element.getNumChildren === 'function') {
    var numChildren = element.getNumChildren();
    for (var i = 0; i < numChildren; i++) {
      collectFootnotes(element.getChild(i), result);
    }
  }
}

/**
 * Applies CONFIG formatting to every paragraph inside a single footnote.
 *
 * @param {GoogleAppsScript.Document.Footnote} footnote
 */
function formatFootnote(footnote) {
  var contents = footnote.getFootnoteContents();
  var numParagraphs = contents.getNumChildren();

  for (var p = 0; p < numParagraphs; p++) {
    var child = contents.getChild(p);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) {
      continue;
    }

    var paragraph = child.asParagraph();
    paragraph.setAlignment(CONFIG.alignment);

    // Format each text element inside the paragraph
    var numChildren = paragraph.getNumChildren();
    for (var c = 0; c < numChildren; c++) {
      var inline = paragraph.getChild(c);
      if (inline.getType() === DocumentApp.ElementType.TEXT) {
        formatTextElement(inline.asText());
      }
    }

    // Punctuation / whitespace cleanup on the paragraph's plain text
    if (CONFIG.trimWhitespace || CONFIG.collapseSpaces || CONFIG.ensureTrailingPeriod) {
      cleanParagraphText(paragraph);
    }
  }
}

/**
 * Applies font and style attributes to a Text element.
 *
 * @param {GoogleAppsScript.Document.Text} text
 */
function formatTextElement(text) {
  var content = text.getText();
  if (content.length === 0) {
    return;
  }

  text.setFontFamily(CONFIG.fontFamily);
  text.setFontSize(CONFIG.fontSize);
  text.setBold(CONFIG.bold);
  text.setItalic(CONFIG.italic);
  text.setUnderline(CONFIG.underline);
}

/**
 * Performs whitespace and punctuation cleanup on a paragraph's raw text
 * by replacing the entire paragraph text content while preserving structure.
 *
 * Because replacing individual Text elements' content is the safest approach
 * (avoids fighting with mixed-format runs), we only touch the last Text child
 * for the trailing-period rule and use a simple regex pass for spaces.
 *
 * @param {GoogleAppsScript.Document.Paragraph} paragraph
 */
function cleanParagraphText(paragraph) {
  var numChildren = paragraph.getNumChildren();

  for (var i = 0; i < numChildren; i++) {
    var child = paragraph.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.TEXT) {
      continue;
    }

    var textEl = child.asText();
    var raw = textEl.getText();

    if (CONFIG.collapseSpaces) {
      raw = raw.replace(/ {2,}/g, ' ');
    }

    // Trim leading whitespace on the first text child
    if (i === 0 && CONFIG.trimWhitespace) {
      raw = raw.replace(/^\s+/, '');
    }

    // Trim trailing whitespace on the last text child
    var isLast = (i === numChildren - 1);
    if (isLast && CONFIG.trimWhitespace) {
      raw = raw.replace(/\s+$/, '');
    }

    // Ensure trailing period on the very last text child
    if (isLast && CONFIG.ensureTrailingPeriod && raw.length > 0) {
      var lastChar = raw.charAt(raw.length - 1);
      if (lastChar !== '.' && lastChar !== '?' && lastChar !== '!') {
        raw = raw + '.';
      }
    }

    textEl.setText(raw);
  }
}
