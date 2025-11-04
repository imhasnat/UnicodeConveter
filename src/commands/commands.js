/*
 * Copyright (c) Hasnat Mahbub. All rights reserved. Licensed under the MIT License.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

async function convertSelectionImpl() {
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    // Log entire selection text for debugging
    selection.load("text");
    await context.sync();
    // Prefer splitting selection into text ranges to support non-contiguous/column selections
    const delimiters = ["\r", "\n", "\v", "\u000B", "\u2028", "\u2029"]; // CR, LF, VT, Unicode LS/PS (no TAB)
    const textRanges = selection.getTextRanges(delimiters, true);
    context.load(textRanges, "items");
    await context.sync();

    // If it is a single range (e.g., vertical selection treated as one block),
    // convert the entire selection at once.
    if (textRanges.items.length === 1) {
      const full = selection.text || "";
      if (full) {
        console.warn("[Cmd] Single-range fallback: converting entire selection");
        // eslint-disable-next-line no-undef
        const convertedFull = ConvertToUnicode("bijoy", full);
        selection.insertText(convertedFull, Word.InsertLocation.replace);
        await context.sync();
        return;
      }
    }

    if (textRanges.items.length === 0) {
      const paragraphs = selection.paragraphs;
      context.load(paragraphs, "items");
      await context.sync();
      if (paragraphs.items.length === 0) {
        selection.load("text");
        await context.sync();
        const text = selection.text || "";
        if (!text) return;
        // eslint-disable-next-line no-undef
        const converted = ConvertToUnicode("bijoy", text);
        selection.insertText(converted, Word.InsertLocation.replace);
        await context.sync();
        return;
      }
      const paraRanges = paragraphs.items.map((p) => p.getRange());
      paraRanges.forEach((r) => r.load("text"));
      await context.sync();
      for (let i = paraRanges.length - 1; i >= 0; i--) {
        const r = paraRanges[i];
        const txt = r.text || "";
        if (!txt) continue;
        // eslint-disable-next-line no-undef
        const converted = ConvertToUnicode("bijoy", txt);
        r.insertText(converted, Word.InsertLocation.replace);
        await context.sync();
      }
      return;
    }

    textRanges.items.forEach((tr) => tr.load("text"));
    await context.sync();
    if (textRanges.items.length > 1) {
      for (let i = textRanges.items.length - 1; i >= 0; i--) {
        const tr = textRanges.items[i];
        const txt = tr.text || "";
        if (!txt) continue;
        // eslint-disable-next-line no-undef
        const converted = ConvertToUnicode("bijoy", txt);
        tr.insertText(converted, Word.InsertLocation.replace);
        await context.sync();
      }
      return;
    }

    // Last-resort fallback: split by spaces if still a single range
    const wordRanges = selection.getTextRanges([" "], true);
    context.load(wordRanges, "items");
    await context.sync();
    console.warn("[Cmd] Fallback to space-delimited ranges:", wordRanges.items.length);
    wordRanges.items.forEach((wr) => wr.load("text"));
    await context.sync();
    for (let i = wordRanges.items.length - 1; i >= 0; i--) {
      const wr = wordRanges.items[i];
      const wtxt = wr.text || "";
      if (!wtxt) continue;
      // eslint-disable-next-line no-undef
      const converted = ConvertToUnicode("bijoy", wtxt);
      wr.insertText(converted, Word.InsertLocation.replace);
      await context.sync();
    }
  });
}

function convertSelection(event) {
  convertSelectionImpl()
    .catch(() => {})
    .finally(() => {
      if (event && typeof event.completed === "function") {
        event.completed();
      }
    });
}

Office.actions.associate("convertSelection", convertSelection);
