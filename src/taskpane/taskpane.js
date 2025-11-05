/*
 * Copyright (c) Hasnat Mahbub. All rights reserved. Licensed under the MIT License.
 * See LICENSE in the project root for license information.
 * Task pane JavaScript for Bijoy to Unicode Converter Word Add-in
 */

/* global document, Office, Word */ 

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").classList.remove("app-body-hidden");
     
    document.getElementById("convert-sutonnymj").onclick = convertSelection;
    // document.getElementById("apply-arabic-font-selection").onclick = applyArabicFontToSelection;
     
    document.getElementById("input-text").addEventListener("input", autoConvertText);
    document.getElementById("input-text").addEventListener("paste", handlePaste); 
  }
});

// Function for auto-conversion as user types
function autoConvertText() {
  const inputText = document.getElementById("input-text").value;
  
  if (!inputText.trim()) {
    document.getElementById("output-text").value = "Converted text will appear here...";
    return;
  }
  
  try {  
    showLoadingSpinner("Converting text...", "Processing input text");
     
    setTimeout(() => {
      try {
        const converted = convertMultiLineText(inputText);
        document.getElementById("output-text").value = converted;
        hideLoadingSpinner();
      } catch (error) {
        console.error("Error in auto-conversion:", error);
        document.getElementById("output-text").value = "Error converting text: " + error.message;
        hideLoadingSpinner();
      }
    }, 100);
    
  } catch (error) {
    console.error("Error in auto-conversion:", error);
    document.getElementById("output-text").value = "Error converting text: " + error.message;
    hideLoadingSpinner();
  }
}

// Function to handle paste events
function handlePaste(event) { 
  setTimeout(() => {
    autoConvertText();
  }, 10);
} 

// Function to detect if text contains Arabic characters
function containsArabic(text) { 
  const arabicPattern = /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]/;
  return arabicPattern.test(text);
}

// Function to convert multi-line text line by line
function convertMultiLineText(text) {
  const lines = text.split('\n');
  const convertedLines = lines.map(line => {
    if (line.trim() === '') {
      return '';
    }
    
    try {
      return ConvertToUnicode("bijoy", line);
    } catch (error) {
      console.error("Error converting line:", line, error);
      return line;
    }
  });
  return convertedLines.join('\n');
}

// Function to show loading spinner
function showLoadingSpinner(message = "Converting text...", progressInfo = "") {
  const overlay = document.getElementById("loading-overlay");
  const loadingText = document.getElementById("loading-text");
  const progressInfoElement = document.getElementById("progress-info");
  
  loadingText.textContent = message;
  progressInfoElement.textContent = progressInfo;
  overlay.className = "loading-overlay";
}

// Function to hide loading spinner
function hideLoadingSpinner() {
  const overlay = document.getElementById("loading-overlay");
  overlay.className = "loading-overlay-hidden";
}

// Function to update loading progress
function updateLoadingProgress(message, progressInfo = "") {
  const loadingText = document.getElementById("loading-text");
  const progressInfoElement = document.getElementById("progress-info");
  
  loadingText.textContent = message;
  progressInfoElement.textContent = progressInfo;
}

// Optimized function to get word-level font information and convert SutonnyMJ words
async function getWordFontInfo(context, selection) {
  try {
    // console.log("[Font Info] Starting word-level font analysis..."); 
    showLoadingSpinner("Analyzing document...", "Scanning for MJ fonts");
     
    selection.load("text");
    await context.sync();
    const selectedText = selection.text || "";
    
    if (!selectedText.trim()) {
      // console.log("[Font Info] No text selected");
      hideLoadingSpinner();
      return;
    }
    
    // console.log(`[Font Info] Selected text: "${selectedText}"`);
     
    const selectionRange = selection.getRange();
    selectionRange.load("text, font");
    await context.sync(); 
    const fontName = selectionRange.font.name || "Unknown";
    // console.log(`[Font Info] Selection font: "${fontName}"`);
    
    if (fontName.toLowerCase().includes("sutonnymj")) {  
      try {
        // const originalSize = selectionRange.font.size;
        // const convertedText = ConvertToUnicode("bijoy", selectedText);
        // // console.log(`[Font Info] Converted: "${selectedText}" -> "${convertedText}"`); 
        // await selectionRange.insertText(convertedText, Word.InsertLocation.replace);
         
        // selectionRange.font.name = "Kalpurush";
        // if (typeof originalSize === "number" && !isNaN(originalSize) && originalSize > 0) {
        //   console.log("getWordFontInfo originalSize", originalSize);
        //   selectionRange.font.size = originalSize;
        // }
        // await context.sync();
        
        // Convert paragraph-by-paragraph to preserve per-line bold/normal formatting
        const paragraphs = selection.paragraphs;
        context.load(paragraphs, "items");
        await context.sync();

        for (let i = 0; i < paragraphs.items.length; i++) {
          const p = paragraphs.items[i];
          // Use paragraph content only (exclude paragraph mark) to avoid adding extra blank lines
          const pRange = p.getRange(Word.RangeLocation.content);
          pRange.load("text, font");
        }
        await context.sync();

        let convertedParaCount = 0;
        for (let i = 0; i < paragraphs.items.length; i++) {
          const pRange = paragraphs.items[i].getRange(Word.RangeLocation.content);
          pRange.load("text, font");
          await context.sync();

          const pText = pRange.text || "";
          if (!pText.trim()) {
            continue;
          }

          const pFontName = (pRange.font.name || "").toLowerCase();
          if (!pFontName.includes("sutonnymj")) {
            continue; // leave non-SutonnyMJ paragraphs unchanged
          }

          const originalSize = pRange.font.size;
          const converted = ConvertToUnicode("bijoy", pText);
          await pRange.insertText(converted, Word.InsertLocation.replace);

          // Apply target Unicode font but don't touch bold/italic so existing styling remains
          pRange.font.name = "Kalpurush";
          if (typeof originalSize === "number" && !isNaN(originalSize) && originalSize > 0) {
            pRange.font.size = originalSize;
          }

          convertedParaCount++;
          if (convertedParaCount % 5 === 0) {
            await context.sync();
          }
        }

        await context.sync();
        updateLoadingProgress("Converting text...", "Selection converted successfully");
        
      } catch (conversionError) {
        console.error(`[Font Info] Error converting selection:`, conversionError);
        updateLoadingProgress("Error", "Conversion failed: " + conversionError.message);
      }
    } else { 
      // console.log(`[Font Info] Selection not SutonnyMJ, processing word-by-word within selection`);
      await processWordsWithinSelection(context, selection);
    }
 
    await context.sync(); 
    updateLoadingProgress("Finalizing conversion...", "Conversion completed"); 
    hideLoadingSpinner();
    
    // console.log("[Font Info] Word-level font analysis and conversion completed");
    
  } catch (error) {
    console.error("[Font Info] Error getting word font information:", error);
    hideLoadingSpinner();
  }
}

// Process words within selection only (conservative approach)
async function processWordsWithinSelection(context, selection) {
  try { 
    console.log("[Word Processing] Starting word-by-word processing...");
    const wordDelimiters = [" ",",", "\t", "\r", "\n", "\v", "\u000B", "\u2028", "\u2029", "(", ")", "-", "=", "/"];
    const wordRanges = selection.getTextRanges(wordDelimiters, true);
     
    context.load(wordRanges, "items");
    await context.sync();
     
    wordRanges.items.forEach(range => {
      range.load("text, font");
    });
    await context.sync();
    
    // console.log(`[Font Info] Found ${wordRanges.items.length} word ranges in selection`);
 
    let convertedCount = 0;
    const batchSize = 20;
    let pendingEdits = 0;  

    for (let i = 0; i < wordRanges.items.length; i++) {
      const range = wordRanges.items[i];
      const word = range.text ? range.text.trim() : "";
      const fontName = range.font.name || "Unknown";
 
      if (word && fontName.toLowerCase().includes("sutonnymj") && word !== "(" && word !== ")" && word !== "|") {
        try {
          // console.log(`[Font Info] Converting word: "${word}"`);
          const originalSize = range.font.size;
          const convertedWord = ConvertToUnicode("bijoy", word);
          // console.log(`[Font Info] Converted: "${word}" -> "${convertedWord}"`);
          
          await range.insertText(convertedWord, Word.InsertLocation.replace);
           


          range.font.name = "Kalpurush";
          if (typeof originalSize === "number" && !isNaN(originalSize) && originalSize > 0) {
            console.log("processWordsWithinSelection originalSize", originalSize);
            range.font.size = originalSize;
          }
          
          convertedCount++;
          pendingEdits++; 
          updateLoadingProgress("Converting text...", `Converted ${convertedCount} words`);
           
          if (pendingEdits >= batchSize) {
            await context.sync();
            pendingEdits = 0;
          }
          
        } catch (conversionError) {
          console.error(`[Font Info] Error converting word "${word}":`, conversionError);
        }
      }
    }

    // console.log(`[Font Info] Converted ${convertedCount} words in selection`);
    
  } catch (error) {
    console.error("[Font Info] Error processing words within selection:", error);
  }
}

// Standalone function to get font information and convert SutonnyMJ words
export async function convertSelection() {
  return Word.run(async (context) => {
    try {
      // console.log("[Font Info Only] Starting font analysis and conversion...");
      
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();
      
      const text = selection.text || "";
      if (!text.trim()) {
        //  console.log("[Font Info Only] No text selected");
        return;
      }
      
      // console.log("[Font Info Only] Selected text:", text);
       
      await getWordFontInfo(context, selection);
      
    } catch (error) {
      console.error("[Font Info Only] Error:", error);
    }
  });
}

// Function to apply Al Majeed Quranic Font to Arabic text in selected text (OPTIMIZED - FASTER)
export async function applyArabicFontToSelection() {
  return Word.run(async (context) => {
    try {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();
      
      const selectedText = selection.text || "";
      if (!selectedText.trim()) {
        showLoadingSpinner("No selection", "Please select some text first");
        hideLoadingSpinner();
        return;
      }
       
      if (!containsArabic(selectedText)) {
        showLoadingSpinner("No Arabic text", "No Arabic text found in selection");
        hideLoadingSpinner();
        return;
      }
      
      showLoadingSpinner("Applying Arabic font...", "Processing selected text");
      
      const arabicFontFamily = "Al Majeed Quranic Font";
      const wordDelimiters = [" ", ",", "\t", "\r", "\n", "\v", "\u000B", "\u2028", "\u2029", "(", ")", "-", "=", "/", ".", ";", ":", "!", "?"];
       
      const wordRanges = selection.getTextRanges(wordDelimiters, true);
      context.load(wordRanges, "items");
      await context.sync();
      
      const totalRanges = wordRanges.items.length;
       
      wordRanges.items.forEach(range => {
        range.load("text");
      });
      await context.sync();
       
      const arabicRangeIndices = [];
      for (let i = 0; i < wordRanges.items.length; i++) {
        const text = wordRanges.items[i].text ? wordRanges.items[i].text.trim() : "";
        if (text && containsArabic(text)) {
          arabicRangeIndices.push(i);
        }
      }
      
      const totalArabicRanges = arabicRangeIndices.length;
      
      if (totalArabicRanges === 0) {
        updateLoadingProgress("Complete!", "No Arabic text found in selection");
        hideLoadingSpinner();
        return;
      }
       
      const batchSize = 500;
      let processedCount = 0;
       
      for (let batchStart = 0; batchStart < arabicRangeIndices.length; batchStart += batchSize) {
        const batchEnd = Math.min(batchStart + batchSize, arabicRangeIndices.length);
         
        for (let idx = batchStart; idx < batchEnd; idx++) {
          const rangeIndex = arabicRangeIndices[idx];
          wordRanges.items[rangeIndex].font.name = arabicFontFamily;
          processedCount++;
        }
         
        await context.sync();
         
        updateLoadingProgress("Applying Arabic font...", `Processed ${processedCount}/${totalArabicRanges} Arabic ranges`);
      }
      
      updateLoadingProgress("Complete!", `Applied "${arabicFontFamily}" to ${totalArabicRanges} Arabic text ranges`);
      hideLoadingSpinner();
      
      console.log(`[Apply Arabic Font Selection] Applied "${arabicFontFamily}" to ${totalArabicRanges} Arabic text ranges (from ${totalRanges} total ranges)`);
      
    } catch (error) {
      console.error("[Apply Arabic Font Selection] Error:", error);
      updateLoadingProgress("Error", "Failed to apply Arabic font: " + error.message);
      // Show error briefly before hiding (errors are important)
      await new Promise(resolve => setTimeout(resolve, 1000));
      hideLoadingSpinner();
    }
  });
}
