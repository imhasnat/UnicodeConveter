/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").classList.remove("app-body-hidden");
     
    document.getElementById("font-info").onclick = getFontInfoOnly;
     
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
    
    // Show loading spinner
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
        const convertedText = ConvertToUnicode("bijoy", selectedText);
        // console.log(`[Font Info] Converted: "${selectedText}" -> "${convertedText}"`); 
        await selectionRange.insertText(convertedText, Word.InsertLocation.replace); 
        updateLoadingProgress("Converting text...", "Selection converted successfully");
        
      } catch (conversionError) {
        console.error(`[Font Info] Error converting selection:`, conversionError);
        updateLoadingProgress("Error", "Conversion failed: " + conversionError.message);
      }
    } else { 
      // console.log(`[Font Info] Selection not SutonnyMJ, processing word-by-word within selection`);
      await processWordsWithinSelection(context, selection);
    }

    // Final sync to flush all remaining edits
    await context.sync(); 
    updateLoadingProgress("Finalizing conversion...", "Conversion completed"); 
    await new Promise(resolve => setTimeout(resolve, 500)); 
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
    const wordDelimiters = [" ",",", "\t", "\r", "\n", "\v", "\u000B", "\u2028", "\u2029", "(", ")"];
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
 
      if (word && fontName.toLowerCase().includes("sutonnymj") && word !== "(" && word !== ")") {
        try {
          // console.log(`[Font Info] Converting word: "${word}"`);
          const convertedWord = ConvertToUnicode("bijoy", word);
          // console.log(`[Font Info] Converted: "${word}" -> "${convertedWord}"`);
          
          await range.insertText(convertedWord, Word.InsertLocation.replace);
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
export async function getFontInfoOnly() {
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
      
      // Get word-level font information and convert SutonnyMJ words
      await getWordFontInfo(context, selection);
      
    } catch (error) {
      console.error("[Font Info Only] Error:", error);
    }
  });
}

