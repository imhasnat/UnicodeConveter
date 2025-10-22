/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").classList.remove("app-body-hidden");
    
    // Set up event handlers
    document.getElementById("run").onclick = convertSelectionFromPane;
    document.getElementById("font-info").onclick = getFontInfoOnly;
    
    // Auto-convert on input and paste
    document.getElementById("input-text").addEventListener("input", autoConvertText);
    document.getElementById("input-text").addEventListener("paste", handlePaste); 
  }
});

export async function convertSelectionFromPane() {
  return Word.run(async (context) => {
    const trimTrailingBreaks = (s) => (s || "").replace(/[\r\n\u000B\u2028\u2029]+$/g, "");
    
    // Check if ConvertToUnicode function is available
    if (typeof ConvertToUnicode === 'undefined') {
      console.error("ConvertToUnicode function is not available");
      return;
    }
    
    // Get ignore list
    const ignoreList = getIgnoreList();
    
    // Show progress for large datasets
    const startTime = Date.now();
    console.log("[Pane] Starting conversion with ignore list:", ignoreList);
    
    const selection = context.document.getSelection();   
    selection.load("text, font");
    await context.sync();
    
    const text = selection.text || "";
    if (!text.trim()) return;
    
    console.log("[Pane] Selection text:", text);
    console.log("[Pane] Selection font:", selection.font);
    
    // Get word-level font information
    await getWordFontInfo(context, selection);
    
    // Try to get text ranges for better granular control
    const delimiters = ["\r", "\n", "\v", "\u000B", "\u2028", "\u2029"];  
    const textRanges = selection.getTextRanges(delimiters, true);
    context.load(textRanges, "items");
    await context.sync();
    
    // If we have multiple ranges, process them individually
    if (textRanges.items.length > 1) {
      textRanges.items.forEach((tr) => tr.load("text"));
      await context.sync();
      
      for (let i = textRanges.items.length - 1; i >= 0; i--) {
        const tr = textRanges.items[i];
        const txt = tr.text || "";
        if (!txt) continue;
        
        console.log(`[Pane] Sub-range #${i} text:`, txt); 
        const converted = await convertTextWithIgnoreList(txt, ignoreList);
        tr.insertText(trimTrailingBreaks(converted), Word.InsertLocation.replace);
        await context.sync();
      }
    } else {
      // Single range or no ranges - convert the entire selection
      console.log("[Pane] Converting entire selection");
      const converted = await convertTextWithIgnoreList(text, ignoreList);
      selection.insertText(trimTrailingBreaks(converted), Word.InsertLocation.replace);
      await context.sync();
    }
    
    // Log completion time
    const endTime = Date.now();
    console.log(`[Pane] Conversion completed in ${endTime - startTime}ms`);
  });
}

// Function for auto-conversion as user types
function autoConvertText() {
  const inputText = document.getElementById("input-text").value;
  
  if (!inputText.trim()) {
    document.getElementById("output-text").value = "Converted text will appear here...";
    return;
  }
  
  try {
    // Check if ConvertToUnicode function is available
    if (typeof ConvertToUnicode === 'undefined') {
      console.error("ConvertToUnicode function is not available");
      document.getElementById("output-text").value = "ConvertToUnicode function not loaded";
      return;
    }
    
    const converted = convertMultiLineText(inputText);
    document.getElementById("output-text").value = converted;
  } catch (error) {
    console.error("Error in auto-conversion:", error);
    document.getElementById("output-text").value = "Error converting text: " + error.message;
  }
}

// Function to handle paste events
function handlePaste(event) {
  // Allow the paste to complete first, then convert
  setTimeout(() => {
    autoConvertText();
  }, 10);
} 

// Function to convert multi-line text line by line
function convertMultiLineText(text) {
  const lines = text.split('\n');
  const convertedLines = lines.map(line => {
    if (line.trim() === '') {
      return ''; // Preserve empty lines
    }
    
    try {
      return ConvertToUnicode("bijoy", line);
    } catch (error) {
      console.error("Error converting line:", line, error);
      return line; // Return original line if conversion fails
    }
  });
  return convertedLines.join('\n');
}

// Function to get ignore list from input
function getIgnoreList() {
  const ignoreText = document.getElementById("ignore-list").value;
  if (!ignoreText.trim()) {
    return [];
  }
  
  // Split by both commas, spaces, and newlines, then filter out empty strings
  const words = ignoreText.split(/[,\s\n\r]+/)
    .map(word => word.trim())
    .filter(word => word.length > 0);
  
  console.log("Ignore list words:", words);
  return words;
}

// Function to check if a word should be ignored (simplified)
function shouldIgnoreWord(word, ignoreList) {
  if (!word || !ignoreList || ignoreList.length === 0) {
    return false;
  }
  
  const cleanWord = word.trim();
  
  // Check for exact match (most common case)
  if (ignoreList.includes(cleanWord)) {
    return true;
  }
  
  // Check for case-insensitive match
  const lowerWord = cleanWord.toLowerCase();
  for (const ignoreWord of ignoreList) {
    if (ignoreWord.toLowerCase() === lowerWord) {
      return true;
    }
  }
  
  return false;
}

// Function to convert text with ignore list (optimized for large datasets)
async function convertTextWithIgnoreList(text, ignoreList) {
  if (!text.trim()) {
    return text;
  }
  
  // Early return if no ignore list
  if (!ignoreList || ignoreList.length === 0) {
    return ConvertToUnicode("bijoy", text);
  }
  
  // Split text into words while preserving spaces
  const words = text.split(/(\s+)/);
  const convertedWords = new Array(words.length);
  
  // Process words in batches to avoid blocking the UI
  const batchSize = 50; // Process 50 words at a time
  
  for (let i = 0; i < words.length; i += batchSize) {
    const batch = words.slice(i, i + batchSize);
    
    for (let j = 0; j < batch.length; j++) {
      const wordIndex = i + j;
      const word = batch[j];
      
      // Skip conversion if word is only whitespace
      if (/^\s+$/.test(word)) {
        convertedWords[wordIndex] = word;
        continue;
      }
      
      // Skip conversion for parentheses characters
      if (word === "(" || word === ")") {
        console.log(`[Convert] Skipping conversion for parentheses character: "${word}"`);
        convertedWords[wordIndex] = word;
        continue;
      }
      
      // Use optimized matching logic to check if word should be ignored
      if (shouldIgnoreWord(word, ignoreList)) {
        convertedWords[wordIndex] = word; // Return original word without conversion
        continue;
      }
      
      // Convert the word
      convertedWords[wordIndex] = ConvertToUnicode("bijoy", word);
    }
    
    // Yield control back to the browser for large datasets
    if (words.length > 100 && i + batchSize < words.length) {
      // Use setTimeout to yield control (only for large datasets)
      await new Promise(resolve => setTimeout(resolve, 0));
    }
  }
  
  return convertedWords.join('');
}

// Optimized function to get word-level font information and convert SutonnyMJ words
async function getWordFontInfo(context, selection) {
  try {
    console.log("[Font Info] Starting word-level font analysis...");
    
    // Get text ranges for each word
    const wordDelimiters = [" ",",", "\t", "\r", "\n", "\v", "\u000B", "\u2028", "\u2029", "(", ")"];
    const wordRanges = selection.getTextRanges(wordDelimiters, true);
    
    // First load the items collection
    context.load(wordRanges, "items");
    await context.sync();
    
    // Then load text and font properties for each range
    wordRanges.items.forEach(range => {
      range.load("text, font");
    });
    await context.sync();
    
    console.log(`[Font Info] Found ${wordRanges.items.length} word ranges`);
    
    // Get ignore list for conversion
    const ignoreList = getIgnoreList();
    
    // Process each word range and log font information
    for (let i = wordRanges.items.length - 1; i >= 0; i--) {
      const range = wordRanges.items[i];
      const word = range.text ? range.text.trim() : "";
      
      if (word) {
        const fontName = range.font.name || "Unknown";
        const fontSize = range.font.size || "Unknown";
        const fontColor = range.font.color || "Unknown";
        
        console.log(`[Font Info] Word ${i + 1}: "${word}" | Font: ${fontName} | Size: ${fontSize} | Color: ${fontColor}`);
        
        // Check if font is SutonnyMJ and convert to Unicode
        if (fontName.toLowerCase().includes("sutonnymj")) {
          console.log(`[Font Info] Converting SutonnyMJ word: "${word}"`);
          
          // Skip conversion for parentheses characters
          if (word === "(" || word === ")") {
            console.log(`[Font Info] Skipping conversion for parentheses character: "${word}"`);
            continue;
          }
          
          // Check if word should be ignored
          if (!shouldIgnoreWord(word, ignoreList)) {
            try {
              // Check if ConvertToUnicode function is available
              if (typeof ConvertToUnicode !== 'undefined') {
                const convertedWord = ConvertToUnicode("bijoy", word);
                console.log(`[Font Info] Converted "${word}" to "${convertedWord}"`);
                
                // Replace the word with converted version
                range.insertText(convertedWord, Word.InsertLocation.replace);
                await context.sync();
              } else {
                console.error("[Font Info] ConvertToUnicode function is not available");
              }
            } catch (conversionError) {
              console.error(`[Font Info] Error converting word "${word}":`, conversionError);
            }
          } else {
            console.log(`[Font Info] Skipping conversion for ignored word: "${word}"`);
          }
        }
      }
    }
    
    console.log("[Font Info] Word-level font analysis and conversion completed");
    
  } catch (error) {
    console.error("[Font Info] Error getting word font information:", error);
  }
}


// Standalone function to get font information and convert SutonnyMJ words
export async function getFontInfoOnly() {
  return Word.run(async (context) => {
    try {
      console.log("[Font Info Only] Starting font analysis and conversion...");
      
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();
      
      const text = selection.text || "";
      if (!text.trim()) {
        console.log("[Font Info Only] No text selected");
        return;
      }
      
      console.log("[Font Info Only] Selected text:", text);
      
      // Get word-level font information and convert SutonnyMJ words
      await getWordFontInfo(context, selection);
      
    } catch (error) {
      console.error("[Font Info Only] Error:", error);
    }
  });
}
