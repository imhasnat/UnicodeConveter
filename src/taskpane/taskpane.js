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
    selection.load("text");
    await context.sync(); 
    
    const tableColumns = context.document.sections.items;   
    tableColumns.load("items, text");
    await context.sync();
    
    console.log("tableColumns", tableColumns);

    const text = selection.text || "";
    if (!text.trim()) return;
    
    console.log("[Pane] Selection text:", text);
    
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

// Functional Trie implementation
function createTrieNode() {
  return {
    children: {},
    isEndOfWord: false
  };
}

function createTrie() {
  return {
    root: createTrieNode()
  };
}

function insertWord(trie, word) {
  let node = trie.root;
  for (let i = 0; i < word.length; i++) {
    const char = word[i];
    if (!node.children[char]) {
      node.children[char] = createTrieNode();
    }
    node = node.children[char];
  }
  node.isEndOfWord = true;
}

function searchWord(trie, word) {
  let node = trie.root;
  for (let i = 0; i < word.length; i++) {
    const char = word[i];
    if (!node.children[char]) {
      return false;
    }
    node = node.children[char];
  }
  return node.isEndOfWord;
}

function hasPrefix(trie, prefix) {
  let node = trie.root;
  for (let i = 0; i < prefix.length; i++) {
    const char = prefix[i];
    if (!node.children[char]) {
      return false;
    }
    node = node.children[char];
  }
  return true;
}

// Optimized data structures for exact and prefix matching
let ignoreWordSet = new Set(); // O(1) exact matches
let ignoreWordTrie = createTrie(); // O(m) prefix matching
let ignoreListHash = null; // Track ignore list changes

// Preprocess ignore list into optimized data structures
function preprocessIgnoreList(ignoreList) {
  if (!ignoreList || ignoreList.length === 0) {
    ignoreWordSet.clear();
    ignoreWordTrie = createTrie();
    return;
  }
  
  // Create hash to detect changes
  const newHash = ignoreList.join('|');
  if (ignoreListHash === newHash) {
    return; // No changes, keep existing structures
  }
  
  ignoreListHash = newHash;
  ignoreWordSet.clear();
  ignoreWordTrie = createTrie();
  
  // Build optimized data structures - preserve all Bijoy characters
  for (const word of ignoreList) {
    const cleanWord = word.trim();
    if (cleanWord.length === 0) continue;
    
    // Add exact match (preserves all Bijoy characters)
    ignoreWordSet.add(cleanWord);
    insertWord(ignoreWordTrie, cleanWord);
  }
}

// Optimized word matching - O(1) exact + O(m) prefix matching
function shouldIgnoreWord(word, ignoreList) {
  if (!word || !ignoreList || ignoreList.length === 0) {
    return false;
  }
  
  const cleanWord = word.trim();
  if (cleanWord.length === 0) return false;
  
  // O(1) exact match check (preserves all Bijoy characters)
  if (ignoreWordSet.has(cleanWord)) {
    return true;
  }
  
  // O(m) prefix match check (for cases like "test" matching "testing")
  if (hasPrefix(ignoreWordTrie, cleanWord)) {
    return true;
  }
  
  return false;
}

// Optimized conversion using single-pass algorithm
async function convertTextWithIgnoreList(text, ignoreList) {
  if (!text.trim()) {
    return text;
  }
  
  // Early return if no ignore list
  if (!ignoreList || ignoreList.length === 0) {
    return ConvertToUnicode("bijoy", text);
  }
  
  // Preprocess ignore list into optimized data structures (O(n) once)
  preprocessIgnoreList(ignoreList);
  
  // Single-pass processing with optimized word detection
  const result = [];
  let currentWord = '';
  let inWord = false;
  let processedWords = 0;
  
  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    
    // Check if character is word boundary
    if (/\s/.test(char)) {
      if (inWord && currentWord.length > 0) {
        // Process complete word
        if (shouldIgnoreWord(currentWord, ignoreList)) {
          result.push(currentWord);
        } else {
          result.push(ConvertToUnicode("bijoy", currentWord));
          processedWords++;
        }
        currentWord = '';
        inWord = false;
      }
      result.push(char); // Add whitespace
    } else {
      // Building word
      currentWord += char;
      inWord = true;
    }
    
    // Yield control for very large texts
    if (processedWords > 0 && processedWords % 1000 === 0) {
      await new Promise(resolve => setTimeout(resolve, 0));
    }
  }
  
  // Process final word if exists
  if (inWord && currentWord.length > 0) {
    if (shouldIgnoreWord(currentWord, ignoreList)) {
      result.push(currentWord);
    } else {
      result.push(ConvertToUnicode("bijoy", currentWord));
    }
  }
  
  return result.join('');
}
