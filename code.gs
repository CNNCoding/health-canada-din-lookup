/**
 * @OnlyCurrentDoc
 * 
 * This script uses a caching approach because the Health Canada API endpoints
 * don't support proper filtering. We fetch all data once and cache it.
 */

// Cache duration in milliseconds (6 hours)
const CACHE_DURATION = 6 * 60 * 60 * 1000;


/**
 * Clear the cached data, including all chunks.
 */
function clearCache() {
  const cache = CacheService.getScriptCache();
  const keysToRemove = ['cacheTimestamp', 'drugData_chunk_count'];

  // Prepare to remove up to 100 potential chunks
  for (let i = 0; i < 100; i++) {
    keysToRemove.push(`drugData_${i}`);
  }
  
  cache.removeAll(keysToRemove);
  
}

/**
 * Main function to fetch DINs from Health Canada API
 */
function populateDinsFromHealthCanada_Advanced() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Product Pricing');
  
  if (!sheet) {
    ui.alert("Error: Sheet named 'Product Pricing' not found.");
    return;
  }

  const HEADER_MAP = {
    type: 'Type',
    drug: 'Drug',
    ingredients: 'Ingredients',
    strength: 'Strength',
    form: 'Drug Form',
    country: 'Country of Origin',
    din: 'DIN'
  };

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const header = values[0];
  const body = values.slice(1);

  // Find column indices
  const colIdx = {};
  for (const key in HEADER_MAP) {
    const index = header.indexOf(HEADER_MAP[key]);
    if (index === -1 && key !== 'din') {
      ui.alert(`Error: Required column "${HEADER_MAP[key]}" was not found.`);
      return;
    }
    colIdx[key] = index;
  }

  // Add DIN column if it doesn't exist
  if (colIdx.din === -1) {
    colIdx.din = header.length;
    header.push(HEADER_MAP.din);
    body.forEach(row => row.push(''));
  }

  // Load drug database with caching
  console.log('Loading drug database...');
  SpreadsheetApp.getActiveSpreadsheet().toast('Loading Health Canada drug database...', 'Initializing');
  
  const drugDatabase = loadDrugDatabase();
  
  if (!drugDatabase) {
    ui.alert('Error: Failed to load drug database from Health Canada API.');
    return;
  }

  console.log(`Database loaded: ${Object.keys(drugDatabase).length} products with ingredient data`);
  SpreadsheetApp.getActiveSpreadsheet().toast(`Database loaded: ${Object.keys(drugDatabase).length} products`, 'Ready');

  let dinsFound = 0;
  let rowsProcessed = 0;
  const totalRows = body.length;

  for (let i = 0; i < body.length; i++) {
    const row = body[i];
    rowsProcessed++;

    if (row[colIdx.country]?.toString().trim().toUpperCase() !== 'CAN' || 
        row[colIdx.din]?.toString().trim() !== '') {
      continue;
    }

    if (rowsProcessed % 10 === 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Processing row ${rowsProcessed} of ${totalRows}...`, 
        "Progress"
      );
    }

    console.log(`\n========== Processing row ${i + 2}: ${row[colIdx.drug]} ==========`);

    const type = row[colIdx.type]?.toString().trim().toLowerCase();
    console.log(`Row ${i + 2}: Type = "${type}", Drug = "${row[colIdx.drug]}"`);
    let finalDin = null;

    if (type === 'brand') {
      finalDin = findBrandDin(row, colIdx, drugDatabase);
    } else if (type === 'generic') {
      finalDin = findGenericDin(row, colIdx, drugDatabase);
    }

    if (finalDin) {
      row[colIdx.din] = finalDin;
      dinsFound++;
      console.log(`✓ Final DIN: ${finalDin}`);
    } else {
      console.log(`✗ No DIN found`);
    }
  }

  const finalData = [header, ...body];
  sheet.getRange(1, 1, finalData.length, finalData[0].length).setValues(finalData);
 // ui.alert('Process Complete', `Found and updated ${dinsFound} DIN(s).`, ui.ButtonSet.OK);
}

/**
 * Load and cache the entire drug database, handling large data by splitting it into chunks.
 */
function loadDrugDatabase() {
  const cache = CacheService.getScriptCache();
  const CHUNK_SIZE = 90 * 1024; // 90 KB, safely under the 100 KB limit

  // --- Attempt to load from cache ---
  const cacheTimestamp = cache.get('cacheTimestamp');
  if (cacheTimestamp && (Date.now() - parseInt(cacheTimestamp) < CACHE_DURATION)) {
    const chunkCountStr = cache.get('drugData_chunk_count');
    
    if (chunkCountStr) {
      const chunkCount = parseInt(chunkCountStr);
      const chunkKeys = [];
      for (let i = 0; i < chunkCount; i++) {
        chunkKeys.push(`drugData_${i}`);
      }

      const cachedChunks = cache.getAll(chunkKeys);
      let allChunksFound = true;
      let jsonData = '';

      // Reassemble the JSON string from chunks
      for (let i = 0; i < chunkCount; i++) {
        const chunk = cachedChunks[`drugData_${i}`];
        if (chunk) {
          jsonData += chunk;
        } else {
          allChunksFound = false;
          console.log(`Cache miss: Chunk drugData_${i} was missing. Refetching.`);
          break; 
        }
      }

      if (allChunksFound) {
        console.log(`Using cached drug database from ${chunkCount} chunks.`);
        return JSON.parse(jsonData);
      }
    }
  }

  // --- If cache is invalid or missing, fetch fresh data ---
  console.log('Fetching fresh drug database from API...');
  
  const products = fetchAllProducts();
  if (!products) return null;

  // --- NEW: Fetch statuses ---
  const statuses = fetchAllStatuses();
  if (!statuses) return null;

  const ingredients = fetchAllIngredients();
  if (!ingredients) return null;

  const database = {};
  
  // Step 1: Build the base product list
  products.forEach(product => {
    database[product.drug_code] = {
      din: product.drug_identification_number,
      brand_name: product.brand_name,
      drug_code: product.drug_code,
      status: null, // Initialize status
      ingredients: []
    };
  });

  // Step 2: Merge in the statuses
  statuses.forEach(stat => {
    if (database[stat.drug_code]) {
      database[stat.drug_code].status = stat.status;
    }
  });

  // Step 3: Merge in the ingredients
  ingredients.forEach(ing => {
    if (database[ing.drug_code]) {
      database[ing.drug_code].ingredients.push({
        name: ing.ingredient_name,
        strength: ing.strength || ing.strength_value || ''
      });
    }
  });

  // --- Serialize, chunk, and cache the new data ---
  try {
    const dataStr = JSON.stringify(database);
    const numChunks = Math.ceil(dataStr.length / CHUNK_SIZE);
    
    const chunksToCache = {
      'cacheTimestamp': Date.now().toString(),
      'drugData_chunk_count': numChunks.toString()
    };

    for (let i = 0; i < numChunks; i++) {
      const key = `drugData_${i}`;
      const chunk = dataStr.substring(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE);
      chunksToCache[key] = chunk;
    }

    cache.putAll(chunksToCache, 21600); // Cache all for 6 hours
    console.log(`Successfully cached database in ${numChunks} chunks.`);

  } catch (e) {
    console.error(`Failed to cache database: ${e.message}`);
  }

  return database;
}

/**
 * Fetch all drug products
 */
function fetchAllProducts() {
  const url = 'https://health-products.canada.ca/api/drug/drugproduct/';
  
  try {
    const response = UrlFetchApp.fetch(url, { 'muteHttpExceptions': true });
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      return Array.isArray(data) ? data : null;
    }
    return null;
  } catch (e) {
    console.error(`Failed to fetch products: ${e.message}`);
    return null;
  }
}

/**
 * Fetch all drug statuses
 */
function fetchAllStatuses() {
  const url = 'https://health-products.canada.ca/api/drug/status/';
  try {
    const response = UrlFetchApp.fetch(url, { 'muteHttpExceptions': true });
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      return Array.isArray(data) ? data : null;
    }
    console.error(`Failed to fetch statuses: Status code ${response.getResponseCode()}`);
    return null;
  } catch (e) {
    console.error(`Failed to fetch statuses: ${e.message}`);
    return null;
  }
}

/**
 * Fetch all active ingredients
 */
function fetchAllIngredients() {
  const url = 'https://health-products.canada.ca/api/drug/activeingredient/';
  
  try {
    const response = UrlFetchApp.fetch(url, { 'muteHttpExceptions': true });
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      return Array.isArray(data) ? data : null;
    }
    return null;
  } catch (e) {
    console.error(`Failed to fetch ingredients: ${e.message}`);
    return null;
  }
}

function findBrandDin(row, colIdx, database) {
  const brandName = cleanSearchTerm(row[colIdx.drug]?.toString() || '');
  const sheetStrength = normalizeStrength(row[colIdx.strength]?.toString() || '');
  const ingredientsRaw = row[colIdx.ingredients]?.toString() || '';
  
  // Split ingredients by semicolon for combination products
  const sheetIngredients = ingredientsRaw
    .split(';')
    .map(ing => normalizeText(cleanSearchTerm(ing.trim())))
    .filter(ing => ing.length > 0);

  console.log(`Brand: "${brandName}", Strength: "${sheetStrength}", Ingredients: ${JSON.stringify(sheetIngredients)}`);

  // Filter for products that are currently marketed
  const marketedProducts = Object.values(database).filter(product => product.status === 'Marketed');
  
  if (marketedProducts.length === 0) {
    return null;
  }

  // Exact brand match - handles brands with strength in name
let matches = marketedProducts.filter(product => {
  const normalizedProductBrand = normalizeText(product.brand_name);
  const normalizedSearchBrand = normalizeText(brandName);
  
  return normalizedProductBrand === normalizedSearchBrand || 
         normalizedProductBrand.startsWith(normalizedSearchBrand + ' ') ||
         normalizedProductBrand.startsWith(normalizedSearchBrand + '-');
});  // <-- This closing was missing or misplaced

// DEBUG - should be HERE, AFTER the filter closes
console.log(`Found ${matches.length} brand matches for "${brandName}"`);
if (matches.length > 0) {
  console.log(`Sample matches: ${JSON.stringify(matches.slice(0, 3).map(p => ({
    brand: p.brand_name,
    din: p.din,
    numIngredients: p.ingredients.length,
    strengths: p.ingredients.map(i => normalizeStrength(i.strength))
  })))}`);
}

  // Fuzzy brand match if no exact match
  if (matches.length === 0) {
    const baseBrandName = brandName.replace(/\s+(ER|CD|XL|SR|LA|CR|DR|XR)\s*$/i, '').trim();
    if (baseBrandName !== brandName) {
      matches = marketedProducts.filter(product => {
        const productBaseName = normalizeText(product.brand_name.replace(/\s+(ER|CD|XL|SR|LA|CR|DR|XR)\s*$/i, '').trim());
        const brandMatches = productBaseName === normalizeText(baseBrandName);

        // If we have ingredients, verify they match
        if (brandMatches && sheetIngredients.length > 0) {
          const hasAllIngredients = sheetIngredients.every(sheetIng => {
            return product.ingredients.some(ing => {
              const apiIngredient = normalizeText(ing.name);
              return apiIngredient.includes(sheetIng) || sheetIng.includes(apiIngredient);
            });
          });
          return hasAllIngredients;
        }

        return brandMatches;
      });
    }
  }

  if (matches.length === 0) return null;

  if (matches.length === 1) return matches[0].din;

  // Filter by strength - IMPROVED for combination products
  const strengthMatches = matches.filter(product => {
    // Get all strength numbers from the API
    const apiStrengths = product.ingredients
      .map(ing => normalizeStrength(ing.strength))
      .filter(s => s.length > 0)
      .sort();
    
    // Original joined comparison
    const apiStrengthsJoined = apiStrengths.join('');
    if (apiStrengthsJoined === sheetStrength) {
      return true;
    }
    
    // For combination products: check if each API strength appears in sheet strength
    if (apiStrengths.length > 1) {
      return apiStrengths.every(strength => sheetStrength.includes(strength));
    }
    
    // Single ingredient: check if any ingredient matches
    return product.ingredients.some(ing => normalizeStrength(ing.strength) === sheetStrength);
  });

  if (strengthMatches.length > 0) {
    return strengthMatches[0].din;
  }

  return matches[0].din;
}

function findGenericDin(row, colIdx, database) {
  const ingredientsRaw = cleanSearchTerm(row[colIdx.ingredients]?.toString() || '');
  const sheetStrength = normalizeStrength(row[colIdx.strength]?.toString() || '');

  // Split ingredients by semicolon for combination products
  const sheetIngredients = ingredientsRaw.split(';').map(ing => normalizeText(ing.trim())).filter(ing => ing.length > 0);

  console.log(`Generic - Ingredients: ${JSON.stringify(sheetIngredients)}, Strength: "${sheetStrength}"`);

  // Filter for products that are currently marketed first
  const marketedProducts = Object.values(database).filter(product => product.status === 'Marketed');

  const matches = marketedProducts.filter(product => {
    // Check if product has ALL the ingredients from the sheet
    const hasAllIngredients = sheetIngredients.every(sheetIng => {
      return product.ingredients.some(ing => {
        const apiIngredient = normalizeText(ing.name);
        // Partial match in either direction (handles salt forms like "Xinafoate")
        return apiIngredient.includes(sheetIng) || sheetIng.includes(apiIngredient);
      });
    });
    
    if (!hasAllIngredients) return false;

    // For combination products: check if each API strength appears in sheet strength
    const apiStrengths = product.ingredients
      .map(ing => normalizeStrength(ing.strength))
      .filter(s => s.length > 0)
      .sort();
    
    // Check if all API strengths are present in the sheet strength
    if (apiStrengths.length > 1) {
      return apiStrengths.every(strength => sheetStrength.includes(strength));
    }
    
    // Single ingredient: exact match
    return product.ingredients.some(ing => {
      const apiStrength = normalizeStrength(ing.strength);
      return apiStrength === sheetStrength;
    });
  });

  if (matches.length > 0) {
    return matches[0].din;
  }

  return null;
}

/**
 * Normalize strength values
 */
function normalizeStrength(strength) {
  if (!strength || typeof strength !== 'string') return '';
  const numbers = strength.match(/\d+\.?\d*/g);
  if (!numbers) return '';
  return numbers.join('');
}

/**
 * Normalize text for comparison
 */
function normalizeText(text) {
  if (!text || typeof text !== 'string') return '';
  return text.toLowerCase().trim().replace(/\s+/g, ' ');
}

/**
 * Clean search terms
 */
function cleanSearchTerm(term) {
  if (!term || typeof term !== 'string') return '';
  return term.replace(/[®™©]/g, '').replace(/\s*\(.*?\)\s*/g, '').trim();
}


/**
 * ===================================================================
 *                    NDC LOOKUP FUNCTIONALITY
 * ===================================================================
 */




// Use script properties to store the last processed row between runs
const SCRIPT_PROPS = PropertiesService.getScriptProperties();
const NDC_ROW_TRACKER_KEY = 'ndc_last_processed_row'; // A unique key for this specific task

/**
 * STARTER FUNCTION: Kicks off the batch process for finding NDCs.
 * Run this from the menu.
 */
function startNdcPopulation() {
  // Clear any previous progress to start fresh
  SCRIPT_PROPS.deleteProperty(NDC_ROW_TRACKER_KEY);
  deleteAllTriggersByName('continueNdcPopulation'); // Clean up old triggers

  SpreadsheetApp.getActiveSpreadsheet().toast('Starting the NDC lookup process in batches. This may take several minutes.', 'Process Started', 10);

  // Call the main worker function to run the first batch
  continueNdcPopulation();
}

/**
 * STOPPER FUNCTION: Manually stops the batch process.
 * Run this from the menu if you need to cancel the operation.
 */
function stopNdcPopulation() {
  deleteAllTriggersByName('continueNdcPopulation');
  SCRIPT_PROPS.deleteProperty(NDC_ROW_TRACKER_KEY);
  SpreadsheetApp.getActiveSpreadsheet().toast('The NDC lookup process has been stopped.', 'Process Stopped', 5);
}


/**
 * WORKER FUNCTION: Processes rows in batches, resuming where it left off.
 * This function is called automatically by triggers after the first run. DO NOT RUN MANUALLY.
 */
function continueNdcPopulation() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Sheet1');

  if (!sheet) {
    ui.alert("Error: The target sheet named 'Sheet1' was not found.");
    return;
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const header = values[0];
  const body = values.slice(1);
  
  // --- Determine start row and execution time limit ---
  const startTime = new Date();
  const timeLimitInMinutes = 4.5; // Run for 4.5 minutes to be safe
  
  // Get the last row we processed from script properties, or start at 0 if it's the first run
  let startRow = parseInt(SCRIPT_PROPS.getProperty(NDC_ROW_TRACKER_KEY) || '0');

  const dinColIdx = header.indexOf('DIN');
  let ndcColIdx = header.indexOf('NDC');
  if (dinColIdx === -1) {
    ui.alert("Error: 'DIN' column not found in 'Sheet1'.");
    return;
  }
  if (ndcColIdx === -1) {
    ndcColIdx = header.length;
    sheet.getRange(1, ndcColIdx + 1).setValue('NDC');
    body.forEach(row => row.push(''));
  }

  // --- Main Processing Loop ---
  for (let i = startRow; i < body.length; i++) {
    const elapsedMinutes = (new Date() - startTime) / 1000 / 60;
    
    // Check if we've run out of time
    if (elapsedMinutes >= timeLimitInMinutes) {
      // 1. Save our current progress
      SCRIPT_PROPS.setProperty(NDC_ROW_TRACKER_KEY, i);

      // 2. Schedule the next run in ~1 minute
      deleteAllTriggersByName('continueNdcPopulation'); // Clean up first
      ScriptApp.newTrigger('continueNdcPopulation')
        .timeBased()
        .after(60 * 1000) // 60 seconds * 1000 milliseconds
        .create();
      
      // 3. Inform the user and exit
      SpreadsheetApp.getActiveSpreadsheet().toast(`Processed up to row ${i + 1}. Pausing to avoid timeout. Will resume automatically in 1 minute...`, 'Continuing in Background', 10);
      return; // Stop this execution
    }

    const row = body[i];
    const din = row[dinColIdx] ? row[dinColIdx].toString().trim() : '';

    // Only process rows that have a DIN and an empty NDC cell
    if (din && !row[ndcColIdx]) {
      console.log(`Processing DIN: ${din} from Sheet1, row ${i + 2}`);
      const apiData = fetchNdcDataForDin(din);
      body[i][ndcColIdx] = apiData ? parseNdcData(apiData) : 'API Error or no data';
    }
  }

  // --- If the loop finishes, the whole process is done ---
  // 1. Write all data back to the sheet
  if (body.length > 0) {
    sheet.getRange(2, 1, body.length, body[0].length).setValues(body);
  }

  // 2. Clean up properties and triggers
  stopNdcPopulation();

  // 3. Notify the user that everything is complete
 // console.log('NDC population process completed successfully.');
  // ui.alert('Process Complete', 'All NDCs have been successfully retrieved and updated in Sheet1.', ui.ButtonSet.OK);
}


/**
 * UTILITY FUNCTION: Deletes all triggers with a specific function name.
 * This prevents orphaned triggers from building up.
 */
function deleteAllTriggersByName(functionName) {
  const allTriggers = ScriptApp.getProjectTriggers();
  for (const trigger of allTriggers) {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

/**
 * Fetches data from the ndc-din.com API for a single DIN.
 * @param {string} din The Drug Identification Number.
 * @return {object|null} The parsed JSON object or null on error.
 */
function fetchNdcDataForDin(din) {
  const url = `https://ndc-din.com/API-Endpoints/din_details.php?din=${din}&special=true&cached=none`;
  try {
    const response = UrlFetchApp.fetch(url, { 'muteHttpExceptions': true });
    const responseCode = response.getResponseCode();

    if (responseCode === 200) {
      const content = response.getContentText();
      return JSON.parse(content);
    } else {
      console.error(`API error for DIN ${din}: Status ${responseCode}, Response: ${response.getContentText()}`);
      return null;
    }
  } catch (e) {
    console.error(`Failed to fetch or parse JSON for DIN ${din}: ${e.message}`);
    return null;
  }
}

/**
 * Parses the API response to extract NDC-9 codes and descriptions.
 * @param {object} apiData The parsed JSON from the API.
 * @return {string} A formatted string of unique NDC-9 codes and descriptions.
 */
function parseNdcData(apiData) {
  const ndcMatching = apiData['ndc-matching'];
  if (!ndcMatching || Object.keys(ndcMatching).length === 0) {
    return 'No NDC matches found';
  }

  const results = [];
  const seenNdc9 = new Set(); // Use a Set to track unique NDC-9 codes

  // Loop through each full NDC entry (e.g., "0169-4130-13")
  for (const fullNdc in ndcMatching) {
    const details = ndcMatching[fullNdc];
    
    // --- Requirement 1: Get the NDC 9 code ---
    const parts = fullNdc.split('-');
    const ndc9Code = parts.length > 2 ? `${parts[0]}-${parts[1]}` : fullNdc;

    // Only add unique NDC-9 codes to the output
    if (!seenNdc9.has(ndc9Code)) {
      // --- Requirement 2: Get the description ---
      const description = details.description || "No description available";
      results.push(`${ndc9Code}: ${description}`);
      seenNdc9.add(ndc9Code);
    }
  }
  
  // Join all found entries with a newline character to display nicely in one cell
  return results.join('\n');
}
