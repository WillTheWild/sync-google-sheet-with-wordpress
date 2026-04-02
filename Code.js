
// --- Configuration ---
const CONFIG = {
  SHEET_NAME: "SHEET_NAME",
  DRIVE_FOLDER_ID: "DRIVE_FOLDER_ID",

  // WordPress site
  SITE_URL: "SITE_URL",

  // Auth — WooCommerce REST API (consumer key:secret)
  WOO_AUTH: Utilities.base64Encode("consumer_key:consumer_secret"),
  // Auth — WordPress REST API (application password)
  WP_AUTH: Utilities.base64Encode("username:application_password"),

  MAX_SYNC_PRODUCTS: 10, // Max products to fetch when syncing Web → Sheet
  COLUMNS: {
    SKU: 1, NAME: 2, BRAND: 3, CATEGORY: 4, TAGS: 5, SHORT_DESC: 6,
    SPEC: 7, PRICE: 8, SALE_PRICE: 9, STOCK: 10, IMAGE: 11,
    FEATURED: 12, WP_ID: 13, ACTION: 14, STATUS: 15, LAST_SYNC: 16
  }
};

// Derived URLs (don't edit these)
const WOO_API = CONFIG.SITE_URL + "/wp-json/wc/v3";
const WP_API = CONFIG.SITE_URL + "/wp-json/wp/v2";

/**
 * Maps sheet rows to a clean object structure.
 */
function getProductRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const rows = sheet.getDataRange().getValues();
  const C = CONFIG.COLUMNS; // 1-based column numbers

  return rows.slice(1).map((row, i) => ({
    rowIndex: i + 2,
    sku: row[C.SKU - 1],
    name: row[C.NAME - 1],
    brand: row[C.BRAND - 1],
    category: row[C.CATEGORY - 1],
    tags: row[C.TAGS - 1],
    shortDesc: row[C.SHORT_DESC - 1],
    specification: row[C.SPEC - 1],
    price: row[C.PRICE - 1],
    salePrice: row[C.SALE_PRICE - 1],
    stock: row[C.STOCK - 1],
    imageFile: row[C.IMAGE - 1],
    featured: row[C.FEATURED - 1] === true,
    wpId: row[C.WP_ID - 1],
    action: String(row[C.ACTION - 1]).trim().toLowerCase(),
    status: row[C.STATUS - 1]
  }));
}

/**
 * Writes success result to a row: ✅ in STATUS, clears ACTION, sets LAST_SYNC.
 */
function markSuccess(rowIndex, wpId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const cols = CONFIG.COLUMNS;

  if (wpId !== undefined) sheet.getRange(rowIndex, cols.WP_ID).setValue(wpId);
  sheet.getRange(rowIndex, cols.ACTION).setValue("no action");
  sheet.getRange(rowIndex, cols.STATUS).setValue("✅");
  sheet.getRange(rowIndex, cols.LAST_SYNC).setValue(new Date());
}

/**
 * Writes error result to a row: error message in STATUS, keeps ACTION unchanged.
 */
function markError(rowIndex, errorMessage) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const cols = CONFIG.COLUMNS;

  sheet.getRange(rowIndex, cols.STATUS).setValue("❌ " + errorMessage);
  sheet.getRange(rowIndex, cols.LAST_SYNC).setValue(new Date());
}

// --- Image Upload ---

function uploadImageToWP(filename) {
  try {
    const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
    const files = folder.getFilesByName(filename);

    if (files.hasNext()) {
      const file = files.next();
      const blob = file.getBlob();

      const options = {
        method: "POST",
        headers: {
          "Authorization": "Basic " + CONFIG.WP_AUTH,
          "Content-Disposition": `attachment; filename="${filename}"`,
          "Content-Type": blob.getContentType()
        },
        payload: blob.getBytes(),
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch(WP_API + "/media", options);
      const data = JSON.parse(response.getContentText());

      if (data.id) {
        return data.id;
      } else {
        throw new Error(data.message || "Image upload failed");
      }
    } else {
      throw new Error(`Image file "${filename}" not found in Drive`);
    }
  } catch (e) {
    throw new Error("Image upload: " + e.message);
  }
}

// --- WooCommerce API ---

function wpRequest(method, endpoint, payload) {
  const options = {
    method: method,
    headers: {
      "Authorization": "Basic " + CONFIG.WOO_AUTH
    },
    muteHttpExceptions: true
  };

  // Only set Content-Type and payload for non-GET requests
  if (method !== "GET") {
    options.headers["Content-Type"] = "application/json";
    if (payload) {
      options.payload = JSON.stringify(payload);
    }
  }

  const response = UrlFetchApp.fetch(WOO_API + "/products" + endpoint, options);
  const result = JSON.parse(response.getContentText());

  if (response.getResponseCode() >= 400) {
    throw new Error(result.message || `API error (${response.getResponseCode()})`);
  }

  return result;
}

/**
 * Looks up a WooCommerce product category by name.
 * If it doesn't exist, creates it and returns the new ID.
 */
function getOrCreateCategory(categoryName) {
  if (!categoryName) return null;

  // Search for existing category
  const searchUrl = WOO_API + `/products/categories?search=${encodeURIComponent(categoryName)}&per_page=100`;

  const searchOptions = {
    method: "GET",
    headers: {
      "Authorization": "Basic " + CONFIG.WOO_AUTH,
      "Content-Type": "application/json"
    },
    muteHttpExceptions: true
  };

  const searchResponse = UrlFetchApp.fetch(searchUrl, searchOptions);
  const categories = JSON.parse(searchResponse.getContentText());

  // Find exact match (search is fuzzy)
  const match = categories.find(c => c.name.toLowerCase() === categoryName.toLowerCase());
  if (match) return match.id;

  // Category not found — create it
  const createUrl = WOO_API + '/products/categories';
  const createOptions = {
    method: "POST",
    headers: {
      "Authorization": "Basic " + CONFIG.WOO_AUTH,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify({ name: categoryName }),
    muteHttpExceptions: true
  };

  const createResponse = UrlFetchApp.fetch(createUrl, createOptions);
  const newCategory = JSON.parse(createResponse.getContentText());

  if (newCategory.id) return newCategory.id;

  throw new Error(`Failed to create category "${categoryName}": ${newCategory.message || 'Unknown error'}`);
}

/**
 * Looks up a WooCommerce product tag by name.
 * If it doesn't exist, creates it and returns the new ID.
 */
function getOrCreateTag(tagName) {
  if (!tagName) return null;

  const searchUrl = WOO_API + `/products/tags?search=${encodeURIComponent(tagName)}&per_page=100`;

  const searchOptions = {
    method: "GET",
    headers: {
      "Authorization": "Basic " + CONFIG.WOO_AUTH,
      "Content-Type": "application/json"
    },
    muteHttpExceptions: true
  };

  const searchResponse = UrlFetchApp.fetch(searchUrl, searchOptions);
  const tags = JSON.parse(searchResponse.getContentText());

  // Find exact match (search is fuzzy)
  const match = tags.find(t => t.name.toLowerCase() === tagName.toLowerCase());
  if (match) return match.id;

  // Tag not found — create it
  const createUrl = WOO_API + '/products/tags';
  const createOptions = {
    method: "POST",
    headers: {
      "Authorization": "Basic " + CONFIG.WOO_AUTH,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify({ name: tagName }),
    muteHttpExceptions: true
  };

  const createResponse = UrlFetchApp.fetch(createUrl, createOptions);
  const newTag = JSON.parse(createResponse.getContentText());

  if (newTag.id) return newTag.id;

  throw new Error(`Failed to create tag "${tagName}": ${newTag.message || 'Unknown error'}`);
}

function buildPayload(row, existingProduct) {
  const payload = {
    name: row.name,
    sku: row.sku,
    regular_price: String(row.price).replace(/,/g, ''),
    sale_price: row.salePrice ? String(row.salePrice).replace(/,/g, '') : "",
    stock_quantity: parseInt(row.stock) || 0,
    manage_stock: true,
    short_description: row.shortDesc || "",
    description: row.specification || "",
    categories: row.category ? [{ id: getOrCreateCategory(row.category) }] : [],
    tags: row.tags
      ? String(row.tags).split(',').map(t => t.trim()).filter(Boolean).map(t => ({ id: getOrCreateTag(t) }))
      : [],
    featured: row.featured
  };

  // Handle image
  if (row.imageFile) {
    // Check if existing product already has this image (skip re-upload)
    let existingImageName = '';
    if (existingProduct && existingProduct.images && existingProduct.images.length > 0) {
      existingImageName = existingProduct.images[0].src.split('/').pop().split('?')[0];
    }

    if (existingImageName === row.imageFile) {
      // Same image — keep existing, don't re-upload
      payload.images = [{ id: existingProduct.images[0].id }];
    } else {
      // New or changed image — upload it
      try {
        const wpImageId = uploadImageToWP(row.imageFile);
        if (wpImageId) {
          payload.images = [{ id: wpImageId }];
        }
      } catch (imgErr) {
        Logger.log('Image upload error for row ' + row.rowIndex + ': ' + imgErr.message);
        // Don't fail the whole product update over an image error
      }
    }
  }

  return payload;
}

// --- Action Handlers ---

function handleAdd(row) {
  const payload = buildPayload(row, null);
  const result = wpRequest("POST", "", payload);

  if (!result.id) throw new Error(result.message || "Failed to create product");

  markSuccess(row.rowIndex, result.id);
}

function handleUpdate(row) {
  if (!row.wpId) throw new Error("Cannot update: no WP ID found");

  // Fetch existing product to compare images
  let existingProduct = null;
  try {
    existingProduct = wpRequest("GET", `/${row.wpId}`);
  } catch (e) {
    Logger.log('Could not fetch existing product ' + row.wpId + ': ' + e.message);
  }

  const payload = buildPayload(row, existingProduct);
  const result = wpRequest("PUT", `/${row.wpId}`, payload);

  if (!result.id) throw new Error(result.message || "Failed to update product");

  markSuccess(row.rowIndex, row.wpId);
}

function handleDelete(row) {
  if (!row.wpId) throw new Error("Cannot delete: no WP ID found");

  wpRequest("DELETE", `/${row.wpId}?force=true`, null);

  // Clear WP ID after successful deletion
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  sheet.getRange(row.rowIndex, CONFIG.COLUMNS.WP_ID).setValue("");
  sheet.getRange(row.rowIndex, CONFIG.COLUMNS.ACTION).setValue("no action");
  sheet.getRange(row.rowIndex, CONFIG.COLUMNS.STATUS).setValue("✅ Deleted");
  sheet.getRange(row.rowIndex, CONFIG.COLUMNS.LAST_SYNC).setValue(new Date());
}

// --- Main Sync (Sheet → Web) ---

/**
 * Processes each row based on its ACTION column value.
 * Actions: add, update, delete, no action
 */
function runSync() {
  const rows = getProductRows();

  rows.forEach(row => {
    if (!row.sku) return;

    const action = row.action;

    // Skip rows with no action or already completed
    if (!action || action === "no action" || action === "✅") return;

    try {
      switch (action) {
        case "add":
          handleAdd(row);
          break;

        case "update":
          handleUpdate(row);
          break;

        case "delete":
          handleDelete(row);
          break;

        default:
          // Unknown action — skip silently
          return;
      }
    } catch (err) {
      markError(row.rowIndex, err.message);
    }
  });
}

// --- Sync from Web → Sheet ---

/**
 * Fetches products from WooCommerce (handles pagination).
 * Uses direct fetch to avoid parameter issues.
 * WooCommerce returns 10 per page by default.
 */
function fetchAllWooProducts(maxItems) {
  maxItems = maxItems || CONFIG.MAX_SYNC_PRODUCTS;
  const allProducts = [];
  let page = 1;

  while (true) {
    const url = WOO_API + `/products?page=${page}&orderby=id&order=asc`;
    const options = {
      method: "GET",
      headers: {
        "Authorization": "Basic " + CONFIG.WOO_AUTH
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();

    if (code >= 400) {
      throw new Error("API error fetching products (HTTP " + code + "): " + response.getContentText());
    }

    const batch = JSON.parse(response.getContentText());

    if (!Array.isArray(batch) || batch.length === 0) break;

    allProducts.push(...batch);

    // Stop if we've reached the limit
    if (allProducts.length >= maxItems) break;

    // If we got fewer than 10 (default page size), we've reached the last page
    if (batch.length < 10) break;
    page++;
  }

  // Trim to exact max
  return allProducts.slice(0, maxItems);
}

/**
 * Downloads an image from a URL and saves it to the configured Google Drive folder.
 * Skips download if a file with the same name already exists in the folder.
 * Returns the filename on success, or empty string on failure.
 */
function downloadImageToDrive(imageUrl) {
  if (!imageUrl) return '';

  try {
    const filename = imageUrl.split('/').pop().split('?')[0]; // strip query params
    if (!filename) return '';

    const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);

    // Check if file already exists — skip to avoid duplicates
    const existing = folder.getFilesByName(filename);
    if (existing.hasNext()) {
      return filename;
    }

    // Download the image
    const response = UrlFetchApp.fetch(imageUrl, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      Logger.log('Failed to download image: ' + imageUrl + ' (HTTP ' + response.getResponseCode() + ')');
      return filename; // still return filename for the sheet
    }

    const blob = response.getBlob().setName(filename);
    folder.createFile(blob);

    return filename;
  } catch (e) {
    Logger.log('Image download error: ' + e.message + ' | URL: ' + imageUrl);
    return imageUrl.split('/').pop() || '';
  }
}

/**
 * Maps a WooCommerce product object to a sheet row array.
 * Also downloads the product image to Google Drive.
 * Columns: SKU, Name, Brand, Category, Tags, Short Desc,
 *          Spec, Price, Sale Price, Stock, Image, Featured, WP_ID, Action, Status, Last Sync
 */
function mapProductToRow(product) {
  // Extract category names
  const categoryNames = (product.categories || []).map(c => c.name).join(', ');

  // Extract tag names
  const tagNames = (product.tags || []).map(t => t.name).join(', ');

  // Extract brand from attributes (or meta_data)
  let brand = '';
  if (product.attributes && product.attributes.length > 0) {
    const brandAttr = product.attributes.find(a => a.name.toLowerCase() === 'brand');
    if (brandAttr && brandAttr.options && brandAttr.options.length > 0) {
      brand = brandAttr.options[0];
    }
  }
  // Fallback: check meta_data for brand
  if (!brand && product.meta_data) {
    const brandMeta = product.meta_data.find(m => m.key === '_brand' || m.key === 'brand');
    if (brandMeta) brand = brandMeta.value;
  }

  // Download first image to Google Drive and get the filename
  // Wrapped in try/catch so image errors never stop the sync
  let imageFile = '';
  if (product.images && product.images.length > 0) {
    try {
      imageFile = downloadImageToDrive(product.images[0].src);
    } catch (imgErr) {
      Logger.log('Image error for product ' + product.id + ': ' + imgErr.message);
      // Fall back to just the filename from the URL
      imageFile = product.images[0].src.split('/').pop() || '';
    }
  }

  return [
    product.sku || '',                                        // SKU
    product.name || '',                                       // Name
    brand,                                                    // Brand
    categoryNames,                                            // Category
    tagNames,                                                 // Tags
    product.short_description || '',                           // Short Desc
    product.description || '',                                // Spec (description)
    product.regular_price || '',                              // Price
    product.sale_price || '',                                 // Sale Price
    product.stock_quantity != null ? product.stock_quantity : '',  // Stock
    imageFile,                                                // Image
    product.featured || false,                                // Featured
    product.id,                                               // WP_ID
    'no action',                                              // Action
    '✅ Synced from web',                                     // Status
    new Date()                                                // Last Sync
  ];
}

/**
 * Pulls all products from WooCommerce and writes them into the sheet.
 * Existing data (below header) is cleared first.
 * Products are sorted by WP ID (ascending).
 */
function syncFromWebToSheet() {
  const ui = SpreadsheetApp.getUi();

  // Confirm before overwriting
  const confirm = ui.alert(
    'Sync from Web to Sheet',
    'This will REPLACE all product rows in the sheet with data from WooCommerce.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);

  try {
    // Fetch all products
    const products = fetchAllWooProducts();

    if (products.length === 0) {
      ui.alert('No products found on WooCommerce.');
      return;
    }

    // Sort by WP ID ascending
    products.sort((a, b) => a.id - b.id);

    // Map to sheet rows
    const rows = products.map(mapProductToRow);

    // Clear existing data (keep header row)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 16).clearContent();
    }

    // Write new data
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 16).setValues(rows);
    }

    ui.alert(`✅ Successfully synced ${products.length} products from WooCommerce to Sheet.`);

  } catch (err) {
    ui.alert('❌ Error syncing from web: ' + err.message);
  }
}

// --- Triggers ---

/**
 * Auto-sets ACTION to "update" when product data columns (A–J) are edited
 * on a row that was previously synced.
 */
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== CONFIG.SHEET_NAME) return;

  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();

  if (row <= 1) return; // Skip header

  // Only trigger if data columns (SKU through IMAGE, 1 to 11) are edited
  if (col >= 1 && col <= 11) {
    const wpId = sheet.getRange(row, CONFIG.COLUMNS.WP_ID).getValue();
    const actionCell = sheet.getRange(row, CONFIG.COLUMNS.ACTION);
    const currentAction = String(actionCell.getValue()).trim().toLowerCase();

    // If the product exists on WP and isn't already queued for delete, set to "update"
    if (wpId && currentAction !== "delete") {
      actionCell.setValue("update");
      sheet.getRange(row, CONFIG.COLUMNS.STATUS).setValue(""); // Clear old status
    }
  }
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('WooCommerce Sync')
    .addItem('📤 Sync Sheet → Web', 'runSync')
    .addItem('📥 Sync Web → Sheet', 'syncFromWebToSheet')
    .addSeparator()
    .addItem('⏱ Setup Time Trigger (15m)', 'createTimeTrigger')
    .addToUi();
}

function createTimeTrigger() {
  // Remove existing triggers to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => { if (t.getHandlerFunction() === 'runSync') ScriptApp.deleteTrigger(t); });

  ScriptApp.newTrigger("runSync")
    .timeBased()
    .everyMinutes(15)
    .create();
}
