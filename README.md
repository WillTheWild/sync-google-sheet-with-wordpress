# WooCommerce ↔ Google Sheets Sync

A Google Apps Script that provides **bidirectional sync** between a Google Sheets spreadsheet and a WooCommerce store. Manage your product catalog from a spreadsheet — add, update, and delete products — or pull your entire store inventory back into the sheet.

## ✨ Features

- **📤 Sheet → Web**: Push product data from Google Sheets to WooCommerce (add / update / delete)
- **📥 Web → Sheet**: Pull all WooCommerce products back into the sheet, sorted by product ID
- **🖼️ Image sync**: Upload images from Google Drive to WordPress Media, and download product images from WooCommerce back to Drive
- **🏷️ Auto-create categories & tags**: Categories and tags are created on-the-fly if they don't exist
- **⚡ Auto-detect changes**: Editing product data columns automatically sets the row action to "update"
- **⏱️ Scheduled sync**: Optional time-based trigger to sync every 15 minutes
- **✅ Status tracking**: Each row shows sync status (✅ / ❌) and last sync timestamp

## 📋 Sheet Structure

Your Google Sheet must have these columns (in order):

| Col | Field | Description |
|-----|-------|-------------|
| A | **SKU** | Product SKU (unique identifier) |
| B | **Name** | Product name |
| C | **Brand** | Brand (synced via WooCommerce attributes) |
| D | **Category** | Product category name |
| E | **Tags** | Comma-separated tag names |
| F | **Short Desc** | Short description |
| G | **Specification** | Full description / specs |
| H | **Price** | Regular price |
| I | **Sale Price** | Sale price (optional) |
| J | **Stock** | Stock quantity |
| K | **Image** | Image filename (must exist in Google Drive folder) |
| L | **Featured** | `TRUE` / `FALSE` |
| M | **WP ID** | WordPress product ID (auto-filled after sync) |
| N | **Action** | `add` / `update` / `delete` / `no action` |
| O | **Status** | Sync result (auto-filled) |
| P | **Last Sync** | Timestamp of last sync (auto-filled) |

## 🚀 Setup

### 1. Prerequisites

- A WordPress site with **WooCommerce** installed
- **WooCommerce REST API** keys ([Generate here](https://woocommerce.com/document/woocommerce-rest-api/))
- A **WordPress Application Password** ([How to create](https://make.wordpress.org/core/2020/11/05/application-passwords-integration-guide/))
- A **Google Drive folder** for product images

### 2. Installation

1. Open your Google Sheet
2. Go to **Extensions → Apps Script**
3. Delete any existing code and paste the contents of `sync_sheet_with_wordpress.js`
4. Update the `CONFIG` object at the top:

```javascript
const CONFIG = {
  SHEET_NAME: "Products",                          // Your sheet tab name
  DRIVE_FOLDER_ID: "your-google-drive-folder-id",  // Google Drive folder for images

  SITE_URL: "https://your-site.com",               // Your WordPress site URL

  WOO_AUTH: Utilities.base64Encode("ck_xxx:cs_xxx"),  // WooCommerce consumer key:secret
  WP_AUTH: Utilities.base64Encode("user:app_password"), // WP username:application_password

  MAX_SYNC_PRODUCTS: 10,  // Max products to fetch when syncing Web → Sheet
  // ...
};
```

5. Save the script (`Ctrl+S`)
6. Reload the spreadsheet — you'll see a **WooCommerce Sync** menu appear

### 3. Authorize

The first time you run a sync, Google will ask you to authorize the script. Grant the required permissions (Sheets, Drive, UrlFetch).

## 📖 Usage

### Menu Options

After setup, a **WooCommerce Sync** menu appears in your spreadsheet:

| Menu Item | Description |
|-----------|-------------|
| 📤 **Sync Sheet → Web** | Processes all rows with an action (`add` / `update` / `delete`) |
| 📥 **Sync Web → Sheet** | Fetches products from WooCommerce and writes them to the sheet |
| ⏱ **Setup Time Trigger** | Creates a 15-minute recurring trigger for auto-sync (Sheet → Web) |

### Syncing Sheet → Web

1. Fill in your product data in columns A–L
2. Set the **Action** column (N):
   - `add` — Create a new product on WooCommerce
   - `update` — Update an existing product (requires WP ID in column M)
   - `delete` — Delete the product from WooCommerce
   - `no action` — Skip this row
3. Run **📤 Sync Sheet → Web** from the menu
4. Check columns O–P for sync results

> **Tip**: When you edit columns A–K on a previously synced row, the Action is automatically set to `update`.

### Syncing Web → Sheet

1. Run **📥 Sync Web → Sheet** from the menu
2. Confirm the action (this replaces all existing rows)
3. Products are fetched, sorted by WP ID, and written to the sheet
4. Product images are downloaded to your Google Drive folder

## ⚙️ Configuration Reference

| Config Key | Description |
|------------|-------------|
| `SHEET_NAME` | Name of the sheet tab containing product data |
| `DRIVE_FOLDER_ID` | Google Drive folder ID where product images are stored |
| `SITE_URL` | Your WordPress site URL (no trailing slash) |
| `WOO_AUTH` | Base64-encoded WooCommerce API `consumer_key:consumer_secret` |
| `WP_AUTH` | Base64-encoded WordPress `username:application_password` |
| `MAX_SYNC_PRODUCTS` | Maximum number of products to fetch in Web → Sheet sync |
| `COLUMNS` | Column mapping (1-based index) — edit if your sheet layout differs |

## License

MIT
