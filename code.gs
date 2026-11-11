function doGet(e) {
  const params = e.parameter;
  const page = params.page || "home";
  const isAdmin = params.admin !== undefined || page === "admin";

  const pageMap = {
    home: "index.html",
    privacy: "privacyPolicy.html",
    tnc: "tnc.html",
    about: "about.html",
    admin: "admin.html",
  };

  const pageTitles = {
    home: "Online Catalog",
    privacy: "Privacy Policy",
    tnc: "Terms and Conditions",
    about: "About",
    admin: "Admin Panel",
  };

  const pageTitle = pageTitles[page] || "Online Catalog";

  // ✅ Admin page (any Google user allowed)
  if (isAdmin) {
    const email = Session.getActiveUser().getEmail(); // Optional
    const template = HtmlService.createTemplateFromFile("admin.html");
    template.userEmail = email || "Unknown User";
    return template
      .evaluate()
      .setTitle(pageTitle)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // ✅ Public pages
  return HtmlService.createTemplateFromFile(pageMap[page] || "index.html")
    .evaluate()
    .setTitle(pageTitle)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- Include HTML Partials ---
function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}

const PRODUCT_SCHEMA_REFERENCE = {
  title: "string (required)",
  description: "string",
  price: "number",
  discount: "number (0-100%)",
  category: "string",
  image: { fileId: "string", fileName: "string" },
  tags: "array of strings",
  features: "array of strings",
  specs: "array of strings",
  section: "string (see SECTIONS)",
  rank: "number (lower = higher priority)",
  dateAdded: "YYYY-MM-DD",
  paymentUrl: "string (Stripe URL)",
};

// --- Function to get all sheets in the spreadsheet ---
function getSheets() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    return sheets.map(function (sheet) {
      return {
        name: sheet.getName(),
        id: sheet.getSheetId(),
      };
    });
  } catch (error) {
    Logger.log("Error getting sheets: " + error.toString());
    return { error: "Error getting sheets: " + error.toString() };
  }
}

// --- Function to get headers from a specific sheet ---
function getSheetHeaders(sheetName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return { error: "Sheet not found" };
    }

    // Get the first row (headers)
    var headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    var headers = headerRange.getValues()[0];

    // Filter out empty headers and return the array
    headers = headers.filter(function (header) {
      return header && header.toString().trim() !== "";
    });

    if (headers.length === 0) {
      return { error: "No headers found in the sheet" };
    }

    return headers;
  } catch (error) {
    Logger.log("Error getting headers: " + error.toString());
    return { error: "Error getting headers: " + error.toString() };
  }
}

function getSiteSettings() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const settings = scriptProperties.getProperty("siteSettings");
    return settings;
  } catch (e) {
    console.error("Error in getSiteSettings: " + e.toString());
    throw new Error("Failed to retrieve site settings.");
  }
}

function saveSiteSettings(settingsJSON) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty("siteSettings", settingsJSON);
    return true; // Indicate success
  } catch (e) {
    console.error("Error in saveSiteSettings: " + e.toString());
    throw new Error("Failed to save site settings.");
  }
}

function getAllSheetsInfo() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var sheetsInfo = {};

    sheets.forEach(function (sheet) {
      var sheetName = sheet.getName();
      var headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
      var headers = headerRange.getValues()[0];

      // Filter out empty headers
      headers = headers.filter(function (header) {
        return header && header.toString().trim() !== "";
      });

      sheetsInfo[sheetName] = headers;
    });

    return {
      success: true,
      sheetsInfo,
    };
  } catch (error) {
    Logger.log("Error getting sheets info: " + error.toString());
    return {
      success: false,
      error: "Error getting sheets info: " + error.toString(),
    };
  }
}

function getSheetData() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getDataRange().getValues();
    return JSON.stringify(data);
  } catch (error) {
    console.error("Error getting sheet data:", error);
    return JSON.stringify({ error: "Failed to get sheet data" });
  }
}

function getLoggedInUserEmail() {
  try {
    return Session.getActiveUser().getEmail();
  } catch (error) {
    Logger.log("Error getting user email: " + error.toString());
    return { error: "Error getting user email: " + error.toString() };
  }
}

function getSpreadsheetOwner() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    return ss.getOwner().getEmail();
  } catch (error) {
    Logger.log("Error getting spreadsheet owner: " + error.toString());
    return { error: "Error getting spreadsheet owner: " + error.toString() };
  }
}

function isUserOwner() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ownerEmail = ss.getOwner().getEmail();
    Logger.log("User email: " + userEmail);
    Logger.log("Owner email: " + ownerEmail);
    return userEmail === ownerEmail;
  } catch (error) {
    Logger.log("Error checking user permissions: " + error.toString());
    return false;
  }
}

function saveProductSchema(schema) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty("productSchema", JSON.stringify(schema));
  return true;
}

function getProductSchema() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const schema = scriptProperties.getProperty("productSchema");
  return schema ? JSON.parse(schema) : null;
}

// =============================================================================
// IMAGE UPLOAD FUNCTIONS
// =============================================================================

// Create or get the uploads folder in Google Drive
function getUploadsFolder() {
  try {
    const folderName = "Product Images";
    const folders = DriveApp.getFoldersByName(folderName);

    if (folders.hasNext()) {
      return folders.next();
    } else {
      // Create the folder if it doesn't exist
      return DriveApp.createFolder(folderName);
    }
  } catch (error) {
    Logger.log("Error getting uploads folder: " + error.toString());
    throw error;
  }
}

// Upload image to Google Drive
function uploadImageToDrive(base64Data, fileName, mimeType) {
  try {
    // Remove data URL prefix if present
    const base64Content = base64Data.replace(/^data:image\/[a-z]+;base64,/, "");

    // Convert base64 to blob
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Content),
      mimeType,
      fileName
    );

    // Get uploads folder
    const folder = getUploadsFolder();

    // Create file in the folder
    const file = folder.createFile(blob);

    // Make file publicly viewable
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Return file information
    return {
      success: true,
      fileId: file.getId(),
      fileName: file.getName(),
    };
  } catch (error) {
    Logger.log("Error uploading image: " + error.toString());
    return {
      success: false,
      error: "Error uploading image: " + error.toString(),
    };
  }
}

// Delete image from Google Drive
function deleteImageFromDrive(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);
    file.setTrashed(true);

    return {
      success: true,
      message: "Image deleted successfully",
    };
  } catch (error) {
    Logger.log("Error deleting image: " + error.toString());
    return {
      success: false,
      error: "Error deleting image: " + error.toString(),
    };
  }
}

// =============================================================================
// PRODUCTS ORM - CRUD Operations for Products Sheet
// =============================================================================

const PRODUCTS_SHEET_NAME = "Products";

// Initialize Products sheet if it doesn't exist
function initializeProductsSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(PRODUCTS_SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(PRODUCTS_SHEET_NAME);
      // Set headers
      sheet.getRange(1, 1, 1, 2).setValues([["id", "formData"]]);
      sheet.getRange(1, 1, 1, 2).setFontWeight("bold");
    }

    return { success: true, message: "Products sheet initialized" };
  } catch (error) {
    Logger.log("Error initializing products sheet: " + error.toString());
    return { success: false, error: error.toString() };
  }
}

// Generate unique ID for new products
function generateProductId() {
  return "prod_" + Date.now() + "_" + Math.random().toString(36).substr(2, 9);
}

// CREATE - Add new product
function createProduct(formData) {
  try {
    initializeProductsSheet();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(PRODUCTS_SHEET_NAME);

    const id = generateProductId();
    const formDataString = JSON.stringify(formData);

    // Add new row
    sheet.appendRow([id, formDataString]);

    return {
      success: true,
      id: id,
      message: "Product created successfully",
    };
  } catch (error) {
    Logger.log("Error creating product: " + error.toString());
    return { success: false, error: error.toString() };
  }
}

// READ - Get all products
function getAllProducts() {
  try {
    initializeProductsSheet();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(PRODUCTS_SHEET_NAME);

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, products: [] };
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    const products = data.map((row) => {
      return {
        id: row[0],
        formData: row[1] ? JSON.parse(row[1]) : {},
      };
    });
    return { success: true, products: products };
  } catch (error) {
    Logger.log("Error getting products: " + error.toString());
    return { success: false, error: error.toString() };
  }
}

// READ - Get single product by ID
function getProductById(id) {
  try {
    initializeProductsSheet();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(PRODUCTS_SHEET_NAME);

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: false, error: "Product not found" };
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === id) {
        return {
          success: true,
          product: {
            id: data[i][0],
            formData: data[i][1] ? JSON.parse(data[i][1]) : {},
          },
        };
      }
    }

    return { success: false, error: "Product not found" };
  } catch (error) {
    Logger.log("Error getting product: " + error.toString());
    return { success: false, error: error.toString() };
  }
}

// UPDATE - Update existing product
function updateProduct(id, formData) {
  try {
    initializeProductsSheet();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(PRODUCTS_SHEET_NAME);

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: false, error: "Product not found" };
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === id) {
        const rowIndex = i + 2; // +2 because array is 0-indexed and we start from row 2
        const formDataString = JSON.stringify(formData);
        sheet.getRange(rowIndex, 2).setValue(formDataString);

        return {
          success: true,
          message: "Product updated successfully",
        };
      }
    }

    return { success: false, error: "Product not found" };
  } catch (error) {
    Logger.log("Error updating product: " + error.toString());
    return { success: false, error: error.toString() };
  }
}

// DELETE - Delete product by ID
function deleteProduct(id) {
  try {
    initializeProductsSheet();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(PRODUCTS_SHEET_NAME);

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: false, error: "Product not found" };
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === id) {
        const rowIndex = i + 2; // +2 because array is 0-indexed and we start from row 2
        sheet.deleteRow(rowIndex);

        return {
          success: true,
          message: "Product deleted successfully",
        };
      }
    }

    return { success: false, error: "Product not found" };
  } catch (error) {
    Logger.log("Error deleting product: " + error.toString());
    return { success: false, error: error.toString() };
  }
}

// SEARCH - Search products by any field in formData
function searchProducts(query) {
  try {
    const result = getAllProducts();
    if (!result.success) {
      return result;
    }

    const searchTerm = query.toLowerCase();
    const filteredProducts = result.products.filter((product) => {
      const searchableText = JSON.stringify(product.formData).toLowerCase();
      return searchableText.includes(searchTerm);
    });

    return { success: true, products: filteredProducts };
  } catch (error) {
    Logger.log("Error searching products: " + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Subscribe email to newsletter
 * Creates a Subscribers sheet if it doesn't exist and adds the email with timestamp
 * @param {string} email - The email address to subscribe
 * @returns {Object} - Success/error response
 */
function subscribeToNewsletter(email) {
  try {
    // Validate email format
    if (!email || !isValidEmail(email)) {
      return {
        success: false,
        error: "Please enter a valid email address",
      };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let subscribersSheet;

    // Check if Subscribers sheet exists, create if not
    try {
      subscribersSheet = ss.getSheetByName("Subscribers");
    } catch (e) {
      subscribersSheet = null;
    }

    if (!subscribersSheet) {
      subscribersSheet = ss.insertSheet("Subscribers");
      // Add headers
      subscribersSheet
        .getRange(1, 1, 1, 4)
        .setValues([["Email", "Subscription Date", "Status", "Source"]]);

      // Format headers
      const headerRange = subscribersSheet.getRange(1, 1, 1, 4);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#4F46E5");
      headerRange.setFontColor("white");

      // Set column widths
      subscribersSheet.setColumnWidth(1, 250); // Email
      subscribersSheet.setColumnWidth(2, 150); // Date
      subscribersSheet.setColumnWidth(3, 100); // Status
      subscribersSheet.setColumnWidth(4, 120); // Source
    }

    // Check if email already exists
    const lastRow = subscribersSheet.getLastRow();
    if (lastRow > 1) {
      const emailColumn = subscribersSheet
        .getRange(2, 1, lastRow - 1, 1)
        .getValues();
      const existingEmails = emailColumn.map((row) =>
        row[0].toString().toLowerCase()
      );

      if (existingEmails.includes(email.toLowerCase())) {
        return {
          success: false,
          error: "This email is already subscribed to our newsletter",
        };
      }
    }

    // Add new subscription
    const currentDate = new Date();
    const formattedDate = Utilities.formatDate(
      currentDate,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd HH:mm:ss"
    );

    const newRow = lastRow + 1;
    subscribersSheet
      .getRange(newRow, 1, 1, 4)
      .setValues([[email, formattedDate, "Active", "Website"]]);

    // Format the new row
    const newRowRange = subscribersSheet.getRange(newRow, 1, 1, 4);
    newRowRange.setBorder(true, true, true, true, true, true);

    // Alternate row coloring
    if (newRow % 2 === 0) {
      newRowRange.setBackground("#F8FAFC");
    }

    // Log the subscription
    console.log(`New newsletter subscription: ${email} at ${formattedDate}`);

    return {
      success: true,
      message:
        "Successfully subscribed to our newsletter! Thank you for joining us.",
    };
  } catch (error) {
    console.error("Error in subscribeToNewsletter:", error);
    return {
      success: false,
      error:
        "An error occurred while processing your subscription. Please try again.",
    };
  }
}

/**
 * Helper function to validate email format
 * @param {string} email - Email to validate
 * @returns {boolean} - True if valid email format
 */
function isValidEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * Get all newsletter subscribers
 * @returns {Object} - List of subscribers or error
 */
function getNewsletterSubscribers() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let subscribersSheet;

    try {
      subscribersSheet = ss.getSheetByName("Subscribers");
    } catch (e) {
      return {
        success: true,
        subscribers: [],
        message: "No subscribers sheet found",
      };
    }

    if (!subscribersSheet) {
      return {
        success: true,
        subscribers: [],
        message: "No subscribers found",
      };
    }

    const lastRow = subscribersSheet.getLastRow();
    if (lastRow <= 1) {
      return {
        success: true,
        subscribers: [],
        message: "No subscribers found",
      };
    }

    // Get all data except headers
    const data = subscribersSheet.getRange(2, 1, lastRow - 1, 4).getValues();

    const subscribers = data.map((row, index) => ({
      id: index + 1,
      email: row[0],
      subscriptionDate: row[1],
      status: row[2],
      source: row[3],
    }));

    return {
      success: true,
      subscribers: subscribers,
      total: subscribers.length,
    };
  } catch (error) {
    console.error("Error getting newsletter subscribers:", error);
    return {
      success: false,
      error: "Error retrieving subscribers: " + error.toString(),
    };
  }
}

/**
 * Unsubscribe email from newsletter
 * @param {string} email - Email to unsubscribe
 * @returns {Object} - Success/error response
 */
function unsubscribeFromNewsletter(email) {
  try {
    if (!email || !isValidEmail(email)) {
      return {
        success: false,
        error: "Please enter a valid email address",
      };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const subscribersSheet = ss.getSheetByName("Subscribers");

    if (!subscribersSheet) {
      return {
        success: false,
        error: "Subscribers sheet not found",
      };
    }

    const lastRow = subscribersSheet.getLastRow();
    if (lastRow <= 1) {
      return {
        success: false,
        error: "Email not found in subscribers list",
      };
    }

    // Find the email
    const data = subscribersSheet.getRange(2, 1, lastRow - 1, 4).getValues();
    let rowToUpdate = -1;

    for (let i = 0; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase() === email.toLowerCase()) {
        rowToUpdate = i + 2; // +2 because array is 0-indexed and we start from row 2
        break;
      }
    }

    if (rowToUpdate === -1) {
      return {
        success: false,
        error: "Email not found in subscribers list",
      };
    }

    // Update status to "Unsubscribed"
    subscribersSheet.getRange(rowToUpdate, 3).setValue("Unsubscribed");

    console.log(`Unsubscribed email: ${email}`);

    return {
      success: true,
      message: "Successfully unsubscribed from newsletter",
    };
  } catch (error) {
    console.error("Error in unsubscribeFromNewsletter:", error);
    return {
      success: false,
      error:
        "An error occurred while processing your unsubscription. Please try again.",
    };
  }
}

/**
 * Adds a new subscriber to the Subscribers sheet
 * @param {string} email - The subscriber's email address
 * @return {Object} Result of the operation
 */
function addSubscriber(email) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Subscribers");
  const timestamp = new Date();
  const formattedDate = Utilities.formatDate(
    timestamp,
    Session.getScriptTimeZone(),
    "yyyy-MM-dd HH:mm:ss"
  );
  sheet.appendRow([email, formattedDate, "Active"]);
  return { success: true };
}

function unsubscribeEmail(email) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Subscribers");
  const data = sheet.getDataRange().getValues();
  const emailCol = 0;
  const statusCol = 2;

  for (let i = 1; i < data.length; i++) {
    if (data[i][emailCol] === email) {
      sheet.getRange(i + 1, statusCol + 1).setValue("Unsubscribed");
      return { success: true };
    }
  }
  return { success: false, error: "Email not found" };
}
