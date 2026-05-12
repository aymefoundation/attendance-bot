const { google } = require("googleapis");
const path = require("path");

let auth;
let sheets;
const SPREADSHEET_ID = "1hjV7t0kgcMnQfHOSM55lCJeIk8zdBrD_7tKAxvEuD38";

try {
  const credPath = path.join(__dirname, "credentials.json");
  console.log(`📂 Loading credentials from: ${credPath}`);

  auth = new google.auth.GoogleAuth({
    keyFile: credPath,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });

  sheets = google.sheets({
    version: "v4",
    auth,
  });

  console.log("✅ Google Sheets authentication initialized");
} catch (error) {
  console.error("❌ Failed to initialize Google Sheets:", error.message);
  process.exit(1);
}

// Test connection function
async function testConnection() {
  try {
    console.log("🔍 Testing Google Sheets connection...");
    const res = await sheets.spreadsheets.get({
      spreadsheetId: SPREADSHEET_ID,
    });
    console.log(`✅ Connected to sheet: ${res.data.properties.title}`);
    return true;
  } catch (error) {
    if (error.message.includes("does not have permission")) {
      console.warn("\\n⚠️  PERMISSION ERROR - Share the Google Sheet with:");
      console.warn(
        "   📧 attendance-bot@attendance-bot-496018.iam.gserviceaccount.com",
      );
      console.warn(
        "   Link: https://docs.google.com/spreadsheets/d/1hjV7t0kgcMnQfHOSM55lCJeIk8zdBrD_7tKAxvEuD38\\n",
      );
    } else {
      console.error("❌ Connection test failed:", error.message);
    }
    return false;
  }
}

module.exports = {
  sheets,
  SPREADSHEET_ID,
  testConnection,
};
