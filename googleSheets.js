const { google } = require("googleapis");
const path = require("path");
const fs = require("fs");

let auth;
let sheets;

const SPREADSHEET_ID = "1hjV7t0kgcMnQfHOSM55lCJeIk8zdBrD_7tKAxvEuD38";

try {
  let credentials;

  if (process.env.GOOGLE_CREDENTIALS) {
    console.log("📂 Loading credentials from Render environment variable");
    credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS);
  } else {
    const credPath = path.join(__dirname, "credentials.json");

    if (fs.existsSync(credPath)) {
      console.log("📂 Loading credentials from local credentials.json");
      credentials = require(credPath);
    } else {
      throw new Error(
        "Missing Google credentials. Create .env with GOOGLE_CREDENTIALS or add credentials.json from your service account key."
      );
    }
  }

  auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });

  sheets = google.sheets({
    version: "v4",
    auth,
  });

  console.log("✅ Google Sheets authentication initialized");

} catch (error) {
  console.error("❌ Failed to initialize Google Sheets:", error.message);
  console.error("💡 Fix: add BOT_TOKEN and GOOGLE_CREDENTIALS in .env, or place your service-account JSON in credentials.json.");
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

      console.warn("\n⚠️ PERMISSION ERROR - Share the Google Sheet with:");
      console.warn(
        "📧 attendance-bot@attendance-bot-496018.iam.gserviceaccount.com"
      );

      console.warn(
        "Link: https://docs.google.com/spreadsheets/d/1hjV7t0kgcMnQfHOSM55lCJeIk8zdBrD_7tKAxvEuD38"
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