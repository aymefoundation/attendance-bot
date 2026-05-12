require("dotenv").config();

const TelegramBot = require("node-telegram-bot-api");
const { sheets, SPREADSHEET_ID, testConnection } = require("./googleSheets");

const bot = new TelegramBot(process.env.BOT_TOKEN, {
  polling: { interval: 300, autoStart: true },
});

console.log("🚀 Attendance Bot running (Google Sheets)");

const ADMIN_USERNAME = "guided_soulll";

let attendanceOpen = false;
let todayColumnIndex = null;

// =======================
// COLUMN LETTER
// =======================
function getCol(n) {
  let s = "";
  while (n >= 0) {
    s = String.fromCharCode((n % 26) + 65) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}

// =======================
// INIT SHEET
// =======================
async function initSheet() {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "Sheet1!1:1",
  });

  const headers = res.data.values?.[0];

  const requiredHeaders = [
    "ID",
    "Name",
    "Username",
    "TOTAL_PRESENT",
    "TOTAL_ABSENT",
    "PERCENTAGE",
  ];

  if (!headers || headers.length === 0 || headers[0] !== "ID") {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: "Sheet1!A1:F1",
      valueInputOption: "RAW",
      requestBody: {
        values: [requiredHeaders],
      },
    });

    console.log("✅ Headers initialized");
  }
}

initSheet();

// =======================
// TEST CONNECTION
// =======================
(async () => {
  await testConnection();
})();

// =======================
// MENUS
// =======================
function sendMainMenu(chatId) {
  bot.sendMessage(chatId, "✨ Welcome to Attendance Bot", {
    reply_markup: {
      keyboard: [[{ text: "👨‍🎓 Student" }], [{ text: "👨‍💼 Admin" }]],
      resize_keyboard: true,
    },
  });
}

function sendStudentMenu(chatId) {
  bot.sendMessage(chatId, "🎓 Student Panel", {
    reply_markup: {
      keyboard: [
        [{ text: "📝 Register" }],
        [{ text: "✅ Present" }, { text: "📊 Status" }],
        [{ text: "⬅️ Back To Main Menu" }],
      ],
      resize_keyboard: true,
    },
  });
}

function sendAdminMenu(chatId) {
  bot.sendMessage(chatId, "👨‍💼 Admin Panel", {
    reply_markup: {
      keyboard: [
        [{ text: "📅 Open Attendance" }],
        [{ text: "🔒 Close Attendance" }],
        [{ text: "🗑 Delete Student" }],
        [{ text: "⬅️ Back To Main Menu" }],
      ],
      resize_keyboard: true,
    },
  });
}

// =======================
// GET STUDENTS
// =======================
async function getStudents() {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "Sheet1!A2:C1000",
  });

  return res.data.values || [];
}

// =======================
// UPDATE STATS (FIXED)
// =======================
async function updateStudentStats(row) {
  const headerRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "Sheet1!1:1",
  });

  const headers = headerRes.data.values[0];

  const totalPresentIndex = headers.indexOf("TOTAL_PRESENT");
  const totalAbsentIndex = headers.indexOf("TOTAL_ABSENT");
  const percentageIndex = headers.indexOf("PERCENTAGE");

  const rowRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `Sheet1!${getCol(3)}${row}:${getCol(todayColumnIndex)}${row}`,
  });

  const values = rowRes.data.values?.[0] || [];

  let present = 0;
  let absent = 0;

  values.forEach((v) => {
    if (v === "✔") present++;
    if (v === "❌") absent++;
  });

  const total = present + absent;
  const percentage = total === 0 ? 0 : Math.round((present / total) * 100);

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `Sheet1!${getCol(totalPresentIndex)}${row}:${getCol(percentageIndex)}${row}`,
    valueInputOption: "RAW",
    requestBody: {
      values: [[present, absent, `${percentage}%`]],
    },
  });
}

// =======================
// START
// =======================
bot.onText(/\/start/, (msg) => {
  sendMainMenu(msg.chat.id);
});

// =======================
// ROUTER
// =======================
bot.on("message", async (msg) => {
  const text = msg.text;
  if (!text) return;

  const chatId = msg.chat.id;
  const username = (msg.from.username || "").toLowerCase();

  switch (text) {
    case "👨‍🎓 Student":
      return sendStudentMenu(chatId);

    case "👨‍💼 Admin":
      if (username !== ADMIN_USERNAME) {
        return bot.sendMessage(chatId, "⛔ Not authorized");
      }
      return sendAdminMenu(chatId);

    case "⬅️ Back To Main Menu":
      return sendMainMenu(chatId);

    case "📝 Register":
      return bot.sendMessage(chatId, "Use: /register YourName");

    case "📊 Status":
      return handleStatus(msg);

    case "✅ Present":
      return handlePresent(msg);

    case "📅 Open Attendance":
      if (username !== ADMIN_USERNAME) return;
      return handleOpen(msg);

    case "🔒 Close Attendance":
      if (username !== ADMIN_USERNAME) return;
      return handleClose(msg);

    case "🗑 Delete Student":
      return bot.sendMessage(chatId, "Use: /delete ID");
  }
});

// =======================
// REGISTER
// =======================
bot.onText(/^\/register (.+)/, async (msg, match) => {
  const name = match[1].trim();
  const username = (msg.from.username || "").toLowerCase();

  const students = await getStudents();

  if (students.find((s) => s[2] === username)) {
    return bot.sendMessage(msg.chat.id, "⚠️ Already registered");
  }

  const id =
    students.length > 0
      ? Math.max(...students.map((s) => parseInt(s[0]) || 0)) + 1
      : 1;

  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: "Sheet1!A:C",
    valueInputOption: "RAW",
    requestBody: {
      values: [[id, name, username]],
    },
  });

  bot.sendMessage(msg.chat.id, `✅ Registered\n👤 ${name}\n🆔 ${id}`);
});

// =======================
// OPEN ATTENDANCE (FIXED)
// =======================
async function handleOpen(msg) {
  attendanceOpen = true;

  const today = new Date().toISOString().split("T")[0];

  const headersRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "Sheet1!1:1",
  });

  const headers = headersRes.data.values?.[0] || [];

  let existingIndex = headers.findIndex((h) => h === today);

  if (existingIndex !== -1) {
    todayColumnIndex = existingIndex;
    return bot.sendMessage(
      msg.chat.id,
      `📅 Using existing column for ${today}`,
    );
  }

  todayColumnIndex = headers.length;

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `Sheet1!${getCol(todayColumnIndex)}1`,
    valueInputOption: "RAW",
    requestBody: { values: [[today]] },
  });

  const students = await getStudents();

  for (let i = 0; i < students.length; i++) {
    const row = i + 2;

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `Sheet1!${getCol(todayColumnIndex)}${row}`,
      valueInputOption: "RAW",
      requestBody: { values: [["❌"]] },
    });
  }

  bot.sendMessage(msg.chat.id, `📅 Attendance opened: ${today}`);
}

// =======================
// PRESENT (FIXED)
// =======================
async function handlePresent(msg) {
  if (!attendanceOpen) {
    return bot.sendMessage(msg.chat.id, "⛔ Closed");
  }

  const username = (msg.from.username || "").toLowerCase();
  const students = await getStudents();

  const index = students.findIndex((s) => s[2] === username);

  if (index === -1) {
    return bot.sendMessage(msg.chat.id, "❌ Not registered");
  }

  const row = index + 2;
  const col = getCol(todayColumnIndex);

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `Sheet1!${col}${row}`,
    valueInputOption: "RAW",
    requestBody: { values: [["✔"]] },
  });

  await updateStudentStats(row);

  bot.sendMessage(msg.chat.id, "✅ Present marked");
}

// =======================
// STATUS (FIXED)
// =======================
async function handleStatus(msg) {
  const username = (msg.from.username || "").toLowerCase();
  const students = await getStudents();

  const index = students.findIndex((s) => s[2] === username);

  if (index === -1) {
    return bot.sendMessage(msg.chat.id, "❌ Not registered");
  }

  const row = index + 2;

  const headerRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "Sheet1!1:1",
  });

  const headers = headerRes.data.values[0];

  const presentCol = headers.indexOf("TOTAL_PRESENT");
  const absentCol = headers.indexOf("TOTAL_ABSENT");
  const percentCol = headers.indexOf("PERCENTAGE");

  const dataRes = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `Sheet1!${getCol(presentCol)}${row}:${getCol(percentCol)}${row}`,
  });

  const [present, absent, percentage] = dataRes.data.values?.[0] || [];

  const percentValue = parseInt(percentage) || 0;

  let text = `📊 STATUS\n\n👤 ${students[index][1]}\n✔ Present: ${present}\n❌ Absent: ${absent}\n📈 Percentage: ${percentage}`;

 if (percentValue < 70) {
   text += `\n\n⚠️ ALERT: Low attendance (${percentage})`;
 } else {
   text += `\n\n👍 Good attendance (${percentage})`;
 }
  bot.sendMessage(msg.chat.id, text);
}

// =======================
// CLOSE
// =======================
async function handleClose(msg) {
  attendanceOpen = false;
  bot.sendMessage(msg.chat.id, "🔒 Attendance closed");
}

// =======================
// DELETE (FIXED)
// =======================
bot.onText(/^\/delete (\d+)/, async (msg, match) => {
  const username = (msg.from.username || "").toLowerCase();

  if (username !== ADMIN_USERNAME) {
    return bot.sendMessage(msg.chat.id, "⛔ Only admin");
  }

  const id = parseInt(match[1]);
  const students = await getStudents();

  const index = students.findIndex((s) => parseInt(s[0]) === id);

  if (index === -1) {
    return bot.sendMessage(msg.chat.id, "❌ Not found");
  }

  const name = students[index][1];
  const row = index + 2;

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [
        {
          deleteDimension: {
            range: {
              sheetId: 0,
              dimension: "ROWS",
              startIndex: row - 1,
              endIndex: row,
            },
          },
        },
      ],
    },
  });

  bot.sendMessage(msg.chat.id, `🗑 Deleted Successfully\n👤 ${name}\n🆔 ${id}`);
});
