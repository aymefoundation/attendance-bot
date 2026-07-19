require("dotenv").config();
const http = require("http");

const TelegramBot = require("node-telegram-bot-api");
const { sheets, SPREADSHEET_ID, testConnection } = require("./googleSheets");

// Web server
const server = http.createServer((req, res) => {
  res.writeHead(200, { "Content-Type": "text/plain" });
  res.end("Attendance Bot is active and running!\n");
});

const PORT = process.env.PORT || 10000;
server.listen(PORT, () => {
  console.log(`🚀 Web server active on port ${PORT}`);
});

// Bot initialization
const bot = new TelegramBot(process.env.BOT_TOKEN, {
  polling: { interval: 300, autoStart: true },
});

console.log("🚀 Attendance Bot running (Google Sheets)");

// =======================
// ADMIN USERNAMES
// =======================
const ADMIN_USERNAMES = ["guided_soulll", "ibnu_umeyr"];

let attendanceOpen = false;
let todayColumnIndex = null;

// User registration state (for multi-step registration)
const userRegistrationState = {};

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
// EXPAND SHEET GRID
// =======================
async function expandSheetGrid() {
  try {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [
          {
            updateSheetProperties: {
              properties: {
                sheetId: 0,
                gridProperties: {
                  rowCount: 1000,
                  columnCount: 500,
                },
              },
              fields: "gridProperties(columnCount,rowCount)",
            },
          },
        ],
      },
    });
    console.log("✅ Sheet grid expanded to 1000 rows × 500 columns");
  } catch (error) {
    console.warn("⚠️ Could not expand grid (might already be max):", error.message);
  }
}

// =======================
// INIT SHEET
// =======================
async function initSheet() {
  await expandSheetGrid();

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "Sheet1!1:1",
  });

  const headers = res.data.values?.[0];

  const requiredHeaders = [
    "ID",
    "Name",
    "Gender",
    "Username",
    "TOTAL_PRESENT",
    "TOTAL_ABSENT",
    "PERCENTAGE",
  ];

  if (!headers || headers.length === 0 || headers[0] !== "ID") {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: "Sheet1!A1:G1",
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
        [{text: "✅ Mark Permission (/mark_permission ID)"}],

        [{text: "🔄 Attendance Recycle"}],
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
    range: "Sheet1!A2:D1000",
  });

  return res.data.values || [];
}

// =======================
// REORDER STUDENTS ALPHABETICALLY
// =======================
async function reorderStudentsAlphabetically() {
  try {
    const students = await getStudents();
    
    if (students.length === 0) return;

    // Sort students by name (A-Z)
    const sorted = students.sort((a, b) => {
      const nameA = (a[1] || "").toLowerCase();
      const nameB = (b[1] || "").toLowerCase();
      return nameA.localeCompare(nameB);
    });

    // Reassign IDs (01, 02, 03, etc.)
    const updatedStudents = sorted.map((student, index) => {
      const id = String(index + 1).padStart(2, "0");
      return [id, student[1], student[2], student[3]];
    });

    // Update all rows in the sheet
    const startRow = 2;
    const endRow = startRow + updatedStudents.length;
    const range = `Sheet1!A${startRow}:D${endRow}`;

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: range,
      valueInputOption: "RAW",
      requestBody: { values: updatedStudents },
    });

    console.log(`✅ Students reordered alphabetically. Total: ${updatedStudents.length}`);
  } catch (error) {
    console.error("❌ Error reordering students:", error.message);
  }
}

// =======================
// UPDATE STATS
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
      if (!ADMIN_USERNAMES.includes(username)) {
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
      if (!ADMIN_USERNAMES.includes(username)) return;
      return handleOpen(msg);

    case "✅ Mark Permission (/mark_permission ID)":
      if (!ADMIN_USERNAMES.includes(username)) return;
      return bot.sendMessage(chatId, "Use: /mark_permission ID\nExample: /mark_permission 1");

    case "🔄 Attendance Recycle":
      if (!ADMIN_USERNAMES.includes(username)) return;
      return handleAttendanceRecycle(msg);

    case "🔒 Close Attendance":
      if (!ADMIN_USERNAMES.includes(username)) return;
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
  const chatId = msg.chat.id;
  const username = (msg.from.username || "").toLowerCase();

  try {
    const students = await getStudents();

    if (students.find((s) => s[3] === username)) {
      return bot.sendMessage(chatId, "⚠️ Already registered");
    }

    // Store registration state (id will be assigned after reordering)
    userRegistrationState[chatId] = { name, username };

    // Ask for gender
    bot.sendMessage(chatId, "👤 Select your gender:", {
      reply_markup: {
        keyboard: [[{ text: "👨 Male" }], [{ text: "👩 Female" }]],
        resize_keyboard: true,
      },
    });
  } catch (error) {
    console.error("❌ Error in register:", error.message);
    bot.sendMessage(chatId, `❌ Error: ${error.message}`);
  }
});

// Handle gender selection during registration
const originalMessageHandler = bot.on("message", async (msg) => {
  const text = msg.text;
  const chatId = msg.chat.id;

  if ((text === "👨 Male" || text === "👩 Female") && userRegistrationState[chatId]) {
    try {
      const { name, username } = userRegistrationState[chatId];
      const gender = text === "👨 Male" ? "Male" : "Female";

      // Add student with temporary ID (will be reassigned)
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: "Sheet1!A:D",
        valueInputOption: "RAW",
        requestBody: {
          values: [["0", name, gender, username]],
        },
      });

      // Reorder all students alphabetically and reassign IDs
      await reorderStudentsAlphabetically();

      // Get updated student list to show the new ID
      const updatedStudents = await getStudents();
      const newStudent = updatedStudents.find((s) => s[3] === username);
      const newId = newStudent ? newStudent[0] : "0";

      delete userRegistrationState[chatId];
      bot.sendMessage(chatId, `✅ Registered\n👤 ${name}\n⚤ ${gender}\n🆔 ${newId}`);
      sendStudentMenu(chatId);
    } catch (error) {
      console.error("❌ Error in gender selection:", error.message);
      bot.sendMessage(chatId, `❌ Error: ${error.message}`);
      delete userRegistrationState[chatId];
    }
  }
});

// =======================
// OPEN ATTENDANCE
// =======================
async function handleOpen(msg) {
  try {
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

    bot.sendMessage(msg.chat.id, `📅 Attendance opened for ${today}\n✅ All students marked absent. Send /present to mark present.`);
  } catch (error) {
    console.error("❌ Error in handleOpen:", error.message);
    bot.sendMessage(msg.chat.id, `❌ Error opening attendance: ${error.message}`);
  }
}

// =======================
// PRESENT
// =======================
async function handlePresent(msg) {
  try {
    if (!attendanceOpen) {
      return bot.sendMessage(msg.chat.id, "⛔ Attendance is closed. Admin needs to open it.");
    }

    const username = (msg.from.username || "").toLowerCase();
    const students = await getStudents();

    const index = students.findIndex((s) => s[3] === username);

    if (index === -1) {
      return bot.sendMessage(msg.chat.id, "❌ Not registered. Use /register YourName");
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
  } catch (error) {
    console.error("❌ Error in handlePresent:", error.message);
    bot.sendMessage(msg.chat.id, `❌ Error marking present: ${error.message}`);
  }
}

// =======================
// STATUS
// =======================
async function handleStatus(msg) {
  try {
    const username = (msg.from.username || "").toLowerCase();
    const students = await getStudents();

    const index = students.findIndex((s) => s[3] === username);

    if (index === -1) {
      return bot.sendMessage(msg.chat.id, "❌ Not registered");
    }

    const row = index + 2;

    const headerRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Sheet1!1:1",
    });

    const headers = headerRes.data.values?.[0] || [];

    const presentCol = headers.indexOf("TOTAL_PRESENT");
    const absentCol = headers.indexOf("TOTAL_ABSENT");
    const percentCol = headers.indexOf("PERCENTAGE");

    const dataRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `Sheet1!${getCol(presentCol)}${row}:${getCol(percentCol)}${row}`,
    });

    const [present, absent, percentage] = dataRes.data.values?.[0] || [];

    const percentValue = parseInt(percentage) || 0;

    let text = `📊 STATUS\n\n👤 ${students[index][1]}\n✔ Present: ${present || 0}\n❌ Absent: ${absent || 0}\n📈 Percentage: ${percentage || "0%"}`;

    if (percentValue < 70) {
      text += `\n\n⚠️ ALERT: Low attendance (${percentage})`;
    } else {
      text += `\n\n👍 Good attendance (${percentage})`;
    }

    bot.sendMessage(msg.chat.id, text);
  } catch (error) {
    console.error("❌ Error in handleStatus:", error.message);
    bot.sendMessage(msg.chat.id, `❌ Error getting status: ${error.message}`);
  }
}

// =======================
// CLOSE
// =======================
async function handleClose(msg) {
  try {
    attendanceOpen = false;
    bot.sendMessage(msg.chat.id, "🔒 Attendance closed");
  } catch (error) {
    console.error("❌ Error in handleClose:", error.message);
    bot.sendMessage(msg.chat.id, `❌ Error closing attendance: ${error.message}`);
  }
}

// =======================
// ATTENDANCE RECYCLE (Clear all attendance, reset stats - OPTIMIZED)
// =======================
async function handleAttendanceRecycle(msg) {
  try {
    const chatId = msg.chat.id;

    // Send confirmation request
    const confirmMsg = await bot.sendMessage(chatId, "🔄 Are you sure? This will clear ALL attendance data and reset stats.\n⚠️ This cannot be undone!", {
      reply_markup: {
        inline_keyboard: [
          [
            { text: "✅ Yes, Clear All", callback_data: "confirm_recycle_yes" },
            { text: "❌ Cancel", callback_data: "confirm_recycle_no" },
          ],
        ],
      },
    });
  } catch (error) {
    console.error("❌ Error in handleAttendanceRecycle:", error.message);
    bot.sendMessage(msg.chat.id, `❌ Error: ${error.message}`);
  }
}

// Handle recycle confirmation
bot.on("callback_query", async (query) => {
  const data = query.data;
  const chatId = query.message.chat.id;
  const username = (query.from.username || "").toLowerCase();

  if (!ADMIN_USERNAMES.includes(username)) {
    return bot.answerCallbackQuery(query.id, "⛔ Not authorized", true);
  }

  if (data === "confirm_recycle_yes") {
    try {
      // Get all students
      const students = await getStudents();
      const headerRes = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "Sheet1!1:1",
      });

      const headers = headerRes.data.values?.[0] || [];
      const totalPresentIndex = headers.indexOf("TOTAL_PRESENT");
      const totalAbsentIndex = headers.indexOf("TOTAL_ABSENT");
      const percentageIndex = headers.indexOf("PERCENTAGE");

      // Batch updates to reduce API calls
      const requests = [];

      // Clear all attendance columns for all students (H7:ZZ + row number)
      requests.push({
        updateRange: {
          range: `Sheet1!${getCol(7)}2:${getCol(499)}${students.length + 1}`,
          sheetId: 0,
        },
        values: [],
      });

      // Build batch update data for stats reset
      const statsUpdateValues = [];
      for (let i = 0; i < students.length; i++) {
        statsUpdateValues.push([0, 0, "0%"]);
      }

      // Prepare batch update with single call
      const batchRequests = [];
      
      // Use batch clear values instead of individual updates
      await sheets.spreadsheets.values.batchClear({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
          ranges: [`Sheet1!${getCol(7)}2:${getCol(499)}${students.length + 1}`],
        },
      });

      // Reset stats in one batch
      const statUpdates = [];
      for (let i = 0; i < students.length; i++) {
        const row = i + 2;
        statUpdates.push({
          range: `Sheet1!${getCol(totalPresentIndex)}${row}:${getCol(percentageIndex)}${row}`,
          values: [[0, 0, "0%"]],
        });
      }

      // Batch update all stats at once (max 100 per batch)
      for (let j = 0; j < statUpdates.length; j += 100) {
        const batch = statUpdates.slice(j, j + 100);
        await sheets.spreadsheets.values.batchUpdate({
          spreadsheetId: SPREADSHEET_ID,
          requestBody: { data: batch, valueInputOption: "RAW" },
        });
      }

      // Clear date headers
      await sheets.spreadsheets.values.batchClear({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
          ranges: [`Sheet1!${getCol(7)}1:${getCol(499)}1`],
        },
      });

      attendanceOpen = false;
      todayColumnIndex = null;

      await bot.editMessageText("🔄 ✅ Attendance recycled! All attendance data cleared. Student names, usernames, gender, and stat columns preserved.", {
        chat_id: chatId,
        message_id: query.message.message_id,
      });

      bot.sendMessage(chatId, "✅ System is ready for new attendance. Use 📅 Open Attendance to start.");
    } catch (error) {
      console.error("❌ Error in recycle confirmation:", error.message);
      bot.editMessageText(`❌ Error: ${error.message}`, {
        chat_id: chatId,
        message_id: query.message.message_id,
      });
    }
  } else if (data === "confirm_recycle_no") {
    await bot.editMessageText("❌ Recycle cancelled.", {
      chat_id: chatId,
      message_id: query.message.message_id,
    });
  }
});

// =======================
// MARK PRESENT (Admin) - /p ID
// =======================
bot.onText(/^\/p (\d+)/, async (msg, match) => {
  try {
    const username = (msg.from.username || "").toLowerCase();

    if (!ADMIN_USERNAMES.includes(username)) {
      return bot.sendMessage(msg.chat.id, "⛔ Only admin can mark present");
    }

    if (!attendanceOpen) {
      return bot.sendMessage(msg.chat.id, "⛔ Attendance is not open. Open it first.");
    }

    const id = parseInt(match[1]);
    const students = await getStudents();

    const index = students.findIndex((s) => parseInt(s[0]) === id);

    if (index === -1) {
      return bot.sendMessage(msg.chat.id, "❌ Student not found");
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

    const studentName = students[index][1];
    bot.sendMessage(msg.chat.id, `✅ ${studentName} marked as present`);
  } catch (error) {
    console.error("❌ Error in /p command:", error.message);
    bot.sendMessage(msg.chat.id, `❌ Error: ${error.message}`);
  }
});

// =======================
// MARK PERMISSION (Admin) - /mark_permission ID
// =======================
bot.onText(/^\/mark_permission (\d+)/, async (msg, match) => {
  try {
    const username = (msg.from.username || "").toLowerCase();

    if (!ADMIN_USERNAMES.includes(username)) {
      return bot.sendMessage(msg.chat.id, "⛔ Only admin can mark permission");
    }

    if (!attendanceOpen) {
      return bot.sendMessage(msg.chat.id, "⛔ Attendance is not open. Open it first.");
    }

    const id = parseInt(match[1]);
    const students = await getStudents();

    const index = students.findIndex((s) => parseInt(s[0]) === id);

    if (index === -1) {
      return bot.sendMessage(msg.chat.id, "❌ Student not found");
    }

    const row = index + 2;
    const col = getCol(todayColumnIndex);

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `Sheet1!${col}${row}`,
      valueInputOption: "RAW",
      requestBody: { values: [["P"]] },
    });

    await updateStudentStats(row);

    const studentName = students[index][1];
    bot.sendMessage(msg.chat.id, `✅ ${studentName} marked with permission (P)`);
  } catch (error) {
    console.error("❌ Error in /mark_permission command:", error.message);
    bot.sendMessage(msg.chat.id, `❌ Error: ${error.message}`);
  }
});

// =======================
// MARK NOT PRESENT (Admin) - /n ID
// =======================
bot.onText(/^\/n (\d+)/, async (msg, match) => {
  try {
    const username = (msg.from.username || "").toLowerCase();

    if (!ADMIN_USERNAMES.includes(username)) {
      return bot.sendMessage(msg.chat.id, "⛔ Only admin can mark not present");
    }

    if (!attendanceOpen) {
      return bot.sendMessage(msg.chat.id, "⛔ Attendance is not open. Open it first.");
    }

    const id = parseInt(match[1]);
    const students = await getStudents();

    const index = students.findIndex((s) => parseInt(s[0]) === id);

    if (index === -1) {
      return bot.sendMessage(msg.chat.id, "❌ Student not found");
    }

    const row = index + 2;
    const col = getCol(todayColumnIndex);

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `Sheet1!${col}${row}`,
      valueInputOption: "RAW",
      requestBody: { values: [["❌"]] },
    });

    await updateStudentStats(row);

    const studentName = students[index][1];
    bot.sendMessage(msg.chat.id, `❌ ${studentName} marked as not present`);
  } catch (error) {
    console.error("❌ Error in /n command:", error.message);
    bot.sendMessage(msg.chat.id, `❌ Error: ${error.message}`);
  }
});

// =======================
// DELETE
// =======================
bot.onText(/^\/delete (\d+)/, async (msg, match) => {
  try {
    const username = (msg.from.username || "").toLowerCase();

    if (!ADMIN_USERNAMES.includes(username)) {
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
  } catch (error) {
    console.error("❌ Error in delete:", error.message);
    bot.sendMessage(msg.chat.id, `❌ Error: ${error.message}`);
  }
});
