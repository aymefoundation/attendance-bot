require("dotenv").config();
const http = require("http");
const TelegramBot = require("node-telegram-bot-api");
const { sheets, SPREADSHEET_ID, testConnection } = require("./googleSheets");

const CLASS_NAMES = ["Class 1", "Class 2", "Class 3", "Class 4", "Class 5", "Class 6"];
const CLASS_ICONS = ["📘", "📗", "📙", "📕", "📓", "📔"];
const ADMIN_USERNAMES = ["guided_soulll", "ibnu_umeyr", "quiet_lightt"];
const FIXED_HEADERS = [
  "ID", "Name", "Gender", "Username", "Telegram ID", "STATUS",
  "TOTAL_PRESENT", "TOTAL_ABSENT", "TOTAL_PERMISSION", "PERCENTAGE",
];
const registrationState = {};
const adminState = {};
const adminChatIds = new Set();
let sheetIds = {};
let attendanceOpen = false;
let activeClass = null;
let todayColumnIndex = null;

const server = http.createServer((req, res) => {
  res.writeHead(200, { "Content-Type": "text/plain" });
  res.end("Attendance Bot is active and running!\n");
});
server.listen(process.env.PORT || 10000, () => console.log("🚀 Web server active"));
const bot = new TelegramBot(process.env.BOT_TOKEN, { polling: false });
let mutationQueue = Promise.resolve();
function withMutationLock(operation) {
  const result = mutationQueue.then(operation, operation);
  mutationQueue = result.catch(() => undefined);
  return result;
}

function isAdmin(msg) { return ADMIN_USERNAMES.includes((msg.from.username || "").toLowerCase()); }
function getCol(index) {
  let result = "";
  while (index >= 0) { result = String.fromCharCode(index % 26 + 65) + result; index = Math.floor(index / 26) - 1; }
  return result;
}
function classKeyboard(prefix) {
  return {
    inline_keyboard: CLASS_NAMES.map((name, index) => [{ text: `${CLASS_ICONS[index]} ${name}`, callback_data: `${prefix}:${index}` }]),
  };
}
function classFromCallback(data, prefix) {
  const match = new RegExp(`^${prefix}:(\\d+)$`).exec(data);
  return match && CLASS_NAMES[Number(match[1])];
}
function classPrompt(chatId, prefix, text) {
  return bot.sendMessage(chatId, text, { reply_markup: classKeyboard(prefix) });
}
function clearRegistrationState(chatId, className) {
  if (!registrationState[chatId] || !className || registrationState[chatId].className === className) delete registrationState[chatId];
}
function range(sheetName, cellRange) { return `'${sheetName}'!${cellRange}`; }
function belongsToUser(row, msg) {
  const telegramId = String(msg.from.id);
  const username = (msg.from.username || "").toLowerCase();
  return row[4] === telegramId || (!row[4] && username && (row[3] || "").toLowerCase() === username);
}

async function setupSheets() {
  const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const existing = spreadsheet.data.sheets || [];
  const byTitle = Object.fromEntries(existing.map((sheet) => [sheet.properties.title, sheet]));
  if (!byTitle["Class 1"] && byTitle.Sheet1) {
    await sheets.spreadsheets.batchUpdate({ spreadsheetId: SPREADSHEET_ID, requestBody: { requests: [{ updateSheetProperties: { properties: { sheetId: byTitle.Sheet1.properties.sheetId, title: "Class 1" }, fields: "title" } }] } });
    byTitle["Class 1"] = { properties: { sheetId: byTitle.Sheet1.properties.sheetId, title: "Class 1" } };
  }
  const requests = CLASS_NAMES.filter((name) => !byTitle[name]).map((title) => ({ addSheet: { properties: { title, gridProperties: { rowCount: 1000, columnCount: 500 } } } }));
  if (requests.length) await sheets.spreadsheets.batchUpdate({ spreadsheetId: SPREADSHEET_ID, requestBody: { requests } });
  const refreshed = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  sheetIds = Object.fromEntries((refreshed.data.sheets || []).map((sheet) => [sheet.properties.title, sheet.properties.sheetId]));
  for (const className of CLASS_NAMES) await ensureHeaders(className);
  const finalSpreadsheet = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  sheetIds = Object.fromEntries((finalSpreadsheet.data.sheets || []).map((sheet) => [sheet.properties.title, sheet.properties.sheetId]));
}

async function ensureHeaders(className) {
  const response = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: range(className, "1:1") });
  const oldHeaders = response.data.values?.[0] || [];
  if (!oldHeaders.length) {
    await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: range(className, `A1:${getCol(FIXED_HEADERS.length - 1)}1`), valueInputOption: "RAW", requestBody: { values: [FIXED_HEADERS] } });
    return;
  }
  const dates = oldHeaders.filter((header) => /^\d{4}-\d{2}-\d{2}$/.test(header));
  const extraHeaders = oldHeaders.filter((header) => header && !FIXED_HEADERS.includes(header) && !/^\d{4}-\d{2}-\d{2}$/.test(header));
  const canonical = [...FIXED_HEADERS, ...dates, ...extraHeaders];
  if (FIXED_HEADERS.every((header, index) => oldHeaders[index] === header) && oldHeaders.length === canonical.length) return;
  const rows = (await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: range(className, `A2:${getCol(Math.max(oldHeaders.length - 1, 0))}`) })).data.values || [];
  const migrated = rows.map((row) => {
    const values = canonical.map((header) => {
      const oldIndex = oldHeaders.indexOf(header);
      if (oldIndex !== -1) return row[oldIndex] || "";
      if (header === "Telegram ID") return "";
      if (header === "STATUS") return "APPROVED";
      return header === "PERCENTAGE" ? "0%" : 0;
    });
    return values;
  });
  const temporaryTitle = `${className} migration ${Date.now()}`;
  await sheets.spreadsheets.batchUpdate({ spreadsheetId: SPREADSHEET_ID, requestBody: { requests: [{ addSheet: { properties: { title: temporaryTitle, gridProperties: { rowCount: 1000, columnCount: 500 } } } }] } });
  await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: range(temporaryTitle, `A1:${getCol(canonical.length - 1)}${migrated.length + 1}`), valueInputOption: "RAW", requestBody: { values: [canonical, ...migrated] } });
  const migrationSpreadsheet = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const temporarySheet = migrationSpreadsheet.data.sheets.find((sheet) => sheet.properties.title === temporaryTitle);
  await sheets.spreadsheets.batchUpdate({ spreadsheetId: SPREADSHEET_ID, requestBody: { requests: [
    { deleteSheet: { sheetId: sheetIds[className] } },
    { updateSheetProperties: { properties: { sheetId: temporarySheet.properties.sheetId, title: className }, fields: "title" } },
  ] } });
}

async function getSheetData(className) {
  const values = (await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: range(className, "A:ZZZ") })).data.values || [];
  return { headers: values[0] || [], rows: values.slice(1) };
}
async function findStudent(className, predicate) {
  const data = await getSheetData(className);
  const index = data.rows.findIndex(predicate);
  return index < 0 ? null : { ...data.rows[index], row: index + 2, headers: data.headers };
}
async function updateCell(className, cell, value) {
  await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: range(className, cell), valueInputOption: "RAW", requestBody: { values: [[value]] } });
}
async function updateStats(className, row) {
  const data = await getSheetData(className);
  const values = data.rows[row - 2] || [];
  const attendance = values.slice(FIXED_HEADERS.length);
  const present = attendance.filter((value) => value === "✔").length;
  const absent = attendance.filter((value) => value === "❌").length;
  const permission = attendance.filter((value) => value === "P").length;
  const total = present + absent + permission;
  const percentage = total ? Math.round(present / total * 100) : 0;
  await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: range(className, `G${row}:J${row}`), valueInputOption: "RAW", requestBody: { values: [[present, absent, permission, `${percentage}%`]] } });
}
async function markStudent(className, row, columnIndex, value) {
  return withMutationLock(async () => {
    if (!attendanceOpen || activeClass !== className || todayColumnIndex !== columnIndex) throw new Error("Attendance session changed; please try again.");
    await updateCell(className, `${getCol(columnIndex)}${row}`, value);
    await updateStats(className, row);
  });
}

function sendMainMenu(chatId) { return bot.sendMessage(chatId, "✨ Welcome to Attendance Bot", { reply_markup: { keyboard: [[{ text: "👨‍🎓 Student" }], [{ text: "👨‍💼 Admin" }]], resize_keyboard: true } }); }
function sendStudentMenu(chatId) { return bot.sendMessage(chatId, "🎓 Student Panel", { reply_markup: { keyboard: [[{ text: "📝 Register" }], [{ text: "✅ Present" }, { text: "📊 Status" }], [{ text: "⬅️ Back To Main Menu" }]], resize_keyboard: true } }); }
function sendAdminMenu(chatId) { return bot.sendMessage(chatId, "👨‍💼 Admin Panel", { reply_markup: { keyboard: [[{ text: "📅 Open Attendance" }], [{ text: "🔒 Close Attendance" }], [{ text: "✅ Mark Permission (/mark_permission ID)" }], [{ text: "🔄 Attendance Recycle" }], [{ text: "🗑 Delete Student" }], [{ text: "⬅️ Back To Main Menu" }]], resize_keyboard: true } }); }

bot.onText(/\/start/, (msg) => { delete registrationState[msg.chat.id]; return sendMainMenu(msg.chat.id); });
bot.on("message", async (msg) => {
  const text = msg.text;
  if (!text || text.startsWith("/")) return;
  const chatId = msg.chat.id;
  if (isAdmin(msg)) adminChatIds.add(chatId);
  if (text === "❌ Cancel Registration") { delete registrationState[chatId]; return sendStudentMenu(chatId); }
  if (registrationState[chatId]?.name === null) {
    registrationState[chatId].name = text.trim();
    return classPrompt(chatId, `register:${chatId}:${msg.from.id}`, "📚 Select the class you want to register for:");
  }
  if ((text === "👨 Male" || text === "👩 Female") && registrationState[chatId]) return finishRegistration(msg, text === "👨 Male" ? "Male" : "Female");
  if (adminState[chatId]?.action === "delete-id") return deleteById(msg, text.trim());
  if (text === "👨‍🎓 Student") return sendStudentMenu(chatId);
  if (text === "👨‍💼 Admin") return isAdmin(msg) ? sendAdminMenu(chatId) : bot.sendMessage(chatId, "⛔ Not authorized");
  if (text === "⬅️ Back To Main Menu") { delete registrationState[chatId]; return sendMainMenu(chatId); }
  if (text === "📝 Register") return bot.sendMessage(chatId, "Use /register YourName");
  if (text === "✅ Present") return handlePresent(msg);
  if (text === "📊 Status") return handleStatus(msg);
  if (text === "📅 Open Attendance" && isAdmin(msg)) return classPrompt(chatId, "open", "📚 Which class do you want to open?");
  if (text === "🔒 Close Attendance" && isAdmin(msg)) return handleClose(msg);
  if (text === "🔄 Attendance Recycle" && isAdmin(msg)) return classPrompt(chatId, "recycle", "📚 Which class do you want to recycle?");
  if (text === "🗑 Delete Student" && isAdmin(msg)) return classPrompt(chatId, "delete", "📚 Which class do you want to delete the student from?");
  if (text === "✅ Mark Permission (/mark_permission ID)" && isAdmin(msg)) return bot.sendMessage(chatId, "Use: /mark_permission ID\nExample: /mark_permission 1");
});

bot.onText(/^\/cancel$/, (msg) => { delete registrationState[msg.chat.id]; return sendStudentMenu(msg.chat.id); });
bot.onText(/^\/register(?:\s+(.+))?$/, async (msg, match) => {
  const name = match[1]?.trim();
  if (!name) { registrationState[msg.chat.id] = { name: null, startedAt: Date.now() }; return bot.sendMessage(msg.chat.id, "👤 Enter your full name:"); }
  registrationState[msg.chat.id] = { name, startedAt: Date.now() };
  return classPrompt(msg.chat.id, `register:${msg.chat.id}:${msg.from.id}`, "📚 Select the class you want to register for:");
});
bot.onText(/^\/delete\s+(\d+)$/, (msg, match) => deleteById(msg, match[1]));

async function finishRegistration(msg, gender) {
  return withMutationLock(async () => {
    const state = registrationState[msg.chat.id];
  if (state?.startedAt && Date.now() - state.startedAt > 15 * 60 * 1000) { delete registrationState[msg.chat.id]; return sendStudentMenu(msg.chat.id); }
  if (!state?.name) { delete registrationState[msg.chat.id]; return bot.sendMessage(msg.chat.id, "Please use /register YourName first."); }
  const username = (msg.from.username || "").toLowerCase();
  const className = state.className;
  if (!className) { delete registrationState[msg.chat.id]; return bot.sendMessage(msg.chat.id, "Please select a class first."); }
  const existing = await findStudent(className, (student) => belongsToUser(student, msg) && ["PENDING", "APPROVED"].includes(student[5]));
  if (existing) { delete registrationState[msg.chat.id]; return bot.sendMessage(msg.chat.id, "⚠️ Already registered or awaiting approval for this class."); }
  const data = await getSheetData(className);
  const nextId = String(data.rows.reduce((max, row) => Math.max(max, Number(row[0]) || 0), 0) + 1).padStart(2, "0");
  await sheets.spreadsheets.values.append({ spreadsheetId: SPREADSHEET_ID, range: range(className, "A:J"), valueInputOption: "RAW", requestBody: { values: [[nextId, state.name, gender, username, String(msg.from.id), "PENDING", 0, 0, 0, "0%"]] } });
  delete registrationState[msg.chat.id];
  await bot.sendMessage(msg.chat.id, `🕒 Registration pending approval for ${className}.`);
  const notificationTargets = adminChatIds.size ? [...adminChatIds] : ADMIN_USERNAMES.map((admin) => `@${admin}`);
  for (const target of notificationTargets) await bot.sendMessage(target, `🆕 New registration\n👤 ${state.name}\n📚 ${className}\n🆔 ${nextId}`, { reply_markup: { inline_keyboard: [[{ text: "✅ Approve", callback_data: `approval:approve:${className}:${nextId}:${msg.from.id}` }, { text: "❌ Reject", callback_data: `approval:reject:${className}:${nextId}:${msg.from.id}` }]] } });
  });
}

async function handleOpenForClass(chatId, className) {
  return withMutationLock(async () => {
  if (attendanceOpen) return bot.sendMessage(chatId, `⚠️ Attendance is already open for ${activeClass}.\n\nPlease close ${activeClass} before opening another class.`);
  const today = new Date().toISOString().split("T")[0];
  const data = await getSheetData(className);
  let index = data.headers.indexOf(today);
  if (index < 0) { index = data.headers.length; await updateCell(className, `${getCol(index)}1`, today); }
  for (let row = 2; row <= data.rows.length + 1; row++) if (data.rows[row - 2][5] === "APPROVED") await updateCell(className, `${getCol(index)}${row}`, "❌");
  todayColumnIndex = index; activeClass = className; attendanceOpen = true;
  return bot.sendMessage(chatId, `📅 Attendance for ${className} is opened.`);
  });
}
async function handlePresent(msg) {
  if (!attendanceOpen) return bot.sendMessage(msg.chat.id, "⛔ Attendance is closed. Admin needs to open it.");
  const session = { className: activeClass, columnIndex: todayColumnIndex };
  const student = await findStudent(session.className, (row) => belongsToUser(row, msg) && row[5] === "APPROVED");
  if (!student) return bot.sendMessage(msg.chat.id, "⛔ You are not registered for this class.");
  await markStudent(session.className, student.row, session.columnIndex, "✔"); return bot.sendMessage(msg.chat.id, `✅ Present marked for ${session.className}`);
}
async function handleStatus(msg) {
  const classes = [];
  for (const className of CLASS_NAMES) { const student = await findStudent(className, (row) => belongsToUser(row, msg) && row[5] === "APPROVED"); if (student) classes.push({ className, student }); }
  if (!classes.length) return bot.sendMessage(msg.chat.id, "❌ Not registered");
  const lines = classes.map(({ className, student }) => `📚 ${className}\n👤 ${student[1]}\n✔ Present: ${student[6] || 0}\n❌ Absent: ${student[7] || 0}\n🅿️ Permission: ${student[8] || 0}\n📈 Percentage: ${student[9] || "0%"}${Number.parseInt(student[9], 10) < 70 ? "\n⚠️ ALERT: Low attendance" : ""}`);
  return bot.sendMessage(msg.chat.id, `📊 STATUS\n\n${lines.join("\n\n")}`);
}
async function handleClose(msg) { if (!attendanceOpen) return bot.sendMessage(msg.chat.id, "⛔ Attendance is already closed."); attendanceOpen = false; activeClass = null; todayColumnIndex = null; return bot.sendMessage(msg.chat.id, "🔒 Attendance closed"); }

async function recycleClass(chatId, className, messageId) {
  const data = await getSheetData(className);
  if (data.headers.length > FIXED_HEADERS.length) await sheets.spreadsheets.values.batchClear({ spreadsheetId: SPREADSHEET_ID, requestBody: { ranges: [range(className, `${getCol(FIXED_HEADERS.length)}1:${getCol(data.headers.length - 1)}${data.rows.length + 1}`)] } });
  for (let row = 2; row <= data.rows.length + 1; row++) await sheets.spreadsheets.values.update({ spreadsheetId: SPREADSHEET_ID, range: range(className, `G${row}:J${row}`), valueInputOption: "RAW", requestBody: { values: [[0, 0, 0, "0%"]] } });
  if (activeClass === className) { attendanceOpen = false; activeClass = null; todayColumnIndex = null; }
  await bot.editMessageText(`✅ ${className} attendance recycled. Student information was preserved.`, { chat_id: chatId, message_id: messageId });
}
async function deleteById(msg, idText) {
  if (!isAdmin(msg)) return bot.sendMessage(msg.chat.id, "⛔ Only admin");
  const className = adminState[msg.chat.id]?.className;
  if (!className) return classPrompt(msg.chat.id, "delete", "📚 Which class do you want to delete the student from?");
  if (!/^\d+$/.test(idText)) return bot.sendMessage(msg.chat.id, "❌ Student ID must contain only numbers.");
  const student = await findStudent(className, (row) => row[0] === String(Number(idText)).padStart(2, "0"));
  if (!student) return bot.sendMessage(msg.chat.id, "❌ Student not found in this class.");
  adminState[msg.chat.id].student = student; delete adminState[msg.chat.id].action;
  return bot.sendMessage(msg.chat.id, `⚠️ Confirm deletion\n\n👤 Name: ${student[1]}\n🆔 ID: ${student[0]}\n📚 Class: ${className}\n\nAre you sure?`, { reply_markup: { inline_keyboard: [[{ text: "✅ Delete", callback_data: `delete:yes:${className}:${student[0]}:${student[4]}` }, { text: "❌ Cancel", callback_data: `delete:no:${className}:${student[0]}:${student[4]}` }]] } });
}

async function deleteStudentRow(className, student) {
  return withMutationLock(async () => {
    const current = await findStudent(className, (row) => row[0] === student[0] && row[4] === student[4]);
    if (!current || current[1] !== student[1]) return null;
    const sheetId = sheetIds[className];
    if (sheetId === undefined) throw new Error(`Sheet ID unavailable for ${className}`);
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { requests: [{ deleteDimension: { range: { sheetId, dimension: "ROWS", startIndex: current.row - 1, endIndex: current.row } } }] },
    });
    return current;
  });
}

bot.on("callback_query", async (query) => {
  if (!query.message) return bot.answerCallbackQuery(query.id, "Unsupported callback", true);
  const data = query.data; const chatId = query.message.chat.id;
  if (!ADMIN_USERNAMES.includes((query.from.username || "").toLowerCase()) && !data.startsWith("register:")) return bot.answerCallbackQuery(query.id, "⛔ Not authorized", true);
  const registerMatch = /^register:(-?\d+):(-?\d+):(\d+)$/.exec(data);
  const registerClass = registerMatch && Number(registerMatch[1]) === Number(chatId) && Number(registerMatch[2]) === Number(query.from.id) && CLASS_NAMES[Number(registerMatch[3])];
  if (registerClass) {
    const state = registrationState[chatId];
    if (!state || (state.startedAt && Date.now() - state.startedAt > 15 * 60 * 1000)) { delete registrationState[chatId]; return bot.answerCallbackQuery(query.id, "Registration expired. Please start again.", true); }
    state.className = registerClass;
    await bot.editMessageReplyMarkup({ inline_keyboard: [] }, { chat_id: chatId, message_id: query.message.message_id });
    await bot.answerCallbackQuery(query.id);
    return bot.sendMessage(chatId, "👤 Select your gender:", { reply_markup: { keyboard: [[{ text: "👨 Male" }], [{ text: "👩 Female" }], [{ text: "❌ Cancel Registration" }]], resize_keyboard: true } });
  }
  const openClass = classFromCallback(data, "open"); if (openClass) { await bot.answerCallbackQuery(query.id); return handleOpenForClass(chatId, openClass); }
  const recycleClassName = classFromCallback(data, "recycle"); if (recycleClassName) { await bot.answerCallbackQuery(query.id); return bot.sendMessage(chatId, `⚠️ Are you sure?\n\nThis will clear ALL attendance records for ${recycleClassName} and reset its attendance statistics.\n\nStudent basic information will NOT be deleted.\n\nThis action cannot be undone.`, { reply_markup: { inline_keyboard: [[{ text: "✅ Yes, Clear", callback_data: `recycle-confirm:${recycleClassName}` }, { text: "❌ Cancel", callback_data: "recycle-cancel" }]] } }); }
  const deleteClass = classFromCallback(data, "delete"); if (deleteClass) { adminState[chatId] = { className: deleteClass, action: "delete-id" }; await bot.answerCallbackQuery(query.id); return bot.sendMessage(chatId, "🆔 Enter the student ID you want to delete.\n\nExample:\n05"); }
  const approvalMatch = /^approval:(approve|reject):(Class [1-6]):(\d+):(-?\d+)$/.exec(data);
  if (approvalMatch) {
    const [, decision, className, id, studentChatId] = approvalMatch;
    const student = await findStudent(className, (row) => row[0] === id && row[4] === studentChatId && row[5] === "PENDING");
    if (!student) return bot.answerCallbackQuery(query.id, "Registration is no longer pending.", true);
    const approved = decision === "approve";
    await updateCell(className, `F${student.row}`, approved ? "APPROVED" : "REJECTED");
    clearRegistrationState(Number(studentChatId), className);
    await bot.editMessageText(`${approved ? "✅ Approved" : "❌ Rejected"}: ${student[1]} (${className})`, { chat_id: chatId, message_id: query.message.message_id });
    await bot.sendMessage(Number(studentChatId), approved
      ? `✅ Your registration for ${className} has been approved!\n\nYou are now registered for ${className}.\n\nPlease use the Student Panel to access the available features.`
      : `❌ Your registration for ${className} was rejected by the administrator.`, { reply_markup: { remove_keyboard: true } });
    return sendStudentMenu(Number(studentChatId)).then(() => bot.answerCallbackQuery(query.id));
  }
  if (data.startsWith("recycle-confirm:")) { await recycleClass(chatId, data.split(":")[1], query.message.message_id); return bot.answerCallbackQuery(query.id); }
  if (data === "recycle-cancel") { await bot.editMessageText("❌ Recycle cancelled.", { chat_id: chatId, message_id: query.message.message_id }); return bot.answerCallbackQuery(query.id); }
  if (data.startsWith("delete:")) {
    const [, decision, className, id, telegramId] = data.split(":");
    const state = adminState[chatId];
    const confirmedStudent = state?.student;
    if (decision === "yes" && confirmedStudent && className === state.className && id === confirmedStudent[0] && telegramId === confirmedStudent[4]) {
      try {
        const deletedStudent = await deleteStudentRow(className, confirmedStudent);
        if (!deletedStudent) await bot.editMessageText("⚠️ This student record has already changed; nothing was deleted.", { chat_id: chatId, message_id: query.message.message_id });
        else await bot.editMessageText(`🗑 Deleted ${deletedStudent[1]} from ${className}.`, { chat_id: chatId, message_id: query.message.message_id });
      } catch (error) {
        await bot.editMessageText(`❌ Could not delete the student row: ${error.message}`, { chat_id: chatId, message_id: query.message.message_id });
      }
    } else await bot.editMessageText("❌ Deletion cancelled or confirmation expired.", { chat_id: chatId, message_id: query.message.message_id });
    delete adminState[chatId];
    return bot.answerCallbackQuery(query.id);
  }
});

async function adminMark(msg, match, value, label) {
  if (!isAdmin(msg)) return bot.sendMessage(msg.chat.id, "⛔ Only admin can mark attendance");
  if (!attendanceOpen) return bot.sendMessage(msg.chat.id, "⛔ Attendance is not open. Open it first.");
  const session = { className: activeClass, columnIndex: todayColumnIndex };
  const id = String(Number(match[1])).padStart(2, "0"); const student = await findStudent(session.className, (row) => row[0] === id && row[5] === "APPROVED");
  if (!student) return bot.sendMessage(msg.chat.id, "❌ Student not found in the active class");
  await markStudent(session.className, student.row, session.columnIndex, value); return bot.sendMessage(msg.chat.id, `${label} ${student[1]} in ${session.className}`);
}
bot.onText(/^\/p\s+(\d+)$/, (msg, match) => adminMark(msg, match, "✔", "✅ Marked present:"));
bot.onText(/^\/n\s+(\d+)$/, (msg, match) => adminMark(msg, match, "❌", "❌ Marked absent:"));
bot.onText(/^\/mark_permission\s+(\d+)$/, (msg, match) => adminMark(msg, match, "P", "🅿️ Marked permission:"));

(async () => {
  try {
    await setupSheets();
    if (!(await testConnection())) throw new Error("Google Sheets connection failed");
    await bot.startPolling();
    console.log("🚀 Attendance Bot running with six class tabs");
  } catch (error) {
    console.error("❌ Startup failed; Telegram polling was not started:", error.message);
  }
})();
