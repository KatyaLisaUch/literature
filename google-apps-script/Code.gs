const SHEETS = {
  teachers: "Teachers",
  students: "Students",
  poems: "Poems",
  grades: "Grades",
  sessions: "Sessions"
};

const HEADERS = {
  Teachers: ["teacherId", "login", "passwordHash", "salt", "teacherName", "active"],
  Students: ["studentId", "teacherId", "fio", "class"],
  Poems: ["poemId", "teacherId", "title", "author", "gradeLevel", "startDate", "endDate"],
  Grades: ["gradeId", "teacherId", "studentId", "poemId", "grade", "updatedAt"],
  Sessions: ["token", "teacherId", "expiresAt"]
};

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    ensureStructure();

    if (body.action === "login") return json(login(body.login, body.password));
    if (body.action === "getData") return json(getDataByToken(body.token));
    if (body.action === "saveAll") return json(saveAll(body.token, body.data));

    return json({ ok: false, error: "Неизвестное действие." });
  } catch (error) {
    return json({ ok: false, error: error.message });
  }
}

function setupInitialData() {
  ensureStructure();
  createTeacher("teacher1", "123456", "Первый учитель");
}

function createTeacher(login, password, teacherName) {
  const sheet = getSheet(SHEETS.teachers);
  const rows = values(sheet);
  const normalizedLogin = text(login).toLowerCase();
  const existing = rows.slice(1).find((row) => text(row[1]).toLowerCase() === normalizedLogin);

  if (existing) {
    throw new Error("Учитель с таким логином уже существует.");
  }

  const teacherId = id("teacher");
  const salt = id("salt");
  sheet.appendRow([
    teacherId,
    normalizedLogin,
    hashPassword(password, salt),
    salt,
    teacherName || login,
    "TRUE"
  ]);
  return teacherId;
}

function login(loginValue, password) {
  const loginText = text(loginValue).toLowerCase();
  const rows = values(getSheet(SHEETS.teachers));
  const row = rows.slice(1).find((item) => text(item[1]).toLowerCase() === loginText && text(item[5]).toUpperCase() !== "FALSE");

  if (!row) {
    throw new Error("Неверный логин или пароль.");
  }

  const passwordHash = hashPassword(password, row[3]);
  if (passwordHash !== row[2]) {
    throw new Error("Неверный логин или пароль.");
  }

  const token = id("session");
  const expiresAt = new Date(Date.now() + 1000 * 60 * 60 * 24 * 30).toISOString();
  getSheet(SHEETS.sessions).appendRow([token, row[0], expiresAt]);

  return {
    ok: true,
    token,
    teacherId: row[0],
    teacherName: row[4],
    data: readTeacherData(row[0])
  };
}

function getDataByToken(token) {
  const teacherId = requireTeacherId(token);
  return {
    ok: true,
    data: readTeacherData(teacherId)
  };
}

function saveAll(token, data) {
  const teacherId = requireTeacherId(token);
  replaceTeacherRows(SHEETS.students, 1, teacherId, studentRows(teacherId, data && data.students));
  replaceTeacherRows(SHEETS.poems, 1, teacherId, poemRows(teacherId, data && data.poems));
  replaceTeacherRows(SHEETS.grades, 1, teacherId, gradeRows(teacherId, data && data.grades));
  return { ok: true, savedAt: new Date().toISOString() };
}

function readTeacherData(teacherId) {
  const students = values(getSheet(SHEETS.students)).slice(1)
    .filter((row) => row[1] === teacherId)
    .map((row) => ({ id: row[0], name: row[2], className: row[3] }));

  const poems = values(getSheet(SHEETS.poems)).slice(1)
    .filter((row) => row[1] === teacherId)
    .map((row) => ({
      id: row[0],
      title: row[2],
      author: row[3],
      gradeLevel: row[4],
      startDate: row[5],
      endDate: row[6]
    }));

  const grades = {};
  values(getSheet(SHEETS.grades)).slice(1)
    .filter((row) => row[1] === teacherId)
    .forEach((row) => {
      grades[`${row[2]}:${row[3]}`] = text(row[4]);
    });

  return { students, poems, grades };
}

function requireTeacherId(token) {
  const tokenText = text(token);
  if (!tokenText) throw new Error("Нет токена входа.");

  const now = new Date();
  const row = values(getSheet(SHEETS.sessions)).slice(1)
    .find((item) => item[0] === tokenText && new Date(item[2]) > now);

  if (!row) {
    throw new Error("Сессия истекла. Войдите заново.");
  }

  return row[1];
}

function studentRows(teacherId, students) {
  return Array.isArray(students) ? students
    .filter((student) => text(student.name) && text(student.className))
    .map((student) => [text(student.id) || id("student"), teacherId, text(student.name), text(student.className)])
    : [];
}

function poemRows(teacherId, poems) {
  return Array.isArray(poems) ? poems
    .filter((poem) => text(poem.title) && text(poem.author) && text(poem.gradeLevel))
    .map((poem) => [
      text(poem.id) || id("poem"),
      teacherId,
      text(poem.title),
      text(poem.author),
      text(poem.gradeLevel),
      text(poem.startDate),
      text(poem.endDate)
    ])
    : [];
}

function gradeRows(teacherId, grades) {
  if (!grades || typeof grades !== "object") return [];
  const updatedAt = new Date().toISOString();
  return Object.keys(grades)
    .filter((key) => ["2", "3", "4", "5"].includes(text(grades[key])) && key.includes(":"))
    .map((key) => {
      const parts = key.split(":");
      return [`${teacherId}_${parts[0]}_${parts[1]}`, teacherId, parts[0], parts[1], text(grades[key]), updatedAt];
    });
}

function replaceTeacherRows(sheetName, teacherIdColumnIndex, teacherId, newRows) {
  const sheet = getSheet(sheetName);
  const all = values(sheet);
  const header = all[0] || HEADERS[sheetName];
  const remaining = all.slice(1).filter((row) => row[teacherIdColumnIndex] !== teacherId);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  const rows = remaining.concat(newRows);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows.map((row) => fitRow(row, header.length)));
  }
}

function ensureStructure() {
  Object.keys(HEADERS).forEach((sheetName) => {
    const sheet = getSheet(sheetName);
    const headers = HEADERS[sheetName];
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    } else {
      const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
      if (current.join("") === "") {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      }
    }
    sheet.setFrozenRows(1);
  });
}

function getSheet(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
}

function values(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  return sheet.getRange(1, 1, lastRow, lastColumn).getValues();
}

function hashPassword(password, salt) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    `${salt}:${password}`,
    Utilities.Charset.UTF_8
  );
  return bytes.map((byte) => (byte < 0 ? byte + 256 : byte).toString(16).padStart(2, "0")).join("");
}

function id(prefix) {
  return `${prefix}_${Date.now()}_${Math.random().toString(16).slice(2)}`;
}

function text(value) {
  return String(value == null ? "" : value).trim();
}

function fitRow(row, length) {
  const result = row.slice(0, length);
  while (result.length < length) result.push("");
  return result;
}

function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
