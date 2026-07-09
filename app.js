const STORAGE_KEY = "poetry-tracker-v1";
const CLOUD_SETTINGS_KEY = "poetry-tracker-cloud-settings-v1";
const CLOUD_SESSION_KEY = "poetry-tracker-cloud-session-v1";
const DEFAULT_API_URL = "https://script.google.com/macros/s/AKfycbx7cr37M2vJYYL_4HEAkGHrkj1HoyrMT5QJkcFX91dv77AIce5wt5zLMWQNnAjexRmcDA/exec";

const state = {
  students: [],
  poems: [],
  grades: {}
};

let dataFileHandle = null;
let dataFileWriteInProgress = false;
let dataFileWritePending = false;
let cloudSaveTimer = null;
let cloudSaveInProgress = false;
let cloudSavePending = false;
let suppressCloudSave = false;

const cloud = {
  apiUrl: "",
  token: "",
  teacherId: "",
  login: "",
  teacherName: ""
};

const els = {
  sidebar: document.querySelector(".sidebar"),
  loginForm: document.getElementById("loginForm"),
  loginPanel: document.getElementById("loginPanel"),
  registerForm: document.getElementById("registerForm"),
  registerPanel: document.getElementById("registerPanel"),
  cloudStatus: document.getElementById("cloudStatus"),
  teacherBadge: document.getElementById("teacherBadge"),
  logoutBtn: document.getElementById("logoutBtn"),
  dataFileInput: document.getElementById("dataFileInput"),
  dataFileStatus: document.getElementById("dataFileStatus"),
  loadDataFileBtn: document.getElementById("loadDataFileBtn"),
  saveDataFileBtn: document.getElementById("saveDataFileBtn"),
  studentsFile: document.getElementById("studentsFile"),
  poemsFile: document.getElementById("poemsFile"),
  poemForm: document.getElementById("poemForm"),
  studentSearchInput: document.getElementById("studentSearchInput"),
  studentSearchList: document.getElementById("studentSearchList"),
  studentSearchResults: document.getElementById("studentSearchResults"),
  classFilter: document.getElementById("classFilter"),
  studentFilter: document.getElementById("studentFilter"),
  debtorsClassFilter: document.getElementById("debtorsClassFilter"),
  exportClassFilter: document.getElementById("exportClassFilter"),
  poemsEditBody: document.getElementById("poemsEditBody"),
  journalHead: document.getElementById("journalHead"),
  journalBody: document.getElementById("journalBody"),
  debtorsBody: document.getElementById("debtorsBody"),
  emptyState: document.getElementById("emptyState"),
  menuToggleBtn: document.getElementById("menuToggleBtn"),
  menuCloseBtn: document.getElementById("menuCloseBtn"),
  menuOverlay: document.getElementById("menuOverlay"),
  appMenu: document.getElementById("appMenu"),
  filtersPanel: document.getElementById("filtersPanel"),
  filtersToggleBtn: document.getElementById("filtersToggleBtn"),
  exportWorkbookBtn: document.getElementById("exportWorkbookBtn"),
  exportDebtorsBtn: document.getElementById("exportDebtorsBtn"),
  clearDataBtn: document.getElementById("clearDataBtn")
};

function removeLegacyUi() {
  document.querySelectorAll('[data-tab="poem-report"], [data-tab="poems"]').forEach((element) => element.remove());
  document.getElementById("poemReportTab")?.remove();
  document.getElementById("poemsTab")?.remove();
  document.getElementById("adminSettingsPanel")?.remove();
}

function uid(prefix) {
  return `${prefix}_${Date.now()}_${Math.random().toString(16).slice(2)}`;
}

function normalizeText(value) {
  return String(value ?? "").trim();
}

function normalizeHeader(value) {
  return normalizeText(value).toLowerCase().replace(/\s+/g, " ");
}

function classParallel(className) {
  const match = normalizeText(className).match(/\d+/);
  return match ? match[0] : normalizeText(className);
}

function displayStudentName(fullName) {
  const parts = normalizeText(fullName).split(/\s+/).filter(Boolean);
  if (parts.length <= 2) return parts.join(" ");
  return `${parts[0]} ${parts[1]}`;
}

function formatDate(value) {
  if (!value) return "";
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10);
  }
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return "";
    return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
  }
  const text = normalizeText(value);
  const parts = text.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/);
  if (parts) {
    const year = parts[3].length === 2 ? `20${parts[3]}` : parts[3];
    return `${year}-${parts[2].padStart(2, "0")}-${parts[1].padStart(2, "0")}`;
  }
  return text;
}

function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target.result);
        resolve(XLSX.read(data, { type: "array", cellDates: true }));
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(file);
  });
}

function firstSheetRows(workbook) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { defval: "" });
}

function pick(row, variants) {
  const entries = Object.entries(row);
  for (const variant of variants) {
    const wanted = normalizeHeader(variant);
    const found = entries.find(([key]) => normalizeHeader(key) === wanted);
    if (found) return found[1];
  }
  return "";
}

function save() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  queueConnectedDataFileWrite();
  queueCloudSave();
}

function load() {
  const saved = localStorage.getItem(STORAGE_KEY);
  if (!saved) return;
  try {
    const parsed = JSON.parse(saved);
    state.students = Array.isArray(parsed.students) ? parsed.students : [];
    state.poems = Array.isArray(parsed.poems) ? parsed.poems : [];
    state.grades = parsed.grades && typeof parsed.grades === "object" ? parsed.grades : {};
  } catch {
    localStorage.removeItem(STORAGE_KEY);
  }
}

function dataSnapshot() {
  return {
    version: 1,
    savedAt: new Date().toISOString(),
    students: state.students,
    poems: state.poems,
    grades: state.grades
  };
}

function loadCloudSettings() {
  try {
    const settings = JSON.parse(localStorage.getItem(CLOUD_SETTINGS_KEY) || "{}");
    const session = JSON.parse(localStorage.getItem(CLOUD_SESSION_KEY) || "{}");
    cloud.apiUrl = normalizeText(settings.apiUrl) || DEFAULT_API_URL;
    cloud.token = normalizeText(session.token);
    cloud.teacherId = normalizeText(session.teacherId);
    cloud.login = normalizeText(session.login);
    cloud.teacherName = normalizeText(session.teacherName);
    updateCloudStatus();
  } catch {
    localStorage.removeItem(CLOUD_SETTINGS_KEY);
    localStorage.removeItem(CLOUD_SESSION_KEY);
    cloud.apiUrl = DEFAULT_API_URL;
    updateCloudStatus();
  }
}

function saveCloudSettings() {
  localStorage.setItem(CLOUD_SETTINGS_KEY, JSON.stringify({ apiUrl: cloud.apiUrl }));
  localStorage.setItem(CLOUD_SESSION_KEY, JSON.stringify({
    token: cloud.token,
    teacherId: cloud.teacherId,
    login: cloud.login,
    teacherName: cloud.teacherName
  }));
}

function clearCloudSession() {
  cloud.token = "";
  cloud.teacherId = "";
  cloud.login = "";
  cloud.teacherName = "";
  localStorage.removeItem(CLOUD_SESSION_KEY);
  updateCloudStatus();
}

function updateCloudStatus(message) {
  const teacherLabel = cloud.teacherName || cloud.teacherId;
  els.teacherBadge.hidden = !cloud.token;
  els.teacherBadge.textContent = cloud.token ? `Текущий учитель: ${teacherLabel}` : "";

  if (message) {
    els.cloudStatus.textContent = cloud.token ? `${message} Аккаунт: ${teacherLabel}.` : message;
  } else if (cloud.token) {
    els.cloudStatus.textContent = `Вход выполнен. Аккаунт: ${teacherLabel}.`;
  } else if (cloud.apiUrl) {
    els.cloudStatus.textContent = "Введите логин и пароль учителя.";
  } else {
    els.cloudStatus.textContent = "Google Sheets не подключен.";
  }
  els.logoutBtn.hidden = !cloud.token;
  els.loginPanel.hidden = !!cloud.token;
  els.registerPanel.hidden = !!cloud.token;
  els.sidebar.hidden = !!cloud.token;
  document.body.classList.toggle("is-authenticated", !!cloud.token);
}

async function apiRequest(action, payload = {}) {
  if (!cloud.apiUrl) {
    throw new Error("Сервис Google Sheets не настроен.");
  }

  const response = await fetch(cloud.apiUrl, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action, ...payload })
  });
  const result = await response.json();
  if (!result.ok) {
    throw new Error(result.error || "Ошибка Google Sheets.");
  }
  return result;
}

function applyRemoteData(data) {
  suppressCloudSave = true;
  try {
    applyDataFile(data || {});
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    render();
  } finally {
    suppressCloudSave = false;
  }
}

function queueCloudSave() {
  if (suppressCloudSave || !cloud.token || !cloud.apiUrl) return;
  clearTimeout(cloudSaveTimer);
  cloudSaveTimer = setTimeout(saveCloudNow, 700);
}

async function saveCloudNow() {
  if (!cloud.token || !cloud.apiUrl) return;
  if (cloudSaveInProgress) {
    cloudSavePending = true;
    return;
  }

  cloudSaveInProgress = true;
  updateCloudStatus("Сохранение в Google Sheets...");
  try {
    await apiRequest("saveAll", {
      token: cloud.token,
      data: dataSnapshot()
    });
    updateCloudStatus("Данные сохранены в Google Sheets.");
  } catch (error) {
    updateCloudStatus(`Ошибка сохранения: ${error.message}`);
  } finally {
    cloudSaveInProgress = false;
    if (cloudSavePending) {
      cloudSavePending = false;
      saveCloudNow();
    }
  }
}

async function loginToCloud(formData) {
  cloud.apiUrl = cloud.apiUrl || DEFAULT_API_URL;
  const login = normalizeText(formData.get("login"));
  const password = normalizeText(formData.get("password"));

  if (!cloud.apiUrl || !login || !password) {
    alert("Заполните логин и пароль.");
    return;
  }

  localStorage.setItem(CLOUD_SETTINGS_KEY, JSON.stringify({ apiUrl: cloud.apiUrl }));
  updateCloudStatus("Выполняется вход...");

  try {
    const result = await apiRequest("login", { login, password });
    cloud.token = result.token;
    cloud.teacherId = result.teacherId;
    cloud.login = result.login || login;
    cloud.teacherName = result.teacherName;
    saveCloudSettings();
    applyRemoteData(result.data);
    updateCloudStatus(`Вход выполнен: ${cloud.teacherName || login}.`);
    els.loginForm.reset();
  } catch (error) {
    clearCloudSession();
    updateCloudStatus(`Ошибка входа: ${error.message}`);
  }
}

async function registerTeacher(formData) {
  cloud.apiUrl = cloud.apiUrl || DEFAULT_API_URL;
  const teacherName = normalizeText(formData.get("teacherName"));
  const login = normalizeText(formData.get("login"));
  const password = normalizeText(formData.get("password"));
  const adminPassword = normalizeText(formData.get("adminPassword"));

  if (!cloud.apiUrl) {
    alert("Сервис Google Sheets не настроен.");
    return;
  }

  if (!teacherName || !login || !password || !adminPassword) {
    alert("Заполните ФИО учителя, логин, пароль и админ-пароль.");
    return;
  }

  if (password.length < 6) {
    alert("Пароль учителя должен быть не короче 6 символов.");
    return;
  }

  localStorage.setItem(CLOUD_SETTINGS_KEY, JSON.stringify({ apiUrl: cloud.apiUrl }));
  updateCloudStatus("Регистрация учителя...");

  try {
    await apiRequest("registerTeacher", {
      teacherName,
      login,
      password,
      adminPassword
    });
    els.registerForm.reset();
    updateCloudStatus(`Учитель зарегистрирован: ${teacherName}. Теперь можно войти.`);
    alert(`Учитель зарегистрирован. Логин: ${login}`);
  } catch (error) {
    updateCloudStatus(`Ошибка регистрации: ${error.message}`);
  }
}

async function refreshCloudData() {
  if (!cloud.token || !cloud.apiUrl) return;
  try {
    updateCloudStatus("Загрузка данных из Google Sheets...");
    const result = await apiRequest("getData", { token: cloud.token });
    applyRemoteData(result.data);
    updateCloudStatus();
  } catch (error) {
    updateCloudStatus(`Не удалось загрузить Google Sheets: ${error.message}`);
  }
}

function setDataFileStatus(message) {
  els.dataFileStatus.textContent = message;
}

function downloadJsonFile(filename, data) {
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function applyDataFile(data) {
  if (!data || typeof data !== "object") {
    throw new Error("Файл базы поврежден или имеет неверный формат.");
  }

  state.students = Array.isArray(data.students)
    ? data.students.map((student) => ({
        id: normalizeText(student.id) || uid("student"),
        name: normalizeText(student.name),
        className: normalizeText(student.className)
      })).filter((student) => student.name && student.className)
    : [];

  state.poems = Array.isArray(data.poems)
    ? data.poems.map((poem) => ({
        id: normalizeText(poem.id) || uid("poem"),
        title: normalizeText(poem.title),
        author: normalizeText(poem.author),
        gradeLevel: normalizeText(poem.gradeLevel),
        startDate: normalizeText(poem.startDate),
        endDate: normalizeText(poem.endDate)
      })).filter((poem) => poem.title && poem.author && poem.gradeLevel)
    : [];

  state.grades = data.grades && typeof data.grades === "object" && !Array.isArray(data.grades)
    ? Object.fromEntries(Object.entries(data.grades).filter(([, grade]) => ["2", "3", "4", "5"].includes(String(grade))))
    : {};
}

async function readDataFile(file) {
  const text = await file.text();
  applyDataFile(JSON.parse(text));
  save();
  render();
}

async function writeConnectedDataFile() {
  if (!dataFileHandle) return;
  const writable = await dataFileHandle.createWritable();
  await writable.write(JSON.stringify(dataSnapshot(), null, 2));
  await writable.close();
  setDataFileStatus("Файл базы подключен. Изменения сохраняются в JSON-файл.");
}

function queueConnectedDataFileWrite() {
  if (!dataFileHandle || !dataFileHandle.createWritable) return;

  if (dataFileWriteInProgress) {
    dataFileWritePending = true;
    return;
  }

  dataFileWriteInProgress = true;
  writeConnectedDataFile()
    .catch(() => setDataFileStatus("Не удалось автоматически записать файл. Нажмите 'Сохранить базу'."))
    .finally(() => {
      dataFileWriteInProgress = false;
      if (dataFileWritePending) {
        dataFileWritePending = false;
        queueConnectedDataFileWrite();
      }
    });
}

async function saveDataFile() {
  if (window.showSaveFilePicker) {
    try {
      dataFileHandle = await window.showSaveFilePicker({
        suggestedName: "baza-ucheta-stihov.json",
        types: [{
          description: "JSON база учета стихов",
          accept: { "application/json": [".json"] }
        }]
      });
      await writeConnectedDataFile();
      return;
    } catch (error) {
      if (error.name === "AbortError") return;
      setDataFileStatus("Не удалось выбрать файл. База будет скачана обычным файлом.");
    }
  }

  downloadJsonFile("baza-ucheta-stihov.json", dataSnapshot());
  setDataFileStatus("База скачана JSON-файлом. После изменений скачайте ее снова.");
}

async function loadDataFile() {
  if (window.showOpenFilePicker) {
    try {
      const [handle] = await window.showOpenFilePicker({
        multiple: false,
        types: [{
          description: "JSON база учета стихов",
          accept: { "application/json": [".json"] }
        }]
      });
      const file = await handle.getFile();
      await readDataFile(file);
      dataFileHandle = handle;
      setDataFileStatus("Файл базы подключен. Изменения сохраняются в JSON-файл.");
      alert("База загружена.");
      return;
    } catch (error) {
      if (error.name === "AbortError") return;
      alert(`Не удалось загрузить базу: ${error.message}`);
      return;
    }
  }

  els.dataFileInput.click();
}

function gradeKey(studentId, poemId) {
  return `${studentId}:${poemId}`;
}

function gradeFor(studentId, poemId) {
  return state.grades[gradeKey(studentId, poemId)] || "";
}

function setGrade(studentId, poemId, grade) {
  const key = gradeKey(studentId, poemId);
  if (grade) {
    state.grades[key] = grade;
  } else {
    delete state.grades[key];
  }
  save();
  render();
}

function selectedClass() {
  return els.classFilter.value;
}

function selectedStudentId() {
  return els.studentFilter.value;
}

function selectedDebtorsClass() {
  return els.debtorsClassFilter.value;
}

function selectedExportClass() {
  return els.exportClassFilter.value;
}

function studentSearchQuery() {
  return normalizeText(els.studentSearchInput.value).toLowerCase();
}

function studentMatchesSearch(student, query) {
  if (!query) return true;
  const fullName = normalizeText(student.name).toLowerCase();
  const displayName = displayStudentName(student.name).toLowerCase();
  const words = fullName.split(/\s+/).filter(Boolean);
  return fullName.startsWith(query)
    || displayName.startsWith(query)
    || words.some((word) => word.startsWith(query));
}

function studentSearchMatches(query) {
  if (!query) return [];
  return state.students
    .filter((student) => studentMatchesSearch(student, query))
    .sort((a, b) => a.name.localeCompare(b.name, "ru"))
    .slice(0, 8);
}

function filteredStudents() {
  const className = selectedClass();
  const studentId = selectedStudentId();
  const query = studentSearchQuery();
  return state.students
    .filter((student) => query || !className || student.className === className)
    .filter((student) => !studentId || student.id === studentId)
    .filter((student) => studentMatchesSearch(student, query))
    .sort((a, b) => a.className.localeCompare(b.className, "ru") || a.name.localeCompare(b.name, "ru"));
}

function filteredPoemsForJournal() {
  const className = selectedClass();
  const parallel = classParallel(className);
  return state.poems
    .filter((poem) => !className || poem.gradeLevel === parallel)
    .sort((a, b) => a.gradeLevel.localeCompare(b.gradeLevel, "ru") || a.title.localeCompare(b.title, "ru"));
}

function allDebts() {
  const rows = [];
  for (const student of state.students) {
    const parallel = classParallel(student.className);
    const poems = state.poems.filter((poem) => poem.gradeLevel === parallel);
    for (const poem of poems) {
      if (!gradeFor(student.id, poem.id)) {
        rows.push({ student, poem });
      }
    }
  }
  return rows;
}

function debtorSummaryRows(className = "") {
  const grouped = new Map();
  for (const { student, poem } of allDebts()) {
    if (className && student.className !== className) continue;
    if (!grouped.has(student.id)) {
      grouped.set(student.id, {
        student,
        poems: []
      });
    }
    grouped.get(student.id).poems.push(poem);
  }
  return [...grouped.values()]
    .sort((a, b) => a.student.className.localeCompare(b.student.className, "ru") || a.student.name.localeCompare(b.student.name, "ru"));
}

function syncFilters() {
  const currentClass = selectedClass();
  const currentStudent = selectedStudentId();
  const query = studentSearchQuery();
  const classes = [...new Set(state.students.map((student) => student.className).filter(Boolean))]
    .sort((a, b) => a.localeCompare(b, "ru"));

  els.classFilter.innerHTML = `<option value="">Все классы</option>${classes
    .map((className) => `<option value="${escapeHtml(className)}">${escapeHtml(className)}</option>`)
    .join("")}`;
  els.classFilter.value = classes.includes(currentClass) ? currentClass : "";

  const currentDebtorsClass = selectedDebtorsClass();
  const currentExportClass = selectedExportClass();
  els.debtorsClassFilter.innerHTML = `<option value="">Все классы</option>${classes
    .map((className) => `<option value="${escapeHtml(className)}">${escapeHtml(className)}</option>`)
    .join("")}`;
  els.debtorsClassFilter.value = classes.includes(currentDebtorsClass) ? currentDebtorsClass : "";

  els.exportClassFilter.innerHTML = `<option value="">Выберите класс</option>${classes
    .map((className) => `<option value="${escapeHtml(className)}">${escapeHtml(className)}</option>`)
    .join("")}`;
  els.exportClassFilter.value = classes.includes(currentExportClass) ? currentExportClass : "";

  const classStudents = state.students
    .filter((student) => query || !els.classFilter.value || student.className === els.classFilter.value)
    .filter((student) => studentMatchesSearch(student, query))
    .sort((a, b) => a.name.localeCompare(b.name, "ru"));
  els.studentFilter.innerHTML = `<option value="">Все ученики</option>${classStudents
    .map((student) => `<option value="${student.id}">${escapeHtml(displayStudentName(student.name))}</option>`)
    .join("")}`;
  els.studentFilter.value = classStudents.some((student) => student.id === currentStudent) ? currentStudent : "";

  const searchMatches = studentSearchMatches(query);
  els.studentSearchList.innerHTML = searchMatches
    .map((student) => `<option value="${escapeHtml(displayStudentName(student.name))}">${escapeHtml(student.className)}</option>`)
    .join("");
  els.studentSearchResults.hidden = searchMatches.length === 0;
  els.studentSearchResults.innerHTML = searchMatches
    .map((student) => `
      <button class="search-result" type="button" data-search-student-id="${student.id}">
        <span>${escapeHtml(displayStudentName(student.name))}</span>
        <small>${escapeHtml(student.className)}</small>
      </button>
    `)
    .join("");
}

function escapeHtml(value) {
  return normalizeText(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function renderStats() {
  els.emptyState.hidden = state.students.length > 0 || state.poems.length > 0;
}

function renderJournal() {
  const students = filteredStudents();
  const poems = filteredPoemsForJournal();

  els.journalHead.innerHTML = `
    <tr>
      <th class="student-name">ФИО</th>
      <th>Класс</th>
      ${poems.map((poem) => `<th>${escapeHtml(poem.title)}<br><small>${escapeHtml(poem.author)}</small></th>`).join("")}
    </tr>
  `;

  els.journalBody.innerHTML = students
    .map((student) => `
      <tr>
        <td class="student-name">${escapeHtml(displayStudentName(student.name))}</td>
        <td>${escapeHtml(student.className)}</td>
        ${poems.map((poem) => renderGradeCell(student, poem)).join("")}
      </tr>
    `)
    .join("");
}

function renderGradeCell(student, poem) {
  const value = gradeFor(student.id, poem.id);
  return `
    <td class="${value ? `grade-${value}` : ""}">
      <select class="grade-select" data-student-id="${student.id}" data-poem-id="${poem.id}">
        <option value="">-</option>
        <option value="2"${value === "2" ? " selected" : ""}>2</option>
        <option value="3"${value === "3" ? " selected" : ""}>3</option>
        <option value="4"${value === "4" ? " selected" : ""}>4</option>
        <option value="5"${value === "5" ? " selected" : ""}>5</option>
      </select>
    </td>
  `;
}

function renderDebtors() {
  const rows = debtorSummaryRows(selectedDebtorsClass());

  els.debtorsBody.innerHTML = rows.length
    ? rows.map(({ student, poems }) => `
        <tr>
          <td data-label="ФИО">${escapeHtml(displayStudentName(student.name))}</td>
          <td data-label="Класс">${escapeHtml(student.className)}</td>
          <td data-label="Стихи">${escapeHtml(poems.map((poem) => poem.title).join(", "))}</td>
          <td data-label="Количество">${poems.length}</td>
          <td data-label="Срок">${escapeHtml(poems.map((poem) => poem.endDate).filter(Boolean).join(", "))}</td>
        </tr>
      `).join("")
    : `<tr><td data-label="" colspan="5">Должников нет</td></tr>`;
}

function renderPoemsEditor() {
  els.poemsEditBody.innerHTML = state.poems.length
    ? state.poems
        .slice()
        .sort((a, b) => a.gradeLevel.localeCompare(b.gradeLevel, "ru") || a.title.localeCompare(b.title, "ru"))
        .map((poem) => `
          <tr>
            <td><input data-poem-id="${poem.id}" data-poem-field="title" value="${escapeHtml(poem.title)}"></td>
            <td><input data-poem-id="${poem.id}" data-poem-field="author" value="${escapeHtml(poem.author)}"></td>
            <td><input data-poem-id="${poem.id}" data-poem-field="gradeLevel" value="${escapeHtml(poem.gradeLevel)}"></td>
            <td><input data-poem-id="${poem.id}" data-poem-field="startDate" type="date" value="${escapeHtml(poem.startDate)}"></td>
            <td><input data-poem-id="${poem.id}" data-poem-field="endDate" type="date" value="${escapeHtml(poem.endDate)}"></td>
            <td><button class="small-action danger" data-delete-poem="${poem.id}" type="button">Удалить</button></td>
          </tr>
        `).join("")
    : `<tr><td colspan="6">Список стихов пуст</td></tr>`;
}

function render() {
  syncFilters();
  renderStats();
  renderJournal();
  renderDebtors();
  renderPoemsEditor();
}

function tabPanelId(tabName) {
  return `${tabName.replace(/-([a-z])/g, (_, letter) => letter.toUpperCase())}Tab`;
}

function setActiveTab(tabName) {
  document.body.dataset.activeTab = tabName;
  document.querySelectorAll(".tab").forEach((tab) => {
    tab.classList.toggle("active", tab.dataset.tab === tabName);
  });
  document.querySelectorAll(".tab-panel").forEach((panel) => {
    panel.classList.toggle("active", panel.id === tabPanelId(tabName));
  });
  closeMenu();
}

function openMenu() {
  els.appMenu.hidden = false;
  els.menuOverlay.hidden = false;
  document.body.classList.add("menu-open");
  els.menuToggleBtn.setAttribute("aria-expanded", "true");
}

function closeMenu() {
  els.appMenu.hidden = true;
  els.menuOverlay.hidden = true;
  document.body.classList.remove("menu-open");
  els.menuToggleBtn.setAttribute("aria-expanded", "false");
}

function toggleMenu() {
  if (els.appMenu.hidden) {
    openMenu();
  } else {
    closeMenu();
  }
}

function toggleFilters() {
  const isCollapsed = els.filtersPanel.classList.toggle("collapsed");
  els.filtersToggleBtn.setAttribute("aria-expanded", String(!isCollapsed));
  els.filtersToggleBtn.querySelector("strong").textContent = isCollapsed ? "Показать" : "Скрыть";
}

function chooseStudentFromSearch(studentId) {
  const student = state.students.find((item) => item.id === studentId);
  if (!student) return;
  els.studentSearchInput.value = displayStudentName(student.name);
  els.classFilter.value = student.className;
  syncFilters();
  els.studentFilter.value = student.id;
  render();
}

async function importStudents(file) {
  const workbook = await readWorkbook(file);
  const rows = firstSheetRows(workbook);
  const imported = rows
    .map((row) => ({
      id: uid("student"),
      name: normalizeText(pick(row, ["ФИО", "Фамилия Имя Отчество", "Ученик"])),
      className: normalizeText(pick(row, ["Класс", "класс"]))
    }))
    .filter((student) => student.name && student.className);

  state.students = imported;
  state.grades = {};
  save();
  render();
  alert(`Импортировано учеников: ${imported.length}`);
}

async function importPoems(file) {
  const workbook = await readWorkbook(file);
  const rows = firstSheetRows(workbook);
  const imported = rows
    .map((row) => ({
      id: uid("poem"),
      title: normalizeText(pick(row, ["Название", "название"])),
      author: normalizeText(pick(row, ["Автор", "автор"])),
      gradeLevel: classParallel(pick(row, ["Класс", "класс", "Параллель", "параллель"])),
      startDate: formatDate(pick(row, ["Дата начала сдачи", "Начало сдачи", "начало"])),
      endDate: formatDate(pick(row, ["Дата окончания сдачи", "Окончание сдачи", "окончание"]))
    }))
    .filter((poem) => poem.title && poem.author && poem.gradeLevel);

  state.poems = imported;
  state.grades = {};
  save();
  render();
  alert(`Импортировано стихов: ${imported.length}`);
}

function addPoem(formData) {
  state.poems.push({
    id: uid("poem"),
    title: normalizeText(formData.get("title")),
    author: normalizeText(formData.get("author")),
    gradeLevel: classParallel(formData.get("gradeLevel")),
    startDate: normalizeText(formData.get("startDate")),
    endDate: normalizeText(formData.get("endDate"))
  });
  save();
  render();
}

function deletePoem(id) {
  state.poems = state.poems.filter((poem) => poem.id !== id);
  for (const key of Object.keys(state.grades)) {
    if (key.endsWith(`:${id}`)) delete state.grades[key];
  }
  save();
  render();
}

function updatePoemField(id, field, value) {
  const poem = state.poems.find((item) => item.id === id);
  if (!poem) return;
  if (field === "gradeLevel") {
    poem[field] = classParallel(value);
  } else {
    poem[field] = normalizeText(value);
  }
  save();
  render();
}

function excelBorderStyle() {
  return {
    top: { style: "thin", color: { rgb: "808080" } },
    bottom: { style: "thin", color: { rgb: "808080" } },
    left: { style: "thin", color: { rgb: "808080" } },
    right: { style: "thin", color: { rgb: "808080" } }
  };
}

function mergeCellStyle(...styles) {
  return styles.reduce((result, style) => {
    Object.keys(style).forEach((key) => {
      result[key] = typeof style[key] === "object" && !Array.isArray(style[key])
        ? { ...(result[key] || {}), ...style[key] }
        : style[key];
    });
    return result;
  }, {});
}

function exportWorkbook(className = selectedExportClass()) {
  if (!className) {
    alert("Выберите класс для экспорта.");
    return;
  }

  const parallel = classParallel(className);
  const poems = state.poems
    .slice()
    .filter((poem) => poem.gradeLevel === parallel)
    .sort((a, b) => a.title.localeCompare(b.title, "ru"));
  const students = state.students
    .filter((student) => student.className === className)
    .sort((a, b) => a.name.localeCompare(b.name, "ru"));

  if (!students.length) {
    alert("В выбранном классе нет учеников.");
    return;
  }

  if (!poems.length) {
    alert("Для выбранного класса нет стихов.");
    return;
  }

  const journalRows = [
    [`Класс: ${className}`],
    ["Фамилия Имя", ...poems.map((poem) => `${poem.title}\n${poem.author}`)]
  ];

  for (const student of students) {
    journalRows.push([displayStudentName(student.name), ...poems.map(() => "")]);
  }

  const workbook = XLSX.utils.book_new();
  const journalSheet = XLSX.utils.aoa_to_sheet(journalRows);
  journalSheet["!cols"] = [
    { wch: 24 },
    ...poems.map(() => ({ wch: 22 }))
  ];
  journalSheet["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: poems.length } }
  ];
  journalSheet["!rows"] = [
    { hpt: 26 },
    { hpt: 44 },
    ...students.map(() => ({ hpt: 24 }))
  ];
  const headerStyle = {
    font: { bold: true },
    fill: { patternType: "solid", fgColor: { rgb: "E6F4F9" } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true },
    border: excelBorderStyle()
  };
  const submittedStyle = {
    fill: { patternType: "solid", fgColor: { rgb: "C6EFCE" } },
    border: excelBorderStyle()
  };
  const missingStyle = {
    fill: { patternType: "solid", fgColor: { rgb: "FFC7CE" } },
    border: excelBorderStyle()
  };
  const textCellStyle = {
    border: excelBorderStyle(),
    alignment: { vertical: "center", wrapText: true }
  };
  const titleStyle = {
    font: { bold: true, sz: 14 },
    fill: { patternType: "solid", fgColor: { rgb: "D9EAF7" } },
    alignment: { horizontal: "center", vertical: "center" },
    border: excelBorderStyle()
  };

  for (let column = 0; column <= poems.length; column += 1) {
    const titleAddress = XLSX.utils.encode_cell({ r: 0, c: column });
    if (!journalSheet[titleAddress]) journalSheet[titleAddress] = { t: "s", v: "" };
    journalSheet[titleAddress].s = titleStyle;

    const address = XLSX.utils.encode_cell({ r: 1, c: column });
    journalSheet[address].s = headerStyle;
  }

  students.forEach((student, studentIndex) => {
    const rowIndex = studentIndex + 2;
    const nameAddress = XLSX.utils.encode_cell({ r: rowIndex, c: 0 });
    if (journalSheet[nameAddress]) journalSheet[nameAddress].s = textCellStyle;

    poems.forEach((poem, poemIndex) => {
      const columnIndex = poemIndex + 1;
      const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
      if (!journalSheet[address]) journalSheet[address] = { t: "s", v: "" };
      journalSheet[address].s = gradeFor(student.id, poem.id) ? submittedStyle : missingStyle;
    });
  });

  XLSX.utils.book_append_sheet(workbook, journalSheet, className);
  XLSX.writeFile(workbook, `uchet-stihov-${className}.xlsx`);
}

function exportDebtorsWorkbook(className = selectedDebtorsClass()) {
  const debtors = debtorSummaryRows(className);

  if (!debtors.length) {
    alert("В выбранном классе должников нет.");
    return;
  }

  const headers = [
    "ФИО",
    "Класс",
    "Количество несданных",
    "Несданные стихи"
  ];
  const rows = [
    headers,
    ...debtors.map(({ student, poems }) => [
      student.name,
      student.className,
      poems.length,
      poems.map((poem, index) => {
        const deadline = poem.endDate ? `, срок: ${poem.endDate}` : "";
        return `${index + 1}. ${poem.title} (${poem.author}${deadline})`;
      }).join("\n")
    ])
  ];

  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet(rows);
  sheet["!cols"] = [
    { wch: 34 },
    { wch: 12 },
    { wch: 18 },
    { wch: 70 }
  ];
  for (let column = 0; column < headers.length; column += 1) {
    const address = XLSX.utils.encode_cell({ r: 0, c: column });
    sheet[address].s = {
      font: { bold: true },
      fill: { patternType: "solid", fgColor: { rgb: "E6F4F9" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border: excelBorderStyle()
    };
  }
  debtors.forEach(({ poems }, index) => {
    [0, 1, 2].forEach((columnIndex) => {
      const address = XLSX.utils.encode_cell({ r: index + 1, c: columnIndex });
      if (sheet[address]) {
        sheet[address].s = {
          alignment: { vertical: "top", wrapText: true },
          border: excelBorderStyle()
        };
      }
    });
    const address = XLSX.utils.encode_cell({ r: index + 1, c: 3 });
    if (sheet[address]) {
      sheet[address].s = {
        alignment: { vertical: "top", wrapText: true },
        border: excelBorderStyle()
      };
    }
  });
  sheet["!rows"] = [
    { hpt: 24 },
    ...debtors.map(({ poems }) => ({ hpt: Math.max(28, poems.length * 18) }))
  ];
  XLSX.utils.book_append_sheet(workbook, sheet, "Должники");
  const suffix = className ? `-${className}` : "-vse-klassy";
  XLSX.writeFile(workbook, `dolzhniki-po-stiham${suffix}.xlsx`);
}

els.studentsFile.addEventListener("change", (event) => {
  const file = event.target.files[0];
  if (file) importStudents(file).catch((error) => alert(`Не удалось импортировать учеников: ${error.message}`));
  event.target.value = "";
});

els.poemsFile.addEventListener("change", (event) => {
  const file = event.target.files[0];
  if (file) importPoems(file).catch((error) => alert(`Не удалось импортировать стихи: ${error.message}`));
  event.target.value = "";
});

els.poemForm.addEventListener("submit", (event) => {
  event.preventDefault();
  addPoem(new FormData(event.currentTarget));
  event.currentTarget.reset();
});

els.loginForm.addEventListener("submit", (event) => {
  event.preventDefault();
  loginToCloud(new FormData(event.currentTarget));
});

els.registerForm.addEventListener("submit", (event) => {
  event.preventDefault();
  registerTeacher(new FormData(event.currentTarget));
});

els.logoutBtn.addEventListener("click", () => {
  clearCloudSession();
  updateCloudStatus("Вы вышли из Google Sheets. Локальные данные остались на экране.");
  closeMenu();
});

els.menuToggleBtn.addEventListener("click", toggleMenu);
els.menuCloseBtn.addEventListener("click", closeMenu);
els.menuOverlay.addEventListener("click", closeMenu);
els.filtersToggleBtn.addEventListener("click", toggleFilters);

document.addEventListener("click", (event) => {
  if (event.target.closest(".menu-shell") || event.target.closest(".app-menu")) return;
  closeMenu();
});

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") closeMenu();
});

els.classFilter.addEventListener("change", render);
els.studentFilter.addEventListener("change", render);
els.studentSearchInput.addEventListener("input", render);
els.debtorsClassFilter.addEventListener("change", renderDebtors);
els.exportClassFilter.addEventListener("change", () => {
  els.debtorsClassFilter.value = els.exportClassFilter.value;
  renderDebtors();
});
els.saveDataFileBtn.addEventListener("click", saveDataFile);
els.loadDataFileBtn.addEventListener("click", loadDataFile);
els.exportWorkbookBtn.addEventListener("click", () => exportWorkbook(selectedExportClass()));
els.exportDebtorsBtn.addEventListener("click", () => exportDebtorsWorkbook(selectedDebtorsClass()));
document.querySelectorAll("#appMenu button:not([data-tab])").forEach((button) => {
  button.addEventListener("click", closeMenu);
});
els.dataFileInput.addEventListener("change", async (event) => {
  const file = event.target.files[0];
  if (!file) return;
  try {
    await readDataFile(file);
    setDataFileStatus("База загружена из JSON-файла. После изменений нажмите 'Сохранить базу'.");
    alert("База загружена.");
  } catch (error) {
    alert(`Не удалось загрузить базу: ${error.message}`);
  } finally {
    event.target.value = "";
  }
});
els.clearDataBtn.addEventListener("click", () => {
  if (!confirm("Удалить все данные из этого браузера?")) return;
  state.students = [];
  state.poems = [];
  state.grades = {};
  save();
  render();
});

document.addEventListener("change", (event) => {
  if (!event.target.matches(".grade-select")) return;
  setGrade(event.target.dataset.studentId, event.target.dataset.poemId, event.target.value);
});

document.addEventListener("change", (event) => {
  if (!event.target.matches("[data-poem-field]")) return;
  updatePoemField(event.target.dataset.poemId, event.target.dataset.poemField, event.target.value);
});

document.addEventListener("click", (event) => {
  const id = event.target.dataset.deletePoem;
  if (!id) return;
  if (confirm("Удалить стих и связанные оценки?")) deletePoem(id);
});

document.addEventListener("click", (event) => {
  const button = event.target.closest("[data-search-student-id]");
  if (!button) return;
  chooseStudentFromSearch(button.dataset.searchStudentId);
});

document.querySelectorAll(".tab").forEach((button) => {
  button.addEventListener("click", () => {
    setActiveTab(button.dataset.tab);
  });
});

document.querySelectorAll(".menu-item[data-tab]").forEach((button) => {
  button.addEventListener("click", () => setActiveTab(button.dataset.tab));
});

removeLegacyUi();
load();
loadCloudSettings();
render();
setActiveTab("journal");
refreshCloudData();
