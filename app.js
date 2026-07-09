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
  loginForm: document.getElementById("loginForm"),
  loginPanel: document.getElementById("loginPanel"),
  adminSettingsForm: document.getElementById("adminSettingsForm"),
  adminSettingsPanel: document.getElementById("adminSettingsPanel"),
  registerForm: document.getElementById("registerForm"),
  registerPanel: document.getElementById("registerPanel"),
  apiUrlInput: document.getElementById("apiUrlInput"),
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
  classFilter: document.getElementById("classFilter"),
  studentFilter: document.getElementById("studentFilter"),
  journalHead: document.getElementById("journalHead"),
  journalBody: document.getElementById("journalBody"),
  debtorsBody: document.getElementById("debtorsBody"),
  poemReportBody: document.getElementById("poemReportBody"),
  poemsBody: document.getElementById("poemsBody"),
  emptyState: document.getElementById("emptyState"),
  menuToggleBtn: document.getElementById("menuToggleBtn"),
  appMenu: document.getElementById("appMenu"),
  exportWorkbookBtn: document.getElementById("exportWorkbookBtn"),
  clearDataBtn: document.getElementById("clearDataBtn")
};

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
    els.apiUrlInput.value = cloud.apiUrl;
    updateCloudStatus();
  } catch {
    localStorage.removeItem(CLOUD_SETTINGS_KEY);
    localStorage.removeItem(CLOUD_SESSION_KEY);
    cloud.apiUrl = DEFAULT_API_URL;
    els.apiUrlInput.value = cloud.apiUrl;
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
  const isAdmin = !!cloud.token && cloud.login.toLowerCase() === "katyalisa";
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
  els.adminSettingsPanel.hidden = !isAdmin;
}

async function apiRequest(action, payload = {}) {
  if (!cloud.apiUrl) {
    throw new Error("Укажите URL Google Apps Script.");
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
  cloud.apiUrl = normalizeText(els.apiUrlInput.value) || cloud.apiUrl || DEFAULT_API_URL;
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
    els.apiUrlInput.value = cloud.apiUrl;
  } catch (error) {
    clearCloudSession();
    updateCloudStatus(`Ошибка входа: ${error.message}`);
  }
}

function saveAdminSettings(formData) {
  if (cloud.login.toLowerCase() !== "katyalisa") return;
  const apiUrl = normalizeText(formData.get("apiUrl"));
  if (!apiUrl) {
    alert("Укажите URL Google Apps Script.");
    return;
  }
  cloud.apiUrl = apiUrl;
  localStorage.setItem(CLOUD_SETTINGS_KEY, JSON.stringify({ apiUrl: cloud.apiUrl }));
  updateCloudStatus("Ссылка Google Apps Script сохранена.");
}

async function registerTeacher(formData) {
  cloud.apiUrl = normalizeText(els.apiUrlInput.value) || cloud.apiUrl || DEFAULT_API_URL;
  const teacherName = normalizeText(formData.get("teacherName"));
  const login = normalizeText(formData.get("login"));
  const password = normalizeText(formData.get("password"));
  const adminPassword = normalizeText(formData.get("adminPassword"));

  if (!cloud.apiUrl) {
    alert("Сначала укажите URL Google Apps Script в блоке Google Sheets.");
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

  const classStudents = state.students
    .filter((student) => query || !els.classFilter.value || student.className === els.classFilter.value)
    .filter((student) => studentMatchesSearch(student, query))
    .sort((a, b) => a.name.localeCompare(b.name, "ru"));
  els.studentFilter.innerHTML = `<option value="">Все ученики</option>${classStudents
    .map((student) => `<option value="${student.id}">${escapeHtml(displayStudentName(student.name))}</option>`)
    .join("")}`;
  els.studentFilter.value = classStudents.some((student) => student.id === currentStudent) ? currentStudent : "";

  const searchMatches = state.students
    .filter((student) => studentMatchesSearch(student, query))
    .sort((a, b) => a.name.localeCompare(b.name, "ru"))
    .slice(0, 20);
  els.studentSearchList.innerHTML = searchMatches
    .map((student) => `<option value="${escapeHtml(displayStudentName(student.name))}">${escapeHtml(student.className)}</option>`)
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
  const allowedStudentIds = new Set(filteredStudents().map((student) => student.id));
  const rows = allDebts()
    .filter(({ student }) => allowedStudentIds.has(student.id))
    .sort((a, b) => a.student.className.localeCompare(b.student.className, "ru") || a.student.name.localeCompare(b.student.name, "ru"));

  els.debtorsBody.innerHTML = rows.length
    ? rows.map(({ student, poem }) => `
        <tr>
          <td data-label="ФИО">${escapeHtml(displayStudentName(student.name))}</td>
          <td data-label="Класс">${escapeHtml(student.className)}</td>
          <td data-label="Стих">${escapeHtml(poem.title)}</td>
          <td data-label="Автор">${escapeHtml(poem.author)}</td>
          <td data-label="Срок">${escapeHtml(poem.endDate || "")}</td>
        </tr>
      `).join("")
    : `<tr><td data-label="" colspan="5">Должников нет</td></tr>`;
}

function renderPoemReport() {
  const rows = [];
  for (const student of filteredStudents()) {
    const poems = state.poems
      .filter((poem) => poem.gradeLevel === classParallel(student.className))
      .sort((a, b) => a.title.localeCompare(b.title, "ru"));
    for (const poem of poems) {
      rows.push({ student, poem, grade: gradeFor(student.id, poem.id) });
    }
  }

  els.poemReportBody.innerHTML = rows.length
    ? rows.map(({ student, poem, grade }) => `
        <tr>
          <td data-label="ФИО">${escapeHtml(displayStudentName(student.name))}</td>
          <td data-label="Класс">${escapeHtml(student.className)}</td>
          <td data-label="Стих">${escapeHtml(poem.title)}</td>
          <td data-label="Оценка" class="${grade ? `grade-${grade}` : ""}">${grade || "Не сдал"}</td>
        </tr>
      `).join("")
    : `<tr><td data-label="" colspan="4">Нет данных по выбранным фильтрам</td></tr>`;
}

function renderPoems() {
  els.poemsBody.innerHTML = state.poems.length
    ? state.poems
        .slice()
        .sort((a, b) => a.gradeLevel.localeCompare(b.gradeLevel, "ru") || a.title.localeCompare(b.title, "ru"))
        .map((poem) => `
          <tr>
            <td data-label="Название">${escapeHtml(poem.title)}</td>
            <td data-label="Автор">${escapeHtml(poem.author)}</td>
            <td data-label="Параллель">${escapeHtml(poem.gradeLevel)}</td>
            <td data-label="Начало">${escapeHtml(poem.startDate)}</td>
            <td data-label="Окончание">${escapeHtml(poem.endDate)}</td>
            <td data-label=""><button class="small-action danger" data-delete-poem="${poem.id}">Удалить</button></td>
          </tr>
        `).join("")
    : `<tr><td data-label="" colspan="6">Список стихов пуст</td></tr>`;
}

function render() {
  syncFilters();
  renderStats();
  renderJournal();
  renderDebtors();
  renderPoemReport();
  renderPoems();
}

function tabPanelId(tabName) {
  return `${tabName.replace(/-([a-z])/g, (_, letter) => letter.toUpperCase())}Tab`;
}

function setActiveTab(tabName) {
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
  els.menuToggleBtn.setAttribute("aria-expanded", "true");
}

function closeMenu() {
  els.appMenu.hidden = true;
  els.menuToggleBtn.setAttribute("aria-expanded", "false");
}

function toggleMenu() {
  if (els.appMenu.hidden) {
    openMenu();
  } else {
    closeMenu();
  }
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

function exportWorkbook() {
  const poems = state.poems
    .slice()
    .sort((a, b) => a.gradeLevel.localeCompare(b.gradeLevel, "ru") || a.title.localeCompare(b.title, "ru"));
  const journalRows = [
    ["ФИО", "Класс", ...poems.map((poem) => `${poem.gradeLevel} кл. - ${poem.title}`)]
  ];

  for (const student of state.students) {
    journalRows.push([student.name, student.className, ...poems.map(() => "")]);
  }

  const debtRows = allDebts().map(({ student, poem }) => ({
    "ФИО": student.name,
    "Класс": student.className,
    "Стих": poem.title,
    "Автор": poem.author,
    "Дата окончания сдачи": poem.endDate
  }));

  const poemReportRows = [];
  for (const poem of state.poems) {
    for (const student of state.students.filter((item) => classParallel(item.className) === poem.gradeLevel)) {
      poemReportRows.push({
        "Стих": poem.title,
        "Автор": poem.author,
        "ФИО": student.name,
        "Класс": student.className,
        "Оценка": gradeFor(student.id, poem.id) || "Не сдал"
      });
    }
  }

  const workbook = XLSX.utils.book_new();
  const journalSheet = XLSX.utils.aoa_to_sheet(journalRows);
  journalSheet["!cols"] = [
    { wch: 32 },
    { wch: 12 },
    ...poems.map(() => ({ wch: 16 }))
  ];
  const headerStyle = {
    font: { bold: true },
    fill: { patternType: "solid", fgColor: { rgb: "E6F4F9" } },
    alignment: { horizontal: "center", vertical: "center", wrapText: true }
  };
  const submittedStyle = {
    fill: { patternType: "solid", fgColor: { rgb: "C6EFCE" } }
  };
  const missingStyle = {
    fill: { patternType: "solid", fgColor: { rgb: "FFC7CE" } }
  };

  for (let column = 0; column < journalRows[0].length; column += 1) {
    const address = XLSX.utils.encode_cell({ r: 0, c: column });
    journalSheet[address].s = headerStyle;
  }

  state.students.forEach((student, studentIndex) => {
    const rowIndex = studentIndex + 1;
    poems.forEach((poem, poemIndex) => {
      if (poem.gradeLevel !== classParallel(student.className)) return;
      const columnIndex = poemIndex + 2;
      const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
      if (!journalSheet[address]) journalSheet[address] = { t: "s", v: "" };
      journalSheet[address].s = gradeFor(student.id, poem.id) ? submittedStyle : missingStyle;
    });
  });

  XLSX.utils.book_append_sheet(workbook, journalSheet, "Журнал");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(debtRows), "Должники");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(poemReportRows), "По стихам");
  XLSX.writeFile(workbook, "uchet-stihov.xlsx");
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

els.adminSettingsForm.addEventListener("submit", (event) => {
  event.preventDefault();
  saveAdminSettings(new FormData(event.currentTarget));
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

document.addEventListener("click", (event) => {
  if (event.target.closest(".menu-shell")) return;
  closeMenu();
});

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") closeMenu();
});

els.classFilter.addEventListener("change", render);
els.studentFilter.addEventListener("change", render);
els.studentSearchInput.addEventListener("input", render);
els.saveDataFileBtn.addEventListener("click", saveDataFile);
els.loadDataFileBtn.addEventListener("click", loadDataFile);
els.exportWorkbookBtn.addEventListener("click", exportWorkbook);
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

document.addEventListener("click", (event) => {
  const id = event.target.dataset.deletePoem;
  if (!id) return;
  if (confirm("Удалить стих и связанные оценки?")) deletePoem(id);
});

document.querySelectorAll(".tab").forEach((button) => {
  button.addEventListener("click", () => {
    setActiveTab(button.dataset.tab);
  });
});

document.querySelectorAll(".menu-item[data-tab]").forEach((button) => {
  button.addEventListener("click", () => setActiveTab(button.dataset.tab));
});

load();
loadCloudSettings();
render();
refreshCloudData();
