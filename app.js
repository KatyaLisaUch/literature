const STORAGE_KEY = "poetry-tracker-v1";

const state = {
  students: [],
  poems: [],
  grades: {}
};

const els = {
  studentsFile: document.getElementById("studentsFile"),
  poemsFile: document.getElementById("poemsFile"),
  poemForm: document.getElementById("poemForm"),
  classFilter: document.getElementById("classFilter"),
  poemFilter: document.getElementById("poemFilter"),
  journalHead: document.getElementById("journalHead"),
  journalBody: document.getElementById("journalBody"),
  debtorsBody: document.getElementById("debtorsBody"),
  poemReportBody: document.getElementById("poemReportBody"),
  poemsBody: document.getElementById("poemsBody"),
  studentsCount: document.getElementById("studentsCount"),
  poemsCount: document.getElementById("poemsCount"),
  gradesCount: document.getElementById("gradesCount"),
  debtsCount: document.getElementById("debtsCount"),
  emptyState: document.getElementById("emptyState"),
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

function selectedPoemId() {
  return els.poemFilter.value;
}

function filteredStudents() {
  const className = selectedClass();
  return state.students
    .filter((student) => !className || student.className === className)
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
  const currentPoem = selectedPoemId();
  const classes = [...new Set(state.students.map((student) => student.className).filter(Boolean))]
    .sort((a, b) => a.localeCompare(b, "ru"));

  els.classFilter.innerHTML = `<option value="">Все классы</option>${classes
    .map((className) => `<option value="${escapeHtml(className)}">${escapeHtml(className)}</option>`)
    .join("")}`;
  els.classFilter.value = classes.includes(currentClass) ? currentClass : "";

  const poems = state.poems.slice().sort((a, b) => a.gradeLevel.localeCompare(b.gradeLevel, "ru") || a.title.localeCompare(b.title, "ru"));
  els.poemFilter.innerHTML = poems.length
    ? poems.map((poem) => `<option value="${poem.id}">${escapeHtml(poem.gradeLevel)} кл. - ${escapeHtml(poem.title)}</option>`).join("")
    : `<option value="">Нет стихов</option>`;
  els.poemFilter.value = poems.some((poem) => poem.id === currentPoem) ? currentPoem : poems[0]?.id || "";
}

function escapeHtml(value) {
  return normalizeText(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function renderStats() {
  const gradeCount = Object.keys(state.grades).length;
  const debtCount = allDebts().length;
  els.studentsCount.textContent = state.students.length;
  els.poemsCount.textContent = state.poems.length;
  els.gradesCount.textContent = gradeCount;
  els.debtsCount.textContent = debtCount;
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
        <td class="student-name">${escapeHtml(student.name)}</td>
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
  const className = selectedClass();
  const rows = allDebts()
    .filter(({ student }) => !className || student.className === className)
    .sort((a, b) => a.student.className.localeCompare(b.student.className, "ru") || a.student.name.localeCompare(b.student.name, "ru"));

  els.debtorsBody.innerHTML = rows.length
    ? rows.map(({ student, poem }) => `
        <tr>
          <td>${escapeHtml(student.name)}</td>
          <td>${escapeHtml(student.className)}</td>
          <td>${escapeHtml(poem.title)}</td>
          <td>${escapeHtml(poem.author)}</td>
          <td>${escapeHtml(poem.endDate || "")}</td>
        </tr>
      `).join("")
    : `<tr><td colspan="5">Должников нет</td></tr>`;
}

function renderPoemReport() {
  const poem = state.poems.find((item) => item.id === selectedPoemId());
  if (!poem) {
    els.poemReportBody.innerHTML = `<tr><td colspan="3">Выберите стих</td></tr>`;
    return;
  }

  const rows = state.students
    .filter((student) => classParallel(student.className) === poem.gradeLevel)
    .sort((a, b) => a.className.localeCompare(b.className, "ru") || a.name.localeCompare(b.name, "ru"));

  els.poemReportBody.innerHTML = rows.length
    ? rows.map((student) => {
        const grade = gradeFor(student.id, poem.id);
        return `
          <tr>
            <td>${escapeHtml(student.name)}</td>
            <td>${escapeHtml(student.className)}</td>
            <td class="${grade ? `grade-${grade}` : ""}">${grade || "Не сдал"}</td>
          </tr>
        `;
      }).join("")
    : `<tr><td colspan="3">Нет учеников для этой параллели</td></tr>`;
}

function renderPoems() {
  els.poemsBody.innerHTML = state.poems.length
    ? state.poems
        .slice()
        .sort((a, b) => a.gradeLevel.localeCompare(b.gradeLevel, "ru") || a.title.localeCompare(b.title, "ru"))
        .map((poem) => `
          <tr>
            <td>${escapeHtml(poem.title)}</td>
            <td>${escapeHtml(poem.author)}</td>
            <td>${escapeHtml(poem.gradeLevel)}</td>
            <td>${escapeHtml(poem.startDate)}</td>
            <td>${escapeHtml(poem.endDate)}</td>
            <td><button class="small-action danger" data-delete-poem="${poem.id}">Удалить</button></td>
          </tr>
        `).join("")
    : `<tr><td colspan="6">Список стихов пуст</td></tr>`;
}

function render() {
  syncFilters();
  renderStats();
  renderJournal();
  renderDebtors();
  renderPoemReport();
  renderPoems();
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
  const journalRows = [];
  for (const student of state.students) {
    const row = {
      "ФИО": student.name,
      "Класс": student.className
    };
    const poems = state.poems.filter((poem) => poem.gradeLevel === classParallel(student.className));
    for (const poem of poems) {
      row[`${poem.title} (${poem.author})`] = gradeFor(student.id, poem.id) || "";
    }
    journalRows.push(row);
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
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(journalRows), "Журнал");
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

els.classFilter.addEventListener("change", render);
els.poemFilter.addEventListener("change", renderPoemReport);
els.exportWorkbookBtn.addEventListener("click", exportWorkbook);
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
    document.querySelectorAll(".tab").forEach((tab) => tab.classList.remove("active"));
    document.querySelectorAll(".tab-panel").forEach((panel) => panel.classList.remove("active"));
    button.classList.add("active");
    document.getElementById(`${button.dataset.tab.replace(/-([a-z])/g, (_, letter) => letter.toUpperCase())}Tab`).classList.add("active");
  });
});

load();
render();
