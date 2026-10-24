import * as XLSX from '@e965/xlsx';
import { HyperFormula } from 'hyperformula';
import './style.css';

const medicineCount = document.querySelector('#medicine-count');
const searchInput = document.querySelector('#medicine-search');
const suggestions = document.querySelector('#suggestions');
const emptyState = document.querySelector('#empty-state');
const medicineView = document.querySelector('#medicine-view');
const animalSelect = document.querySelector('#animal-select');
const weightInput = document.querySelector('#weight-input');
const weightMessage = document.querySelector('#weight-message');
const doseResults = document.querySelector('#dose-results');
const formulaMessage = document.querySelector('#formula-message');
const resultColumns = [
  { column: 13, label: 'Dos/djur från' },
  { column: 14, label: 'Dos/djur till' },
  { column: 15, label: 'Enhet' },
  { column: 16, label: 'Dos från' },
  { column: 17, label: 'Antal från' },
  { column: 18, label: 'Antal till' },
  { column: 19, label: 'Form/enhet' },
];
const detailFields = [
  ['Styrka', 3],
  ['Enhet', 4],
  ['Licens', 5],
  ['Dos från', 7],
  ['Dos till', 8],
  ['Enhet', 9],
  ['Per kg', 10],
  ['Antal gånger/dygn', 11],
];

let workbook;
let formulaEngine;
let sheetIds = new Map();
let dataByAnimal = new Map();
let activeAnimal = 'Hund';
let selectedMedicine;
let currentMatches = [];
let lastOverriddenMedicine;

function cellAddress(row, column) {
  return { sheet: sheetIds.get(row.sheetName), row: row.sheetRow - 1, col: column };
}

function getEngineValue(row, column) {
  return formulaEngine.getCellValue(cellAddress(row, column));
}

function displayValue(value) {
  if (value === null || value === undefined || value === '') return '–';
  if (typeof value === 'object' && value.type === 'ERROR') return value.value || 'Formelfel';
  if (typeof value === 'number') {
    return new Intl.NumberFormat('sv-SE', { maximumFractionDigits: 6 }).format(value);
  }
  return String(value);
}

function cellValue(sheet, rowNumber, columnNumber) {
  const address = XLSX.utils.encode_cell({ r: rowNumber - 1, c: columnNumber });
  const cell = sheet[address];
  return cell?.v ?? '';
}

function hasHeader(row, name) {
  return String(row ?? '').trim().toLocaleLowerCase('sv-SE') === name.toLocaleLowerCase('sv-SE');
}

function findHeaderRow(sheet) {
  const range = XLSX.utils.decode_range(sheet['!ref']);
  for (let row = range.s.r; row <= range.e.r; row += 1) {
    if (
      hasHeader(cellValue(sheet, row + 1, 0), 'Benämning') &&
      hasHeader(cellValue(sheet, row + 1, 1), 'Substans') &&
      hasHeader(cellValue(sheet, row + 1, 2), 'Kategori ATC')
    ) return row + 1;
  }
  return null;
}

function makeFormulaData(sheet) {
  if (!sheet['!ref']) return [[]];
  const range = XLSX.utils.decode_range(sheet['!ref']);
  const data = [];
  for (let row = range.s.r; row <= range.e.r; row += 1) {
    const values = [];
    for (let column = range.s.c; column <= range.e.c; column += 1) {
      const cell = sheet[XLSX.utils.encode_cell({ r: row, c: column })];
      values[column] = cell?.f ? `=${cell.f}` : cell?.v ?? null;
    }
    data[row] = values;
  }
  return data;
}

function readMedicines(nextWorkbook) {
  const sheets = {};
  const parsed = new Map();
  for (const sheetName of nextWorkbook.SheetNames) {
    const sheet = nextWorkbook.Sheets[sheetName];
    const headerRow = findHeaderRow(sheet);
    if (!headerRow) continue;

    sheets[sheetName] = makeFormulaData(sheet);
    const range = XLSX.utils.decode_range(sheet['!ref']);
    const medicines = [];
    for (let row = headerRow + 1; row <= range.e.r + 1; row += 1) {
      const name = cellValue(sheet, row, 0);
      if (name === '' || name === null) continue;
      medicines.push({
        name: String(name),
        substance: String(cellValue(sheet, row, 1) ?? ''),
        atc: String(cellValue(sheet, row, 2) ?? ''),
        remark: String(cellValue(sheet, row, 6) ?? ''),
        sheetName,
        sheetRow: row,
        raw: Array.from({ length: 20 }, (_, column) => cellValue(sheet, row, column)),
        hasWeightFormula: Boolean(sheet[XLSX.utils.encode_cell({ r: row - 1, c: 12 })]?.f),
      });
    }
    parsed.set(sheetName, medicines);
  }
  if (!parsed.size) throw new Error('Hittade inga flikar med rubrikerna Benämning, Substans och Kategori ATC.');

  const engine = HyperFormula.buildFromSheets(sheets, { licenseKey: 'gpl-v3' });
  workbook = nextWorkbook;
  formulaEngine = engine;
  sheetIds = new Map(Object.keys(sheets).map((name) => [name, engine.getSheetId(name)]));
  dataByAnimal = parsed;
  const availableAnimals = [...parsed.keys()];
  activeAnimal = availableAnimals.find((name) => name.toLocaleLowerCase('sv-SE').includes('hund')) ?? availableAnimals[0];
  selectedMedicine = undefined;
  lastOverriddenMedicine = undefined;
  searchInput.value = '';
  medicineView.hidden = true;
  emptyState.hidden = false;
  hideSuggestions();
  animalSelect.replaceChildren(...availableAnimals.map((name) => new Option(animalLabel(name), name)));
  animalSelect.value = activeAnimal;
  renderAnimalState();
}

function renderAnimalState() {
  const medicines = dataByAnimal.get(activeAnimal) ?? [];
  medicineCount.textContent = `${medicines.length.toLocaleString('sv-SE')} läkemedel i listan`;
  if (selectedMedicine) renderSelectedMedicine();
}

function normalize(value) {
  return String(value ?? '').toLocaleLowerCase('sv-SE').normalize('NFD').replace(/\p{Diacritic}/gu, '');
}

function animalLabel(sheetName) {
  return sheetName.replace(/\s+läkemedel$/i, '');
}

function searchMedicines(query) {
  const search = normalize(query.trim());
  if (!search) return [];
  const terms = search.split(/\s+/).filter(Boolean);
  return (dataByAnimal.get(activeAnimal) ?? [])
    .map((medicine) => {
      const fields = [medicine.name, medicine.substance, medicine.atc].map(normalize);
      const matches = terms.every((term) => fields.some((field) => field.includes(term)));
      const score = terms.reduce((total, term) => total + (fields[0].startsWith(term) ? 3 : 1), 0);
      return { medicine, score, matches };
    })
    .filter((item) => item.matches)
    .sort((a, b) => b.score - a.score || a.medicine.name.localeCompare(b.medicine.name, 'sv'))
    .slice(0, 12)
    .map((item) => item.medicine);
}

function showSuggestions(matches) {
  currentMatches = matches;
  suggestions.replaceChildren();
  if (!matches.length || !searchInput.value.trim()) {
    hideSuggestions();
    return;
  }
  matches.forEach((medicine, index) => {
    const option = document.createElement('button');
    option.className = 'suggestion';
    option.type = 'button';
    option.role = 'option';
    option.id = `suggestion-${index}`;
    option.innerHTML = `<span class="suggestion-main">${escapeHtml(medicine.name)}</span><span class="suggestion-meta">${escapeHtml(medicine.substance || 'Substans saknas')} <span>·</span> ${escapeHtml(medicine.atc || 'ATC saknas')}</span>`;
    option.addEventListener('click', () => selectMedicine(medicine));
    suggestions.append(option);
  });
  suggestions.hidden = false;
  searchInput.setAttribute('aria-expanded', 'true');
}

function hideSuggestions() {
  suggestions.hidden = true;
  searchInput.setAttribute('aria-expanded', 'false');
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, (character) => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;',
  })[character]);
}

function restoreWeightCell(medicine) {
  if (!medicine) return;
  const sourceSheet = workbook.Sheets[medicine.sheetName];
  const sourceCell = sourceSheet[XLSX.utils.encode_cell({ r: medicine.sheetRow - 1, c: 12 })];
  const contents = sourceCell?.f ? `=${sourceCell.f}` : sourceCell?.v ?? null;
  formulaEngine.setCellContents(cellAddress(medicine, 12), [[contents]]);
}

function updateWeight() {
  if (!selectedMedicine) return;
  const rawWeight = weightInput.value;
  const weight = rawWeight === '' ? null : Number(rawWeight);
  const validWeight = weight === null || (Number.isFinite(weight) && weight > 0);
  weightMessage.hidden = validWeight;
  formulaEngine.setCellContents(cellAddress(selectedMedicine, 12), [[validWeight && weight !== null ? weight : null]]);
  lastOverriddenMedicine = selectedMedicine;
  renderResults();
}

function selectMedicine(medicine) {
  if (lastOverriddenMedicine) restoreWeightCell(lastOverriddenMedicine);
  selectedMedicine = medicine;
  const weightCell = workbook.Sheets[medicine.sheetName][XLSX.utils.encode_cell({ r: medicine.sheetRow - 1, c: 12 })];
  const sourceWeight = medicine.hasWeightFormula ? '' : (weightCell?.v ?? '');
  const savedWeight = Number(sourceWeight) > 0 ? sourceWeight : '';
  formulaEngine.setCellContents(cellAddress(medicine, 12), [[savedWeight === '' ? null : savedWeight]]);
  lastOverriddenMedicine = medicine;
  searchInput.value = medicine.name;
  hideSuggestions();
  emptyState.hidden = true;
  medicineView.hidden = false;
  weightInput.value = savedWeight === '' ? '' : savedWeight;
  weightMessage.hidden = true;
  document.querySelector('#selected-species').textContent = animalLabel(activeAnimal);
  document.querySelector('#medicine-name').textContent = medicine.name;
  document.querySelector('#medicine-subtitle').textContent = [medicine.substance, medicine.atc].filter(Boolean).join(' · ');
  renderSelectedMedicine();
}

function renderSelectedMedicine() {
  if (!selectedMedicine) return;
  const medicine = selectedMedicine;
  const details = document.querySelector('#medicine-details');
  details.replaceChildren();
  for (const [label, column] of detailFields) {
    const value = medicine.raw[column];
    if (value === '' || value === null || value === undefined) continue;
    const row = document.createElement('div');
    row.className = 'detail-row';
    row.innerHTML = `<span>${escapeHtml(label)}</span><strong>${escapeHtml(displayValue(value))}</strong>`;
    details.append(row);
  }
  const remarkCard = document.querySelector('#remark-card');
  remarkCard.hidden = !medicine.remark.trim();
  document.querySelector('#medicine-remark').textContent = medicine.remark || '';
  renderResults();
}

function renderResults() {
  if (!selectedMedicine) return;
  doseResults.replaceChildren();
  let hasFormulaError = false;
  for (const { column, label } of resultColumns) {
    const result = getEngineValue(selectedMedicine, column);
    const error = result && typeof result === 'object' && result.type === 'ERROR';
    hasFormulaError ||= error;
    const item = document.createElement('div');
    item.className = `result-item${error ? ' result-error' : ''}`;
    item.innerHTML = `<span class="result-label">${escapeHtml(label)}</span><strong>${escapeHtml(displayValue(result))}</strong>`;
    doseResults.append(item);
  }
  formulaMessage.hidden = !hasFormulaError;
  formulaMessage.textContent = hasFormulaError
    ? 'En eller flera formler i den här raden ger fel i Excel. Kontrollera den ursprungliga raden innan resultatet används.'
    : '';
}

animalSelect.addEventListener('change', () => {
  if (lastOverriddenMedicine) restoreWeightCell(lastOverriddenMedicine);
  activeAnimal = animalSelect.value;
  selectedMedicine = undefined;
  lastOverriddenMedicine = undefined;
  searchInput.value = '';
  medicineView.hidden = true;
  emptyState.hidden = false;
  hideSuggestions();
  renderAnimalState();
  searchInput.focus();
});

searchInput.addEventListener('input', () => showSuggestions(searchMedicines(searchInput.value)));
searchInput.addEventListener('focus', () => searchInput.select());
searchInput.addEventListener('keydown', (event) => {
  if (event.key === 'Escape') hideSuggestions();
  if (event.key === 'Enter' && currentMatches[0]) {
    event.preventDefault();
    selectMedicine(currentMatches[0]);
  }
  if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'k') {
    event.preventDefault();
    searchInput.focus();
  }
});
document.addEventListener('click', (event) => {
  if (!event.target.closest('.search-panel')) hideSuggestions();
});
weightInput.addEventListener('input', updateWeight);
document.querySelector('#clear-selection').addEventListener('click', () => {
  if (lastOverriddenMedicine) restoreWeightCell(lastOverriddenMedicine);
  selectedMedicine = undefined;
  lastOverriddenMedicine = undefined;
  weightInput.value = '';
  weightMessage.hidden = true;
  searchInput.value = '';
  medicineView.hidden = true;
  emptyState.hidden = false;
  searchInput.focus();
});
async function loadDefaultWorkbook() {
  try {
    const response = await fetch('/Doslista.xlsx');
    if (!response.ok) throw new Error(`Kunde inte läsa Doslista.xlsx (${response.status}).`);
    const data = await response.arrayBuffer();
    readMedicines(XLSX.read(data, { type: 'array', cellFormula: true, cellNF: true }));
  } catch (error) {
    medicineCount.textContent = 'Kunde inte läsa läkemedelslistan';
    console.error('Kunde inte läsa standardarbetsboken:', error);
  }
}

loadDefaultWorkbook();
