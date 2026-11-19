"use strict";

const el = (id) => document.getElementById(id);
const wheel = el("wheel");
const ctx = wheel.getContext("2d");
const spin = el("spin");
const removeWinner = el("remove-winner");
const fileInput = el("file");
const sheetSelect = el("sheet");
const columnSelect = el("column");
const header = el("header");
const nameInput = el("participant-name");
const exportButton = el("export");
const colors = ["#6246ba", "#92266a", "#aa4c32", "#936017", "#116f66", "#225d9e", "#4f3a9d", "#a03a59"];
const turn = 2 * Math.PI;
const pointer = -Math.PI / 2;
const radius = wheel.width / 2;
const preferredHeaders = ["nombre", "nombres", "nombre completo", "nombres completos", "participante", "participantes", "name", "full name"];

let workbook;
let columns = new Map();
let participants = [];
let winner = null;
const excludedCells = new Set();
const manualNames = [];
let firstRow = 0;
let rotation = 0;
let busy = false;

function drawWheel() {
  const count = participants.length || colors.length;
  const step = turn / count;
  ctx.clearRect(0, 0, wheel.width, wheel.height);
  ctx.save();
  ctx.translate(radius, radius);
  const fills = colors.map((color) => {
    const shade = ctx.createRadialGradient(0, 0, 0, 0, 0, radius);
    shade.addColorStop(0, "#21143f");
    shade.addColorStop(.68, color);
    shade.addColorStop(1, color);
    return shade;
  });

  for (let index = 0; index < count; index++) {
    const angle = pointer + rotation + index * step;
    ctx.beginPath();
    ctx.moveTo(0, 0);
    ctx.arc(0, 0, radius, angle, angle + step + 0.0001);
    ctx.closePath();
    ctx.fillStyle = fills[index % fills.length];
    ctx.fill();
    if (count <= 60) {
      ctx.beginPath();
      ctx.moveTo(0, 0);
      ctx.lineTo(radius * Math.cos(angle), radius * Math.sin(angle));
      ctx.strokeStyle = "#ffebba65";
      ctx.lineWidth = 2;
      ctx.stroke();
    }
  }

  const size = Math.min(41, Math.floor(step * radius * .39));
  if (participants.length && size >= 21) {
    ctx.font = `700 ${size}px Segoe UI, sans-serif`;
    ctx.fillStyle = "#ffebce";
    ctx.textBaseline = "middle";
    for (let index = 0; index < count; index++) {
      const midpoint = pointer + rotation + (index + .5) * step;
      const leftSide = Math.cos(midpoint) < 0;
      let label = participants[index].name;
      if (label.length > 50) label = label.slice(0, 49) + "…";
      if (ctx.measureText(label).width > radius * .7) {
        while (ctx.measureText(label + "…").width > radius * .7) label = label.slice(0, -1);
        label += "…";
      }
      ctx.save();
      ctx.rotate(midpoint);
      ctx.translate(radius * .89, 0);
      if (leftSide) ctx.rotate(Math.PI);
      ctx.textAlign = leftSide ? "left" : "right";
      ctx.fillText(label, 0, 0);
      ctx.restore();
    }
  }
  ctx.restore();
}

function setResult(label, value, hasWinner = false) {
  el("result-label").textContent = label;
  el("winner").textContent = value;
  el("result").classList.toggle("has-winner", hasWinner);
}

function syncControls() {
  for (const control of [fileInput, sheetSelect, columnSelect, header, nameInput, el("add")]) control.disabled = busy;
  for (const button of el("names").querySelectorAll(".remove")) button.disabled = busy;
  spin.disabled = busy || !participants.length;
  exportButton.disabled = busy || !participants.length;
  removeWinner.hidden = !winner;
  removeWinner.disabled = busy;
}

function refreshNames(removedWinner) {
  const cells = columns.get(Number(columnSelect.value)) ?? [];
  const imported = cells
    .filter(({ row }) => !header.checked || row !== firstRow)
    .map(({ row, value }) => ({
      name: value,
      cellKey: JSON.stringify([sheetSelect.value, columnSelect.value, row]),
    }))
    .filter(({ cellKey }) => !excludedCells.has(cellKey));
  participants = imported.concat(manualNames.map((name, manualIndex) => ({ name, manualIndex })));
  winner = null;
  rotation = 0;
  const list = el("names");
  const items = document.createDocumentFragment();
  for (const { name, manualIndex } of participants) {
    const item = document.createElement("li");
    const label = document.createElement("span");
    label.textContent = name;
    item.append(label);
    if (manualIndex !== undefined) {
      const remove = document.createElement("button");
      remove.type = "button";
      remove.className = "remove";
      remove.dataset.manualIndex = manualIndex;
      remove.setAttribute("aria-label", `Quitar a ${name}`);
      remove.textContent = "×";
      item.append(remove);
    }
    items.append(item);
  }
  list.replaceChildren(items);
  list.hidden = !participants.length;
  el("count").textContent = participants.length;
  el("empty").hidden = !!participants.length;
  el("export-error").hidden = true;
  if (removedWinner) {
    setResult("Ganador retirado de la lista", removedWinner.name, true);
  } else {
    setResult("Resultado", participants.length ? "Todo listo para girar" : "Agrega participantes para empezar");
  }
  drawWheel();
  syncControls();
}

function refreshColumns(preferredColumn) {
  const options = [];
  let suggested;
  for (const [index, cells] of columns) {
    const heading = header.checked && cells[0].row === firstRow ? cells[0].value : "";
    const label = XLSX.utils.encode_col(index);
    const normalized = heading.normalize("NFD").replace(/\p{Diacritic}/gu, "").toLowerCase();
    if (preferredHeaders.includes(normalized)) suggested ??= index;
    options.push(new Option(heading ? `${label} · ${heading}` : `Columna ${label}`, index));
  }
  columnSelect.replaceChildren(...options);
  const selected = preferredColumn ?? suggested;
  if (selected !== undefined) columnSelect.value = selected;
  refreshNames();
}

function loadSheet() {
  const sheet = workbook.Sheets[sheetSelect.value];
  columns = new Map();
  firstRow = Infinity;
  for (const [address, cell] of Object.entries(sheet)) {
    if (address.startsWith("!") || cell.t === "e" || cell.v == null) continue;
    const value = XLSX.utils.format_cell(cell).trim();
    if (!value) continue;
    const { r: row, c: column } = XLSX.utils.decode_cell(address);
    if (!columns.has(column)) columns.set(column, []);
    columns.get(column).push({ row, value });
    firstRow = Math.min(firstRow, row);
  }
  columns = new Map([...columns].sort(([left], [right]) => left - right));
  for (const cells of columns.values()) cells.sort((left, right) => left.row - right.row);
  refreshColumns();
}

fileInput.addEventListener("change", async () => {
  const file = fileInput.files[0];
  if (!file) return;
  const error = el("error");
  busy = true;
  syncControls();
  error.hidden = true;
  try {
    if (!/\.xlsx?$/i.test(file.name)) throw new Error("Selecciona un archivo .xlsx o .xls.");
    const bytes = new Uint8Array(await file.arrayBuffer());
    const zip = bytes[0] === 0x50 && bytes[1] === 0x4b;
    const ole = bytes[0] === 0xd0 && bytes[1] === 0xcf && bytes[2] === 0x11 && bytes[3] === 0xe0;
    const legacy = /\.xls$/i.test(file.name) && bytes[0] === 0x09 && [0, 2, 4, 8].includes(bytes[1]);
    if (!zip && !ole && !legacy) throw new Error("El archivo no tiene un formato de Excel válido.");
    const parsed = XLSX.read(bytes, { type: "array", dense: false });
    if (!parsed.SheetNames.some((name) => parsed.Sheets[name]?.["!ref"])) {
      throw new Error("El archivo no contiene hojas con datos.");
    }
    workbook = parsed;
    excludedCells.clear();
    sheetSelect.replaceChildren(...workbook.SheetNames.map((name) => new Option(name, name)));
    sheetSelect.value = workbook.SheetNames.find((name) => workbook.Sheets[name]?.["!ref"]);
    header.checked = true;
    loadSheet();
    el("settings").hidden = false;
    el("file-title").textContent = file.name;
    el("file-type").textContent = "Seleccionar otro archivo";
  } catch (problem) {
    error.textContent = problem.message.startsWith("Selecciona") || problem.message.startsWith("El archivo")
      ? problem.message : "No se pudo leer el archivo. Verifica que se abra en Excel.";
    error.hidden = false;
  } finally {
    fileInput.value = "";
    busy = false;
    syncControls();
  }
});

sheetSelect.addEventListener("change", loadSheet);
columnSelect.addEventListener("change", () => refreshNames());
header.addEventListener("change", () => {
  const selected = columnSelect.value;
  refreshColumns(selected === "" ? undefined : Number(selected));
});

nameInput.addEventListener("input", () => nameInput.setCustomValidity(""));
el("add-form").addEventListener("submit", (event) => {
  event.preventDefault();
  if (busy) return;
  const name = nameInput.value.trim();
  if (!name) {
    nameInput.setCustomValidity("Escribe un nombre.");
    nameInput.reportValidity();
    return;
  }
  manualNames.push(name);
  nameInput.value = "";
  refreshNames();
  nameInput.focus();
});

el("names").addEventListener("click", (event) => {
  const button = event.target.closest("[data-manual-index]");
  if (!button || busy) return;
  manualNames.splice(Number(button.dataset.manualIndex), 1);
  refreshNames();
  nameInput.focus();
});

removeWinner.addEventListener("click", () => {
  if (busy || !winner) return;
  if (winner.manualIndex === undefined) {
    excludedCells.add(winner.cellKey);
  } else {
    manualNames.splice(winner.manualIndex, 1);
  }
  refreshNames(winner);
  (spin.disabled ? el("result") : spin).focus({ preventScroll: true });
});

exportButton.addEventListener("click", () => {
  if (busy || !participants.length) return;
  const error = el("export-error");
  error.hidden = true;
  try {
    const output = XLSX.utils.book_new();
    const sheet = XLSX.utils.aoa_to_sheet([["Nombre"], ...participants.map(({ name }) => [name])]);
    sheet["!cols"] = [{ wch: 35 }];
    XLSX.utils.book_append_sheet(output, sheet, "Participantes");
    XLSX.writeFile(output, "Lista de participantes.xlsx");
  } catch {
    error.textContent = "No se pudo generar la lista de Excel.";
    error.hidden = false;
  }
});

function randomIndex(count) {
  const value = new Uint32Array(1);
  const limit = 2 ** 32 - (2 ** 32 % count);
  do { crypto.getRandomValues(value); } while (value[0] >= limit);
  return value[0] % count;
}

function celebrate() {
  if (matchMedia("(prefers-reduced-motion: reduce)").matches) return;
  const celebration = el("celebration");
  const pieces = document.createDocumentFragment();
  for (let index = 0; index < 26; index++) {
    const piece = document.createElement("i");
    const angle = index * turn / 26;
    const distance = 75 + Math.random() * 125;
    piece.style.setProperty("--x", `${Math.cos(angle) * distance}px`);
    piece.style.setProperty("--y", `${Math.sin(angle) * distance}px`);
    piece.style.setProperty("--burst-color", ["#ffc979", "#fa8eb7", "#a98dff", "#72dfca"][index % 4]);
    pieces.append(piece);
  }
  celebration.replaceChildren(pieces);
}

spin.addEventListener("click", () => {
  if (busy || !participants.length) return;
  winner = null;
  busy = true;
  syncControls();
  el("spin-label").textContent = "Girando…";
  setResult("Resultado", "Sorteo en curso");
  const winnerIndex = randomIndex(participants.length);
  const target = turn - (winnerIndex + .5) * turn / participants.length;
  const initial = rotation;
  const distance = 5 * turn + (target - initial + turn) % turn;
  const duration = matchMedia("(prefers-reduced-motion: reduce)").matches ? 0 : 3500;
  const started = performance.now();

  function frame(now) {
    const progress = duration ? Math.min(1, Math.max(0, (now - started) / duration)) : 1;
    rotation = (initial + distance * (1 - (1 - progress) ** 4)) % turn;
    drawWheel();
    if (progress < 1) {
      requestAnimationFrame(frame);
      return;
    }
    winner = participants[winnerIndex];
    setResult("Ganador del sorteo", winner.name, true);
    el("spin-label").textContent = "Girar ruleta";
    celebrate();
    busy = false;
    syncControls();
  }
  requestAnimationFrame(frame);
});

drawWheel();
