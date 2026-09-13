import React, { useEffect, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import * as XLSX from "xlsx";
import {
  ArrowDown, ArrowUp, BarChart3, Check, ChevronDown, ChevronLeft, ChevronRight,
  Columns3, Copy, Download, ExternalLink, EyeOff, FileCode2, FileSpreadsheet, Filter, FilterX, FolderOpen,
  GripVertical, Group, Layers3, ListFilter, PaintBucket, Palette, Pencil, Plus, RotateCcw, Search, Sheet, SlidersHorizontal, Trash2,
  Upload, X
} from "lucide-react";
import "./styles.css";

const SAMPLE = {
  id: "sample",
  workbook: "Mathematical Constants.xlsx",
  sheet: "Famous Constants",
  columns: ["Symbol", "Constant", "Value", "Category", "Known since", "Why it matters"],
  rows: [
    { Symbol: "π", Constant: "Pi", Value: 3.141592653589793, Category: "Geometry", "Known since": "Antiquity", "Why it matters": "Ratio of a circle's circumference to its diameter" },
    { Symbol: "e", Constant: "Euler's number", Value: 2.718281828459045, Category: "Analysis", "Known since": "17th century", "Why it matters": "Natural growth, logarithms, and compound interest" },
    { Symbol: "τ", Constant: "Tau", Value: 6.283185307179586, Category: "Geometry", "Known since": "Modern notation", "Why it matters": "One full turn in radians; equal to 2π" },
    { Symbol: "φ", Constant: "Golden ratio", Value: 1.618033988749895, Category: "Geometry", "Known since": "Antiquity", "Why it matters": "Self-similarity, pentagons, and Fibonacci limits" },
    { Symbol: "√2", Constant: "Pythagoras' constant", Value: 1.414213562373095, Category: "Algebra", "Known since": "Antiquity", "Why it matters": "Diagonal of a unit square and the first famous irrational" },
    { Symbol: "γ", Constant: "Euler–Mascheroni constant", Value: 0.577215664901533, Category: "Number theory", "Known since": "18th century", "Why it matters": "Difference between the harmonic series and the natural logarithm" },
    { Symbol: "ζ(3)", Constant: "Apéry's constant", Value: 1.202056903159594, Category: "Number theory", "Known since": "18th century", "Why it matters": "Value of the Riemann zeta function at 3" },
    { Symbol: "G", Constant: "Catalan's constant", Value: 0.915965594177219, Category: "Number theory", "Known since": "19th century", "Why it matters": "Appears in combinatorics, geometry, and special functions" },
    { Symbol: "ln 2", Constant: "Natural log of 2", Value: 0.693147180559945, Category: "Analysis", "Known since": "17th century", "Why it matters": "Exponential half-lives and binary information" },
    { Symbol: "√3", Constant: "Square root of 3", Value: 1.732050807568877, Category: "Algebra", "Known since": "Antiquity", "Why it matters": "Geometry of equilateral triangles and 30°–60°–90° triangles" },
    { Symbol: "δ", Constant: "Feigenbaum delta", Value: 4.66920160910299, Category: "Chaos theory", "Known since": "20th century", "Why it matters": "Universal rate of period-doubling toward chaos" },
    { Symbol: "K", Constant: "Khinchin's constant", Value: 2.685452001065306, Category: "Number theory", "Known since": "20th century", "Why it matters": "Typical geometric mean of continued-fraction coefficients" }
  ]
};

const OPERATORS = [
  ["contains", "contains"], ["notContains", "does not contain"], ["equals", "equals"],
  ["notEquals", "does not equal"], ["startsWith", "starts with"], ["endsWith", "ends with"],
  ["gt", "is greater than"], ["gte", "is at least"], ["lt", "is less than"],
  ["lte", "is at most"], ["blank", "is blank"], ["notBlank", "is not blank"]
];

const uid = () => Math.random().toString(36).slice(2, 10);
const defaultHiddenColumns = () => new Set(["Source workbook", "Source sheet"]);
const defaultColumnWidth = column => Math.min(360, Math.max(120, column.length * 8 + 36));
const text = value => value == null ? "" : String(value);
const escapeHtml = value => text(value).replace(/[&<>"']/g, character => ({
  "&": "&amp;",
  "<": "&lt;",
  ">": "&gt;",
  '"': "&quot;",
  "'": "&#39;"
}[character]));
const normalized = value => text(value).trim().toLocaleLowerCase();
const isBlank = value => value == null || text(value).trim() === "";
const valueColorIndex = (value, seed = 0) => {
  const key = normalized(value);
  let hash = 2166136261;
  for (let index = 0; index < key.length; index++) {
    hash ^= key.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }

  return (Math.abs(hash) + seed) % 8;
};
const VALUE_COLORS = ["#58a6ff", "#3fb950", "#d29922", "#a371f7", "#f85149", "#2dd4bf", "#f472b6", "#fb923c"];
const valueColorClass = (value, seed = 0) => isBlank(value) ? "" : `value-color-${valueColorIndex(value, seed)}`;
const numberValue = value => {
  if (typeof value === "number") return value;
  const parsed = Number(text(value).replace(/[,$%\s]/g, ""));
  return Number.isFinite(parsed) ? parsed : null;
};

function compareValues(a, b) {
  if (isBlank(a) && isBlank(b)) return 0;
  if (isBlank(a)) return 1;
  if (isBlank(b)) return -1;
  const an = numberValue(a);
  const bn = numberValue(b);
  if (an != null && bn != null) return an - bn;
  return text(a).localeCompare(text(b), undefined, { numeric: true, sensitivity: "base" });
}

function matches(value, operator, expected) {
  const actualText = normalized(value);
  const expectedText = normalized(expected);
  if (operator === "blank") return isBlank(value);
  if (operator === "notBlank") return !isBlank(value);
  if (operator === "contains") return actualText.includes(expectedText);
  if (operator === "notContains") return !actualText.includes(expectedText);
  if (operator === "equals") return actualText === expectedText;
  if (operator === "notEquals") return actualText !== expectedText;
  if (operator === "startsWith") return actualText.startsWith(expectedText);
  if (operator === "endsWith") return actualText.endsWith(expectedText);
  const actualNumber = numberValue(value);
  const expectedNumber = numberValue(expected);
  if (actualNumber == null || expectedNumber == null) return false;
  if (operator === "gt") return actualNumber > expectedNumber;
  if (operator === "gte") return actualNumber >= expectedNumber;
  if (operator === "lt") return actualNumber < expectedNumber;
  if (operator === "lte") return actualNumber <= expectedNumber;
  return true;
}

function compileSmartQuery(source) {
  const queryText = text(source).trim();
  if (!queryText) return { test: () => true, error: "" };
  const tokens = [];
  const regexToken = (pattern, flags = "i") => {
    try {
      return { kind: "regex", regex: new RegExp(pattern, flags.replace(/[gy]/g, "")) };
    } catch (error) {
      return { error: `Invalid regular expression: ${error.message}` };
    }
  };
  const bareToken = value => {
    const looksLikeRegex =
      value.startsWith("^") || value.endsWith("$") || value.includes(".*") ||
      /[\[\](){}+?\\|]/.test(value);
    return looksLikeRegex ? regexToken(value) : { kind: "term", value };
  };
  let index = 0;
  while (index < queryText.length) {
    if (/\s/.test(queryText[index])) {
      index++;
      continue;
    }
    if (queryText[index] === "!") {
      tokens.push({ kind: "not" });
      index++;
      continue;
    }
    if (queryText[index] === '"') {
      let value = "";
      let closed = false;
      for (index++; index < queryText.length; index++) {
        if (queryText[index] === "\\" && index + 1 < queryText.length) {
          value += queryText[++index];
          continue;
        }
        if (queryText[index] === '"') {
          index++;
          closed = true;
          break;
        }
        value += queryText[index];
      }
      if (!closed) return { test: () => false, error: "Unclosed quoted phrase" };
      tokens.push({ kind: "term", value });
      continue;
    }
    if (queryText[index] === "/") {
      let pattern = "";
      let closed = false;
      for (index++; index < queryText.length; index++) {
        if (queryText[index] === "\\" && index + 1 < queryText.length) {
          pattern += queryText[index] + queryText[++index];
          continue;
        }
        if (queryText[index] === "/") {
          index++;
          closed = true;
          break;
        }
        pattern += queryText[index];
      }
      if (!closed) return { test: () => false, error: "Unclosed regular expression" };
      let flags = "";
      while (index < queryText.length && /[a-z]/i.test(queryText[index])) flags += queryText[index++];
      const token = regexToken(pattern, flags);
      if (token.error) return { test: () => false, error: token.error };
      tokens.push(token);
      continue;
    }
    let value = "";
    while (index < queryText.length && !/\s/.test(queryText[index])) value += queryText[index++];
    const operator = value.toUpperCase();
    if (operator === "AND" || operator === "OR" || operator === "NOT") {
      tokens.push({ kind: operator.toLowerCase() });
    } else {
      const token = bareToken(value);
      if (token.error) return { test: () => false, error: token.error };
      tokens.push(token);
    }
  }
  const groups = [[]];
  let negateNext = false;
  for (const token of tokens) {
    if (token.kind === "or") {
      if (negateNext || !groups.at(-1).length) return { test: () => false, error: "OR needs a term on both sides" };
      groups.push([]);
    } else if (token.kind === "not") {
      negateNext = !negateNext;
    } else if (token.kind !== "and") {
      groups.at(-1).push({ ...token, negated: negateNext });
      negateNext = false;
    }
  }
  if (negateNext || !groups.at(-1).length) return { test: () => false, error: "Query cannot end with an operator" };
  return {
    error: "",
    test(value, fields = []) {
      const raw = text(value);
      const lower = raw.toLocaleLowerCase();
      const fieldValues = fields.map(text);
      return groups.some(group => group.every(token => {
        const matched = token.kind === "regex"
          ? token.regex.test(raw) || fieldValues.some(field => token.regex.test(field))
          : lower.includes(token.value.toLocaleLowerCase());
        return token.negated ? !matched : matched;
      }));
    }
  };
}

function dedupeHeaders(headers) {
  const counts = new Map();
  return headers.map((header, index) => {
    const base = text(header).trim() || `Column ${index + 1}`;
    const count = (counts.get(base) || 0) + 1;
    counts.set(base, count);
    return count === 1 ? base : `${base} (${count})`;
  });
}

function parseSheet(workbook, sheetName, workbookName) {
  const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
    header: 1,
    defval: null,
    raw: false,
    blankrows: false
  });
  if (!matrix.length) return null;
  let headerIndex = matrix.findIndex(row => row.some(value => !isBlank(value)));
  if (headerIndex < 0) return null;
  const columns = dedupeHeaders(matrix[headerIndex]);
  const rows = matrix.slice(headerIndex + 1)
    .filter(row => row.some(value => !isBlank(value)))
    .map(row => ({
      __rowId: uid(),
      ...Object.fromEntries(columns.map((column, index) => [column, row[index] ?? null]))
    }));
  if (!rows.length && !columns.length) return null;
  return { id: uid(), workbook: workbookName, sheet: sheetName, columns, rows };
}

async function parseFile(file) {
  const data = await file.arrayBuffer();
  const prefix = new TextDecoder().decode(data.slice(0, 128)).trimStart().toLowerCase();
  if (prefix.startsWith("<!doctype html") || prefix.startsWith("<html")) {
    throw new Error(`${file.name} is an HTML page, not an Excel workbook.`);
  }
  let workbook;
  try {
    workbook = XLSX.read(data, { type: "array", cellDates: true });
  } catch {
    throw new Error(`${file.name} could not be read as an Excel workbook.`);
  }
  const parsed = workbook.SheetNames
    .map(sheetName => parseSheet(workbook, sheetName, file.name))
    .filter(Boolean);
  if (!parsed.length) return parsed;
  const sourceUrl = URL.createObjectURL(file);
  return parsed.map(dataset => ({ ...dataset, sourceUrl }));
}

function IconButton({ label, children, className = "", ...props }) {
  return <button className={`icon-button ${className}`} title={label} aria-label={label} {...props}>{children}</button>;
}

function MyDataMark() {
  return (
    <svg className="mydata-mark" viewBox="0 0 64 64" aria-hidden="true">
      <rect x="13" y="14" width="38" height="36" rx="4" fill="var(--cp-surface)" stroke="var(--cp-accent)" strokeWidth="3" />
      <path d="M13 26h38M13 38h38M27 14v36M39 14v36" stroke="var(--cp-accent)" strokeWidth="2" />
      <rect x="28" y="27" width="10" height="10" fill="var(--cp-success)" />
    </svg>
  );
}

function App() {
  const [datasets, setDatasets] = useState([]);
  const [activeId, setActiveId] = useState("all");
  const [query, setQuery] = useState("");
  const [quickFilters, setQuickFilters] = useState({});
  const [filters, setFilters] = useState([]);
  const [sort, setSort] = useState({ column: "", direction: "asc" });
  const [hiddenColumns, setHiddenColumns] = useState(defaultHiddenColumns);
  const [groupBy, setGroupBy] = useState("");
  const [collapsedGroups, setCollapsedGroups] = useState(new Set());
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(50);
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const [sidebarWidth, setSidebarWidth] = useState(260);
  const [filterPanelOpen, setFilterPanelOpen] = useState(false);
  const [quickFiltersVisible, setQuickFiltersVisible] = useState(false);
  const [colorColumnsByView, setColorColumnsByView] = useState({});
  const [rowColorColumnByView, setRowColorColumnByView] = useState({});
  const [paletteSeedsByView, setPaletteSeedsByView] = useState({});
  const [valueColorOverridesByView, setValueColorOverridesByView] = useState({});
  const [conditionalRulesByView, setConditionalRulesByView] = useState({});
  const [colorMenuOpen, setColorMenuOpen] = useState(false);
  const [customizingColorColumn, setCustomizingColorColumn] = useState("");
  const [valueColorSearch, setValueColorSearch] = useState("");
  const [columnsOpen, setColumnsOpen] = useState(false);
  const [exportOpen, setExportOpen] = useState(false);
  const [loadedFileMenuKey, setLoadedFileMenuKey] = useState("");
  const [dragging, setDragging] = useState(false);
  const [draggedRowId, setDraggedRowId] = useState(null);
  const [dragOverRowId, setDragOverRowId] = useState(null);
  const [manualOrders, setManualOrders] = useState({});
  const [columnWidths, setColumnWidths] = useState({});
  const [columnAliases, setColumnAliases] = useState({});
  const [editingColumn, setEditingColumn] = useState("");
  const [renameDraft, setRenameDraft] = useState("");
  const [contextMenu, setContextMenu] = useState(null);
  const [notice, setNotice] = useState(null);
  const inputRef = useRef(null);
  const searchInputRef = useRef(null);
  const exportMenuRef = useRef(null);
  const colorMenuRef = useRef(null);
  const columnsMenuRef = useRef(null);

  const selected = useMemo(() => {
    let result;
    if (activeId.startsWith("file:")) {
      const fileKey = activeId.slice(5);
      const fileSheets = datasets.filter(dataset =>
        (dataset.sourceUrl || `name:${dataset.workbook}`) === fileKey
      );
      if (!fileSheets.length) return null;
      if (fileSheets.length === 1) {
        result = fileSheets[0];
      } else {
        const columns = ["Source sheet"];
        fileSheets.forEach(dataset => dataset.columns.forEach(column => {
          if (!columns.includes(column)) columns.push(column);
        }));
        const rows = fileSheets.flatMap(dataset => dataset.rows.map(row => ({
          "Source sheet": dataset.sheet,
          ...row
        })));
        result = {
          id: activeId,
          workbook: fileSheets[0].workbook,
          sheet: "All workbook sheets",
          sourceUrl: fileSheets[0].sourceUrl,
          columns,
          rows
        };
      }
    } else if (activeId !== "all") {
      result = datasets.find(dataset => dataset.id === activeId) || null;
    } else {
      if (!datasets.length) return null;
      const columns = ["Source workbook", "Source sheet"];
      datasets.forEach(dataset => dataset.columns.forEach(column => {
        if (!columns.includes(column)) columns.push(column);
      }));
      const rows = datasets.flatMap(dataset => dataset.rows.map(row => ({
        "Source workbook": dataset.workbook,
        "Source sheet": dataset.sheet,
        ...row
      })));
      result = { id: "all", workbook: "All workbooks", sheet: "Combined sheets", columns, rows };
    }
    const order = manualOrders[activeId];
    if (!result || !order?.length) return result;
    const rank = new Map(order.map((id, index) => [id, index]));
    return {
      ...result,
      rows: [...result.rows].sort((a, b) =>
        (rank.get(a.__rowId) ?? Number.MAX_SAFE_INTEGER) -
        (rank.get(b.__rowId) ?? Number.MAX_SAFE_INTEGER)
      )
    };
  }, [datasets, activeId, manualOrders]);

  const columns = selected?.columns || [];
  const visibleColumns = columns.filter(column => !hiddenColumns.has(column));
  const displayColumn = column => columnAliases[activeId]?.[column] || column;
  const displayColumns = visibleColumns.map(displayColumn);
  const coloredColumns = new Set(colorColumnsByView[activeId] || []);
  const rowColorColumn = rowColorColumnByView[activeId] || "";
  const paletteSeeds = paletteSeedsByView[activeId] || {};
  const valueColorOverrides = valueColorOverridesByView[activeId] || {};
  const conditionalRules = conditionalRulesByView[activeId] || [];
  const formatCount = coloredColumns.size + conditionalRules.length;
  const loadedFiles = useMemo(() => {
    const files = new Map();
    datasets.forEach(dataset => {
      const key = dataset.sourceUrl || `name:${dataset.workbook}`;
      const current = files.get(key) || { key, name: dataset.workbook, url: dataset.sourceUrl, sheets: 0 };
      current.sheets++;
      files.set(key, current);
    });
    return [...files.values()];
  }, [datasets]);
  const tableWidth = 64 + visibleColumns.reduce((sum, column) => sum + (columnWidths[column] || defaultColumnWidth(column)), 0);
  const smartQuery = useMemo(() => compileSmartQuery(query), [query]);
  const customizableValues = useMemo(() => {
    if (!selected || !customizingColorColumn) return [];
    const values = new Map();
    selected.rows.forEach(row => {
      const value = row[customizingColorColumn];
      if (isBlank(value)) return;
      const key = normalized(value);
      const current = values.get(key) || { key, value: text(value), count: 0 };
      current.count++;
      values.set(key, current);
    });
    const search = normalized(valueColorSearch);
    return [...values.values()]
      .filter(entry => !search || normalized(entry.value).includes(search))
      .sort((a, b) => b.count - a.count || a.value.localeCompare(b.value, undefined, { numeric: true, sensitivity: "base" }));
  }, [selected, customizingColorColumn, valueColorSearch]);

  useEffect(() => {
    setQuickFilters({});
    setFilters([]);
    setSort({ column: "", direction: "asc" });
    setHiddenColumns(defaultHiddenColumns());
    setGroupBy("");
    setEditingColumn("");
    setPage(1);
  }, [activeId]);

  useEffect(() => {
    const focusSearch = event => {
      const target = event.target;
      const isTyping = target instanceof HTMLElement &&
        (target.matches("input, textarea, select") || target.isContentEditable);
      const commandShortcut = event.ctrlKey && event.shiftKey && event.key.toLowerCase() === "f";
      const slashShortcut = event.key === "/" && !event.ctrlKey && !event.metaKey && !event.altKey && !isTyping;
      if (!commandShortcut && !slashShortcut) return;
      event.preventDefault();
      searchInputRef.current?.focus();
      searchInputRef.current?.select();
    };
    window.addEventListener("keydown", focusSearch);
    return () => window.removeEventListener("keydown", focusSearch);
  }, []);

  useEffect(() => {
    if (!contextMenu) return;
    const close = event => {
      if (event.type === "keydown" && event.key !== "Escape") return;
      if (event.target instanceof Element && event.target.closest(".context-menu")) return;
      setContextMenu(null);
    };
    window.addEventListener("pointerdown", close);
    window.addEventListener("keydown", close);
    window.addEventListener("resize", close);
    window.addEventListener("scroll", close, true);
    return () => {
      window.removeEventListener("pointerdown", close);
      window.removeEventListener("keydown", close);
      window.removeEventListener("resize", close);
      window.removeEventListener("scroll", close, true);
    };
  }, [contextMenu]);

  useEffect(() => {
    if (!notice) return;
    const timer = window.setTimeout(() => setNotice(null), notice.type === "error" ? 5000 : 3000);
    return () => window.clearTimeout(timer);
  }, [notice]);

  useEffect(() => {
    if (!exportOpen && !colorMenuOpen && !columnsOpen) return;
    const closeMenus = event => {
      if (event.type === "keydown") {
        if (event.key !== "Escape") return;
        setExportOpen(false);
        setColorMenuOpen(false);
        setColumnsOpen(false);
        return;
      }
      const target = event.target;
      if (target instanceof Node && (
        exportMenuRef.current?.contains(target) ||
        colorMenuRef.current?.contains(target) ||
        columnsMenuRef.current?.contains(target)
      )) return;
      setExportOpen(false);
      setColorMenuOpen(false);
      setColumnsOpen(false);
    };
    document.addEventListener("pointerdown", closeMenus);
    document.addEventListener("keydown", closeMenus);
    return () => {
      document.removeEventListener("pointerdown", closeMenus);
      document.removeEventListener("keydown", closeMenus);
    };
  }, [exportOpen, colorMenuOpen, columnsOpen]);

  useEffect(() => {
    if (!loadedFileMenuKey) return;
    const close = event => {
      if (event.type === "keydown" && event.key !== "Escape") return;
      if (event.type !== "keydown" && event.target instanceof Element && event.target.closest(".loaded-file-wrap")) return;
      setLoadedFileMenuKey("");
    };
    document.addEventListener("pointerdown", close);
    document.addEventListener("keydown", close);
    return () => {
      document.removeEventListener("pointerdown", close);
      document.removeEventListener("keydown", close);
    };
  }, [loadedFileMenuKey]);

  const filteredRows = useMemo(() => {
    if (!selected) return [];
    let rows = selected.rows.filter(row => {
      const values = columns.map(column => row[column]);
      if (query && !smartQuery.test(values.join(" "), values)) return false;
      if (!Object.entries(quickFilters).every(([column, value]) => !value || normalized(row[column]).includes(normalized(value)))) return false;
      return filters.every(filter => matches(row[filter.column], filter.operator, filter.value));
    });
    if (sort.column) {
      rows = [...rows].sort((a, b) => {
        const result = compareValues(a[sort.column], b[sort.column]);
        return sort.direction === "asc" ? result : -result;
      });
    }
    return rows;
  }, [selected, columns, query, smartQuery, quickFilters, filters, sort]);

  const groups = useMemo(() => {
    if (!groupBy) return null;
    const map = new Map();
    filteredRows.forEach(row => {
      const key = isBlank(row[groupBy]) ? "(Blank)" : text(row[groupBy]);
      if (!map.has(key)) map.set(key, []);
      map.get(key).push(row);
    });
    return [...map.entries()].sort(([a], [b]) => a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" }));
  }, [filteredRows, groupBy]);

  const totalPages = pageSize === 0 ? 1 : Math.max(1, Math.ceil(filteredRows.length / pageSize));
  useEffect(() => setPage(current => Math.min(current, totalPages)), [totalPages]);
  const start = pageSize === 0 ? 0 : (page - 1) * pageSize;
  const pageRows = pageSize === 0 ? filteredRows : filteredRows.slice(start, start + pageSize);
  const filterCount = filters.length + Object.values(quickFilters).filter(Boolean).length + (query ? 1 : 0);

  async function importFiles(fileList) {
    const files = [...fileList].filter(file => /\.(xlsx|xlsm|xls|csv)$/i.test(file.name));
    if (!files.length) {
      setNotice({ type: "error", text: "Choose one or more Excel or CSV files." });
      return;
    }
    try {
      const results = await Promise.all(files.map(parseFile));
      const incoming = results.flat();
      if (!incoming.length) throw new Error("No non-empty worksheets were found.");
      setDatasets(current => [...current.filter(dataset => dataset.id !== "sample"), ...incoming]);
      setActiveId(incoming.length === 1 ? incoming[0].id : "all");
      setNotice({ type: "success", text: `Imported ${files.length} workbook${files.length === 1 ? "" : "s"} with ${incoming.length} worksheet${incoming.length === 1 ? "" : "s"}.` });
    } catch (error) {
      setNotice({ type: "error", text: error.message });
    }
  }

  function loadSample() {
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(SAMPLE.rows), SAMPLE.sheet);
    const sourceUrl = URL.createObjectURL(new Blob([
      XLSX.write(workbook, { type: "array", bookType: "xlsx" })
    ], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    setDatasets([{ ...SAMPLE, sourceUrl, rows: SAMPLE.rows.map(row => ({ __rowId: uid(), ...row })) }]);
    setActiveId(SAMPLE.id);
    setNotice({ type: "success", text: "Loaded sample data. Import an Excel file whenever you're ready." });
  }

  function clearData() {
    [...new Set(datasets.map(dataset => dataset.sourceUrl).filter(Boolean))]
      .forEach(url => URL.revokeObjectURL(url));
    setDatasets([]);
    setActiveId("all");
    clearFilters();
  }

  function clearFilters() {
    setQuery("");
    setQuickFilters({});
    setFilters([]);
    setGroupBy("");
    setSort({ column: "", direction: "asc" });
    setPage(1);
  }

  function toggleSort(column) {
    setSort(current => current.column === column
      ? { column, direction: current.direction === "asc" ? "desc" : "asc" }
      : { column, direction: "asc" });
  }

  function toggleColumn(column) {
    setHiddenColumns(current => {
      const next = new Set(current);
      next.has(column) ? next.delete(column) : next.add(column);
      return next;
    });
  }

  function toggleColorColumn(column) {
    setColorColumnsByView(current => {
      const next = new Set(current[activeId] || []);
      next.has(column) ? next.delete(column) : next.add(column);
      return { ...current, [activeId]: [...next] };
    });
    if (coloredColumns.has(column) && customizingColorColumn === column) {
      setCustomizingColorColumn("");
      setValueColorSearch("");
    }
    if (coloredColumns.has(column) && rowColorColumn === column) {
      setRowColorColumnByView(current => ({ ...current, [activeId]: "" }));
    }
  }

  function setAllColorColumns(enabled) {
    setColorColumnsByView(current => ({
      ...current,
      [activeId]: enabled ? [...columns] : []
    }));
    if (!enabled) {
      setRowColorColumnByView(current => ({ ...current, [activeId]: "" }));
      setCustomizingColorColumn("");
      setValueColorSearch("");
    }
  }

  function toggleColumnColorScope(column) {
    setRowColorColumnByView(current => ({
      ...current,
      [activeId]: current[activeId] === column ? "" : column
    }));
  }

  function clearAllFormatting() {
    setColorColumnsByView(current => ({ ...current, [activeId]: [] }));
    setRowColorColumnByView(current => ({ ...current, [activeId]: "" }));
    setPaletteSeedsByView(current => ({ ...current, [activeId]: {} }));
    setValueColorOverridesByView(current => ({ ...current, [activeId]: {} }));
    setConditionalRulesByView(current => ({ ...current, [activeId]: [] }));
    setCustomizingColorColumn("");
    setValueColorSearch("");
  }

  function setValueColor(column, valueKey, color) {
    setValueColorOverridesByView(current => ({
      ...current,
      [activeId]: {
        ...(current[activeId] || {}),
        [column]: { ...(current[activeId]?.[column] || {}), [valueKey]: color }
      }
    }));
  }

  function resetValueColor(column, valueKey) {
    setValueColorOverridesByView(current => {
      const view = { ...(current[activeId] || {}) };
      const columnColors = { ...(view[column] || {}) };
      delete columnColors[valueKey];
      view[column] = columnColors;
      return { ...current, [activeId]: view };
    });
  }

  function openValueColorEditor(column) {
    setColorColumnsByView(current => {
      const next = new Set(current[activeId] || []);
      next.add(column);
      return { ...current, [activeId]: [...next] };
    });
    setCustomizingColorColumn(column);
    setValueColorSearch("");
    setColorMenuOpen(true);
  }

  function colorColumnFromContext() {
    const column = contextMenu.column;
    const wasColored = coloredColumns.has(column);
    setColorColumnsByView(current => {
      const next = new Set(current[activeId] || []);
      next.add(column);
      return { ...current, [activeId]: [...next] };
    });
    setPaletteSeedsByView(current => {
      const currentSeed = current[activeId]?.[column] || 0;
      const nextSeed = wasColored
        ? (currentSeed + 1 + Math.floor(Math.random() * (VALUE_COLORS.length - 1))) % VALUE_COLORS.length
        : currentSeed;
      return {
        ...current,
        [activeId]: { ...(current[activeId] || {}), [column]: nextSeed }
      };
    });
    setColorMenuOpen(false);
    setContextMenu(null);
  }

  function clearColumnColors(column) {
    setColorColumnsByView(current => ({
      ...current,
      [activeId]: (current[activeId] || []).filter(candidate => candidate !== column)
    }));
    setValueColorOverridesByView(current => {
      const view = { ...(current[activeId] || {}) };
      delete view[column];
      return { ...current, [activeId]: view };
    });
    setConditionalRulesByView(current => ({
      ...current,
      [activeId]: (current[activeId] || []).filter(rule => rule.column !== column)
    }));
    setPaletteSeedsByView(current => {
      const view = { ...(current[activeId] || {}) };
      delete view[column];
      return { ...current, [activeId]: view };
    });
    setRowColorColumnByView(current => (
      current[activeId] === column ? { ...current, [activeId]: "" } : current
    ));
    if (customizingColorColumn === column) {
      setCustomizingColorColumn("");
      setValueColorSearch("");
    }
    setContextMenu(null);
  }

  function addConditionalRule() {
    if (!columns.length) return;
    const rule = { id: uid(), column: columns[0], operator: "equals", value: "", color: "#58a6ff", applyToRow: false };
    setConditionalRulesByView(current => ({
      ...current,
      [activeId]: [...(current[activeId] || []), rule]
    }));
  }

  function duplicateConditionalRule(rule) {
    setConditionalRulesByView(current => {
      const rules = [...(current[activeId] || [])];
      const index = rules.findIndex(candidate => candidate.id === rule.id);
      rules.splice(index + 1, 0, { ...rule, id: uid() });
      return { ...current, [activeId]: rules };
    });
  }

  function showContextMenu(event, type, column, value = "") {
    event.preventDefault();
    const width = 250;
    const height = type === "cell" ? 310 : 250;
    setContextMenu({
      type,
      column,
      value,
      x: Math.max(8, Math.min(event.clientX, window.innerWidth - width - 8)),
      y: Math.max(8, Math.min(event.clientY, window.innerHeight - height - 8))
    });
  }

  function addContextFilter(operator) {
    const { column, value } = contextMenu;
    setFilters(current => [...current, { id: uid(), column, operator, value: text(value) }]);
    setFilterPanelOpen(true);
    setPage(1);
    setContextMenu(null);
  }

  function clearColumnFilters(column) {
    setQuickFilters(current => ({ ...current, [column]: "" }));
    setFilters(current => current.filter(filter => filter.column !== column));
    setPage(1);
    setContextMenu(null);
  }

  function colorContextValue() {
    const { column, value } = contextMenu;
    const rule = {
      id: uid(),
      column,
      operator: isBlank(value) ? "blank" : "equals",
      value: isBlank(value) ? "" : text(value),
      color: "#58a6ff"
    };
    setConditionalRulesByView(current => ({
      ...current,
      [activeId]: [...(current[activeId] || []), rule]
    }));
    setColorMenuOpen(true);
    setContextMenu(null);
  }

  async function copyContextValue() {
    const value = text(contextMenu.value);
    try {
      await navigator.clipboard.writeText(value);
    } catch {
      const textarea = document.createElement("textarea");
      textarea.value = value;
      textarea.style.cssText = "position:fixed;opacity:0";
      document.body.appendChild(textarea);
      textarea.select();
      document.execCommand("copy");
      textarea.remove();
    }
    setNotice({ type: "success", text: "Cell value copied." });
    setContextMenu(null);
  }

  function updateConditionalRule(id, patch) {
    setConditionalRulesByView(current => ({
      ...current,
      [activeId]: (current[activeId] || []).map(rule => rule.id === id ? { ...rule, ...patch } : rule)
    }));
  }

  function removeConditionalRule(id) {
    setConditionalRulesByView(current => ({
      ...current,
      [activeId]: (current[activeId] || []).filter(rule => rule.id !== id)
    }));
  }

  function conditionalRuleFor(column, value, applyToRow = false) {
    return conditionalRules.find(rule =>
      rule.column === column && Boolean(rule.applyToRow) === applyToRow &&
      matches(value, rule.operator, rule.value)
    );
  }

  function cellFormat(column, value) {
    const rule = conditionalRuleFor(column, value, false);
    if (rule) return { className: "conditional-cell", style: { "--rule-color": rule.color }, priority: 2 };
    const customColor = valueColorOverrides[column]?.[normalized(value)];
    if (coloredColumns.has(column) && customColor) {
      return { className: "custom-value-cell", style: { "--value-color": customColor }, priority: 1 };
    }
    return {
      className: coloredColumns.has(column) ? valueColorClass(value, paletteSeeds[column] || 0) : "",
      style: undefined,
      priority: coloredColumns.has(column) && !isBlank(value) ? 1 : 0
    };
  }

  function rowFormat(row) {
    const rule = conditionalRules.find(candidate =>
      candidate.applyToRow && matches(row[candidate.column], candidate.operator, candidate.value)
    );
    if (rule) return { className: "conditional-cell", style: { "--rule-color": rule.color }, priority: 2 };
    if (!rowColorColumn || !coloredColumns.has(rowColorColumn)) return { className: "", style: undefined, priority: 0 };
    const value = row[rowColorColumn];
    if (isBlank(value)) return { className: "", style: undefined, priority: 0 };
    const customColor = valueColorOverrides[rowColorColumn]?.[normalized(value)];
    return customColor
      ? { className: "custom-value-cell", style: { "--value-color": customColor }, priority: 1 }
      : { className: valueColorClass(value, paletteSeeds[rowColorColumn] || 0), style: undefined, priority: 1 };
  }

  function beginColumnRename(column) {
    setEditingColumn(column);
    setRenameDraft(displayColumn(column));
  }

  function commitColumnRename(column) {
    const nextName = renameDraft.trim();
    if (!nextName) {
      setNotice({ type: "error", text: "Column names cannot be blank." });
      return;
    }
    const duplicate = columns.some(other =>
      other !== column && displayColumn(other).toLocaleLowerCase() === nextName.toLocaleLowerCase()
    );
    if (duplicate) {
      setNotice({ type: "error", text: `A column named "${nextName}" already exists in this view.` });
      return;
    }
    setColumnAliases(current => ({
      ...current,
      [activeId]: { ...(current[activeId] || {}), [column]: nextName }
    }));
    setEditingColumn("");
    setNotice({ type: "success", text: `Renamed "${column}" to "${nextName}".` });
  }

  function contentWidthFor(column) {
    const canvas = document.createElement("canvas");
    const context = canvas.getContext("2d");
    context.font = '14px "Segoe UI", Aptos, Calibri, sans-serif';
    const values = [displayColumn(column), ...selected.rows.map(row => text(row[column]))];
    const measured = Math.max(...values.map(value => context.measureText(value).width));
    return Math.min(520, Math.max(72, Math.ceil(measured + 30)));
  }

  function fitColumnToContent(column) {
    setColumnWidths(current => ({ ...current, [column]: contentWidthFor(column) }));
  }

  function fitVisibleColumnsToContent() {
    const widths = Object.fromEntries(visibleColumns.map(column => [column, contentWidthFor(column)]));
    setColumnWidths(current => ({ ...current, ...widths }));
  }

  function startColumnResize(event, column) {
    event.preventDefault();
    event.stopPropagation();
    const startX = event.clientX;
    const header = event.currentTarget.parentElement;
    const startWidth = header.getBoundingClientRect().width;
    const move = moveEvent => {
      const width = Math.max(72, Math.round(startWidth + moveEvent.clientX - startX));
      setColumnWidths(current => ({ ...current, [column]: width }));
    };
    const stop = () => {
      document.body.classList.remove("resizing-column");
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", stop);
      window.removeEventListener("pointercancel", stop);
    };
    document.body.classList.add("resizing-column");
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", stop);
    window.addEventListener("pointercancel", stop);
  }

  function startSidebarResize(event) {
    event.preventDefault();
    if (event.detail > 1) {
      setSidebarWidth(260);
      return;
    }
    const startX = event.clientX;
    const startWidth = sidebarWidth;
    const move = moveEvent => {
      const maxWidth = Math.min(520, Math.round(window.innerWidth * 0.45));
      setSidebarWidth(Math.max(180, Math.min(maxWidth, startWidth + moveEvent.clientX - startX)));
    };
    const stop = () => {
      document.body.classList.remove("resizing-sidebar");
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", stop);
      window.removeEventListener("pointercancel", stop);
    };
    document.body.classList.add("resizing-sidebar");
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", stop);
    window.addEventListener("pointercancel", stop);
  }

  function updateFilter(id, patch) {
    setFilters(current => current.map(filter => filter.id === id ? { ...filter, ...patch } : filter));
  }

  function addFilter() {
    if (!columns.length) return;
    setFilters(current => [...current, { id: uid(), column: columns[0], operator: "contains", value: "" }]);
    setFilterPanelOpen(true);
  }

  function reorderRows(sourceId, targetId) {
    if (!selected || !sourceId || !targetId || sourceId === targetId || groupBy) return;
    const visibleIds = filteredRows.map(row => row.__rowId);
    const sourceIndex = visibleIds.indexOf(sourceId);
    const targetIndex = visibleIds.indexOf(targetId);
    if (sourceIndex < 0 || targetIndex < 0) return;
    visibleIds.splice(sourceIndex, 1);
    visibleIds.splice(targetIndex, 0, sourceId);
    const visibleSet = new Set(visibleIds);
    let nextVisibleIndex = 0;
    const mergedOrder = selected.rows.map(row =>
      visibleSet.has(row.__rowId) ? visibleIds[nextVisibleIndex++] : row.__rowId
    );
    setManualOrders(current => ({ ...current, [activeId]: mergedOrder }));
    setSort({ column: "", direction: "asc" });
    setDraggedRowId(null);
    setDragOverRowId(null);
  }

  function filteredExportRows() {
    return filteredRows.map(row =>
      Object.fromEntries(visibleColumns.map(column => [displayColumn(column), row[column] ?? ""]))
    );
  }

  function exportFilteredExcel() {
    if (!selected) return;
    const exportRows = filteredExportRows();
    const workbook = XLSX.utils.book_new();
    const sheet = XLSX.utils.json_to_sheet(exportRows, { header: displayColumns });
    if (sheet["!ref"]) sheet["!autofilter"] = { ref: sheet["!ref"] };
    sheet["!cols"] = displayColumns.map(column => ({ wch: Math.min(60, Math.max(12, column.length + 2)) }));
    XLSX.utils.book_append_sheet(workbook, sheet, "Filtered data");
    XLSX.writeFile(workbook, "mydata-export.xlsx");
    setExportOpen(false);
    setNotice({ type: "success", text: `Exported ${filteredRows.length.toLocaleString()} filtered rows.` });
  }

  function exportFilteredCsv() {
    if (!selected) return;
    const sheet = XLSX.utils.json_to_sheet(filteredExportRows(), { header: displayColumns });
    const blob = new Blob([XLSX.utils.sheet_to_csv(sheet)], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "mydata-export.csv";
    link.click();
    URL.revokeObjectURL(url);
    setExportOpen(false);
    setNotice({ type: "success", text: `Exported ${filteredRows.length.toLocaleString()} filtered rows as CSV.` });
  }

  function exportFilteredHtml() {
    if (!selected) return;
    const title = `${selected.sheet} - MyData export`;
    const exportedAt = new Intl.DateTimeFormat(undefined, {
      dateStyle: "medium",
      timeStyle: "short"
    }).format(new Date());
    const filterLabel = query ? `Smart filter: ${query}` : "No smart filter";
    const headers = visibleColumns.map(column => `<th>${escapeHtml(displayColumn(column))}</th>`).join("");
    const rows = filteredRows.map(row => {
      const rowStyle = rowFormat(row);
      return `<tr>${visibleColumns.map(column => {
      const rule = conditionalRuleFor(column, row[column], false);
      const customColor = valueColorOverrides[column]?.[normalized(row[column])];
      const cellClass = rule ? "conditional-cell" : customColor && coloredColumns.has(column) ? "custom-value-cell" : coloredColumns.has(column) ? valueColorClass(row[column], paletteSeeds[column] || 0) : "";
      const cellStyle = rule ? ` style="--rule-color:${rule.color}"` : customColor && coloredColumns.has(column) ? ` style="--value-color:${customColor}"` : "";
      const className = rule ? cellClass : rowStyle.className || cellClass;
      const style = rule ? cellStyle : rowStyle.style
        ? ` style="${rowStyle.className === "conditional-cell" ? `--rule-color:${rowStyle.style["--rule-color"]}` : `--value-color:${rowStyle.style["--value-color"]}`}"`
        : cellStyle;
      return `<td class="${className}"${style}>${isBlank(row[column]) ? '<span class="blank">—</span>' : escapeHtml(row[column])}</td>`;
      }).join("")}</tr>`;
    }).join("");
    const html = `<!doctype html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>${escapeHtml(title)}</title>
<link rel="icon" type="image/svg+xml" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'%3E%3Crect width='64' height='64' rx='12' fill='%23132d4d'/%3E%3Crect x='13' y='14' width='38' height='36' rx='4' fill='%23142c49' stroke='%2358a6ff' stroke-width='3'/%3E%3Cpath d='M13 26h38M13 38h38M27 14v36M39 14v36' stroke='%2358a6ff' stroke-width='2'/%3E%3Crect x='28' y='27' width='10' height='10' fill='%233fb950'/%3E%3C/svg%3E">
<style>
:root{color-scheme:dark;--cp-bg:#1b2939;--cp-bg-elevated:#17273a;--cp-surface:#142c49;--cp-border:#2c4057;--cp-text:#f2f6fb;--cp-text-muted:#a9bad0;--cp-text-soft:#74869d;--cp-accent:#58a6ff;--cp-success:#3fb950;--cp-value-0:rgba(88,166,255,.18);--cp-value-1:rgba(63,185,80,.18);--cp-value-2:rgba(210,153,34,.2);--cp-value-3:rgba(163,113,247,.2);--cp-value-4:rgba(248,81,73,.17);--cp-value-5:rgba(45,212,191,.18);--cp-value-6:rgba(244,114,182,.18);--cp-value-7:rgba(251,146,60,.18)}
*{box-sizing:border-box}
body{margin:0;background:var(--cp-bg);color:var(--cp-text);font:14px "Segoe UI",Aptos,Calibri,-apple-system,BlinkMacSystemFont,sans-serif}
header{position:sticky;top:0;z-index:3;display:flex;align-items:center;gap:14px;padding:13px 18px;background:var(--cp-bg-elevated);border-bottom:1px solid var(--cp-border)}
.mark{width:34px;height:34px;display:grid;place-items:center;border:1px solid var(--cp-border);border-radius:6px;background:var(--cp-surface)}
.mark svg{width:25px;height:25px}h1{margin:0;font-size:17px}.sub{margin-top:2px;color:var(--cp-text-muted);font-size:11px}
.meta{display:flex;gap:18px;flex-wrap:wrap;padding:10px 18px;color:var(--cp-text-muted);font-size:11px;border-bottom:1px solid var(--cp-border)}
.meta strong{color:var(--cp-text)}main{padding:14px 18px 28px}.table-wrap{overflow:auto;border:1px solid var(--cp-border);border-radius:8px}
table{width:100%;border-collapse:separate;border-spacing:0;background:var(--cp-surface)}
th,td{height:38px;padding:7px 10px;border-right:1px solid var(--cp-border);border-bottom:1px solid var(--cp-border);text-align:left;white-space:nowrap}
th{position:sticky;top:0;background:var(--cp-bg-elevated);font-size:11px}th:last-child,td:last-child{border-right:0}tr:last-child td{border-bottom:0}
tbody tr:hover td{background:color-mix(in srgb,var(--cp-accent) 10%,var(--cp-surface))}.blank{color:var(--cp-text-soft)}
.value-color-0{background:var(--cp-value-0)}.value-color-1{background:var(--cp-value-1)}.value-color-2{background:var(--cp-value-2)}.value-color-3{background:var(--cp-value-3)}.value-color-4{background:var(--cp-value-4)}.value-color-5{background:var(--cp-value-5)}.value-color-6{background:var(--cp-value-6)}.value-color-7{background:var(--cp-value-7)}
.conditional-cell{background:color-mix(in srgb,var(--rule-color) 23%,var(--cp-surface));box-shadow:inset 4px 0 var(--rule-color)}
.custom-value-cell{background:color-mix(in srgb,var(--value-color) 23%,var(--cp-surface));box-shadow:inset 4px 0 var(--value-color)}
</style>
</head>
<body>
<header>
  <div class="mark"><svg viewBox="0 0 64 64" aria-hidden="true"><rect x="13" y="14" width="38" height="36" rx="4" fill="var(--cp-surface)" stroke="var(--cp-accent)" stroke-width="3"/><path d="M13 26h38M13 38h38M27 14v36M39 14v36" stroke="var(--cp-accent)" stroke-width="2"/><rect x="28" y="27" width="10" height="10" fill="var(--cp-success)"/></svg></div>
  <div><h1>${escapeHtml(selected.sheet)}</h1><div class="sub">${escapeHtml(selected.workbook)}</div></div>
</header>
<div class="meta"><span><strong>Rows:</strong> ${filteredRows.length.toLocaleString()}</span><span><strong>Columns:</strong> ${visibleColumns.length}</span><span><strong>View:</strong> ${escapeHtml(filterLabel)}</span><span><strong>Exported:</strong> ${escapeHtml(exportedAt)}</span></div>
<main><div class="table-wrap"><table><thead><tr>${headers}</tr></thead><tbody>${rows || `<tr><td colspan="${visibleColumns.length}">No matching rows.</td></tr>`}</tbody></table></div></main>
</body>
</html>`;
    const blob = new Blob([html], { type: "text/html;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "mydata-filtered-view.html";
    link.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    setExportOpen(false);
    setNotice({ type: "success", text: `Exported ${filteredRows.length.toLocaleString()} filtered rows as HTML.` });
  }

  function exportAllSheets() {
    if (!datasets.length) return;
    const workbook = XLSX.utils.book_new();
    const usedNames = new Set();
    datasets.forEach((dataset, index) => {
      const order = manualOrders[dataset.id];
      const rank = new Map((order || []).map((id, position) => [id, position]));
      const rows = order?.length
        ? [...dataset.rows].sort((a, b) => (rank.get(a.__rowId) ?? Number.MAX_SAFE_INTEGER) - (rank.get(b.__rowId) ?? Number.MAX_SAFE_INTEGER))
        : dataset.rows;
      const labels = dataset.columns.map(column => columnAliases[dataset.id]?.[column] || column);
      const exportRows = rows.map(row => Object.fromEntries(dataset.columns.map((column, columnIndex) => [labels[columnIndex], row[column] ?? ""])));
      const sheet = XLSX.utils.json_to_sheet(exportRows, { header: labels });
      if (sheet["!ref"]) sheet["!autofilter"] = { ref: sheet["!ref"] };
      sheet["!cols"] = labels.map(column => ({ wch: Math.min(60, Math.max(12, column.length + 2)) }));
      const base = (dataset.sheet || `Sheet ${index + 1}`).replace(/[\\/?*[\]:]/g, " ").slice(0, 31) || `Sheet ${index + 1}`;
      let name = base;
      let suffix = 2;
      while (usedNames.has(name.toLocaleLowerCase())) {
        name = `${base.slice(0, 27)} ${suffix++}`.slice(0, 31);
      }
      usedNames.add(name.toLocaleLowerCase());
      XLSX.utils.book_append_sheet(workbook, sheet, name);
    });
    XLSX.writeFile(workbook, "mydata-all-sheets.xlsx");
    setExportOpen(false);
    setNotice({ type: "success", text: `Exported ${datasets.length} worksheet${datasets.length === 1 ? "" : "s"} as Excel.` });
  }

  function renderRows(rows) {
    return rows.map((row, rowIndex) => {
      const rowStyle = rowFormat(row);
      return (
      <tr
        key={row.__rowId}
        className={`${draggedRowId === row.__rowId ? "dragging-row" : ""} ${dragOverRowId === row.__rowId ? "drag-over-row" : ""}`}
        onDragOver={event => {
          if (!draggedRowId || draggedRowId === row.__rowId || groupBy) return;
          event.preventDefault();
          setDragOverRowId(row.__rowId);
        }}
        onDrop={event => {
          event.preventDefault();
          reorderRows(draggedRowId || event.dataTransfer.getData("text/plain"), row.__rowId);
        }}
      >
        <td className="row-number">
          <button
            className="row-drag-handle"
            draggable={!groupBy}
            disabled={Boolean(groupBy)}
            title={groupBy ? "Turn off grouping to reorder rows" : "Drag to reorder"}
            aria-label={`Drag row ${start + rowIndex + 1} to reorder`}
            onDragStart={event => {
              setDraggedRowId(row.__rowId);
              event.dataTransfer.effectAllowed = "move";
              event.dataTransfer.setData("text/plain", row.__rowId);
            }}
            onDragEnd={() => {
              setDraggedRowId(null);
              setDragOverRowId(null);
            }}
          >
            <GripVertical size={13} />
          </button>
          <span>{start + rowIndex + 1}</span>
        </td>
        {visibleColumns.map(column => {
          const cellStyle = cellFormat(column, row[column]);
          const format = cellStyle.priority === 2 ? cellStyle : rowStyle.className ? rowStyle : cellStyle;
          return (
            <td
              key={column}
              className={format.className}
              style={format.style}
              title={text(row[column])}
              onContextMenu={event => showContextMenu(event, "cell", column, row[column])}
            >
              {text(row[column]) || <span className="blank">—</span>}
            </td>
          );
        })}
      </tr>
      );
    });
  }

  return (
    <div
      className={`app-shell ${sidebarOpen ? "" : "sidebar-closed"}`}
      style={{ "--sidebar-width": `${sidebarWidth}px` }}
    >
      <aside className="sidebar">
        <div className="brand">
          <div className="brand-mark"><MyDataMark /></div>
          <div><strong>MyData</strong><span>Explore Excel, CSV, and more</span></div>
        </div>
        <button className="primary wide" onClick={() => inputRef.current?.click()}><Upload size={16} /> Import files</button>
        <input ref={inputRef} type="file" accept=".xlsx,.xlsm,.xls,.csv" multiple hidden onChange={event => { importFiles(event.target.files); event.target.value = ""; }} />
        <div className="source-heading"><span>Data sources</span>{datasets.length > 0 && <span>{datasets.length}</span>}</div>
        <nav className="source-list">
          {datasets.length > 1 && (
            <button className={activeId === "all" ? "active" : ""} onClick={() => setActiveId("all")}>
              <Layers3 size={16} /><span><strong>All sheets</strong><small>{datasets.reduce((sum, dataset) => sum + dataset.rows.length, 0).toLocaleString()} rows</small></span>
            </button>
          )}
          {datasets.map(dataset => (
            <button key={dataset.id} className={activeId === dataset.id ? "active" : ""} onClick={() => setActiveId(dataset.id)}>
              <Sheet size={16} /><span><strong>{dataset.sheet}</strong><small>{dataset.workbook} · {dataset.rows.length.toLocaleString()} rows</small></span>
            </button>
          ))}
        </nav>
        <div className="sidebar-footer">
          {datasets.length ? (
            <button className="ghost wide" onClick={clearData}><Trash2 size={15} /> Clear all data</button>
          ) : (
            <button className="ghost wide" onClick={loadSample}><BarChart3 size={15} /> Load sample data</button>
          )}
          <p>Files stay in this browser. Nothing is uploaded.</p>
        </div>
      </aside>
      <div
        className="sidebar-resizer"
        role="separator"
        aria-orientation="vertical"
        aria-label="Resize source sidebar"
        title="Drag to resize; double-click to reset"
        onPointerDown={startSidebarResize}
        onDoubleClick={() => setSidebarWidth(260)}
      />

      <main className="main">
        <header className="topbar">
          <IconButton label={sidebarOpen ? "Hide sources" : "Show sources"} onClick={() => setSidebarOpen(value => !value)}>
            <ChevronLeft className={sidebarOpen ? "" : "flip"} size={18} />
          </IconButton>
          <div className="title-block">
            <h1>{selected ? selected.sheet : "MyData"}</h1>
            {selected?.sourceUrl ? (
              <button
                className="workbook-link"
                type="button"
                title={`Open ${selected.workbook} in MyData`}
                onClick={() => setActiveId(`file:${selected.sourceUrl}`)}
              >
                {selected.workbook} <ChevronRight size={11} />
              </button>
            ) : (
              <p>{selected ? selected.workbook : "Import one or more workbooks to explore their data"}</p>
            )}
          </div>
          <div className="top-actions">
            <button className="button" disabled={!selected} onClick={clearFilters}><RotateCcw size={15} /> Reset view</button>
            <div className="menu-wrap" ref={exportMenuRef}>
              <button className={`button ${exportOpen ? "active" : ""}`} disabled={!selected} onClick={() => { setExportOpen(value => !value); setColorMenuOpen(false); setColumnsOpen(false); }}>
                <Download size={15} /> Export <ChevronDown size={14} />
              </button>
              {exportOpen && <div className="popover export-menu">
                <button onClick={exportFilteredExcel}><FileSpreadsheet size={16} /><span><strong>Filtered view</strong><small>Excel workbook (.xlsx)</small></span></button>
                <button onClick={exportFilteredCsv}><Sheet size={16} /><span><strong>Filtered view</strong><small>Comma-separated values (.csv)</small></span></button>
                <button onClick={exportFilteredHtml}><FileCode2 size={16} /><span><strong>Filtered view</strong><small>Standalone web page (.html)</small></span></button>
                <button onClick={exportAllSheets}><Layers3 size={16} /><span><strong>All imported sheets</strong><small>Excel workbook (.xlsx)</small></span></button>
              </div>}
            </div>
            <button className="primary" onClick={() => inputRef.current?.click()}><Plus size={16} /> Add workbooks</button>
          </div>
        </header>

        {loadedFiles.length > 0 && <div className="loaded-files" aria-label="Loaded files">
          <strong>Loaded {loadedFiles.length === 1 ? "file" : "files"}</strong>
          <div>
            {loadedFiles.map(file => {
              const fileSheets = datasets.filter(dataset =>
                (dataset.sourceUrl || `name:${dataset.workbook}`) === file.key
              );
              const active = activeId === `file:${file.key}` || fileSheets.some(dataset => dataset.id === activeId);
              return <div className="loaded-file-wrap" key={file.key}>
                <button
                  className={active ? "active" : ""}
                  type="button"
                  aria-expanded={loadedFileMenuKey === file.key}
                  title={`Choose a sheet from ${file.name}`}
                  onClick={() => setLoadedFileMenuKey(current => current === file.key ? "" : file.key)}
                >
                  <FileSpreadsheet size={13} />
                  <span>{file.name}</span>
                  <small>{file.sheets} sheet{file.sheets === 1 ? "" : "s"}</small>
                  <ChevronDown size={11} />
                </button>
                {loadedFileMenuKey === file.key && <div className="loaded-file-menu">
                  {fileSheets.length > 1 && <button
                    className={activeId === `file:${file.key}` ? "active" : ""}
                    onClick={() => { setActiveId(`file:${file.key}`); setLoadedFileMenuKey(""); }}
                  >
                    <Layers3 size={14} /><span><strong>All workbook sheets</strong><small>{fileSheets.reduce((sum, dataset) => sum + dataset.rows.length, 0)} rows</small></span>
                  </button>}
                  {fileSheets.map(dataset => <button
                    key={dataset.id}
                    className={activeId === dataset.id ? "active" : ""}
                    onClick={() => { setActiveId(dataset.id); setLoadedFileMenuKey(""); }}
                  >
                    <Sheet size={14} /><span><strong>{dataset.sheet}</strong><small>{dataset.rows.length} rows</small></span>
                  </button>)}
                </div>}
              </div>;
            })}
          </div>
        </div>}

        {notice && <div className={`toast ${notice.type}`} role={notice.type === "error" ? "alert" : "status"}>
          <span>{notice.text}</span>
          <IconButton label="Dismiss" onClick={() => setNotice(null)}><X size={14} /></IconButton>
        </div>}

        {!selected ? (
          <section
            className={`drop-zone ${dragging ? "dragging" : ""}`}
            onDragOver={event => { event.preventDefault(); setDragging(true); }}
            onDragLeave={() => setDragging(false)}
            onDrop={event => { event.preventDefault(); setDragging(false); importFiles(event.dataTransfer.files); }}
          >
            <div className="drop-icon"><FolderOpen size={30} /></div>
            <h2>Drop Excel files here</h2>
            <p>Open multiple workbooks, switch between sheets, or combine everything in one table.</p>
            <div><button className="primary" onClick={() => inputRef.current?.click()}><Upload size={16} /> Choose files</button><button className="button" onClick={loadSample}>Try sample data</button></div>
            <small>Supports .xlsx, .xlsm, .xls, and .csv</small>
          </section>
        ) : (
          <>
            <section className="metrics">
              <div><span>Rows</span><strong>{selected.rows.length.toLocaleString()}</strong></div>
              <div><span>Matching</span><strong>{filteredRows.length.toLocaleString()}</strong></div>
              <div><span>Columns</span><strong>{columns.length.toLocaleString()}</strong></div>
              <div><span>Sheets loaded</span><strong>{datasets.length.toLocaleString()}</strong></div>
            </section>

            <section className="toolbar">
              <div className="search-box">
                <Search size={16} />
                <input
                  ref={searchInputRef}
                  value={query}
                  aria-invalid={Boolean(smartQuery.error)}
                  title={smartQuery.error || 'Words are ANDed; use OR, NOT or !, "quoted phrases", ^regex shorthand, or /regex/i'}
                  onChange={event => { setQuery(event.target.value); setPage(1); }}
                  placeholder='Search every column…'
                />
                {!query && <kbd title="Press / or Ctrl+Shift+F to focus">/</kbd>}
                {query && <IconButton label="Clear search" onClick={() => setQuery("")}><X size={14} /></IconButton>}
              </div>
              <button className={`button ${filterPanelOpen ? "active" : ""}`} onClick={() => setFilterPanelOpen(value => !value)}><Filter size={15} /> Filters {filterCount > 0 && <span className="count">{filterCount}</span>}</button>
              <button
                className={`button ${quickFiltersVisible ? "active" : ""}`}
                onClick={() => setQuickFiltersVisible(value => !value)}
                title={quickFiltersVisible ? "Hide the filter boxes below column headers" : "Show filter boxes below column headers"}
              >
                <ListFilter size={15} /> Column filters
              </button>
              <div className="menu-wrap" ref={colorMenuRef}>
                <button
                  className={`button ${formatCount ? "active" : ""}`}
                  onClick={() => { setColorMenuOpen(value => !value); setExportOpen(false); setColumnsOpen(false); }}
                  title="Choose columns whose equal values receive the same color"
                >
                  <Palette size={15} /> Color values
                  {formatCount > 0 && <span className="count">{formatCount}</span>}
                  <ChevronDown size={14} />
                </button>
                {colorMenuOpen && <div className="popover color-menu">
                  <div className="popover-head">
                    <strong>Color by columns</strong>
                    <div><button onClick={() => setAllColorColumns(true)}>All</button><button onClick={clearAllFormatting}>Clear</button></div>
                  </div>
                  <p>Equal values receive the same color.</p>
                  {columns.map(column => (
                    <div className="color-column-row" key={column}>
                      <label title={`Color values in ${displayColumn(column)}`}>
                        <input type="checkbox" checked={coloredColumns.has(column)} onChange={() => toggleColorColumn(column)} />
                        <span className="check"><Check size={12} /></span>
                        <span>{displayColumn(column)}</span>
                      </label>
                      {coloredColumns.has(column) && <>
                        <button className={`color-scope ${rowColorColumn === column ? "rows" : ""}`} onClick={() => toggleColumnColorScope(column)} title="Toggle between coloring cells and entire rows">
                          {rowColorColumn === column ? "Rows" : "Cells"}
                        </button>
                        <button className="edit-value-colors" onClick={() => openValueColorEditor(column)}>Edit colors</button>
                      </>}
                    </div>
                  ))}
                  {customizingColorColumn && coloredColumns.has(customizingColorColumn) && <div className="value-color-editor">
                    <div className="value-color-editor-head">
                      <div><strong>{displayColumn(customizingColorColumn)}</strong><small>Choose a color for each distinct value</small></div>
                      <IconButton label="Close value color editor" onClick={() => setCustomizingColorColumn("")}><X size={13} /></IconButton>
                    </div>
                    <input className="value-color-search" type="search" value={valueColorSearch} onChange={event => setValueColorSearch(event.target.value)} placeholder="Find a value…" />
                    <div className="value-color-list">
                      {customizableValues.map(entry => {
                        const currentColor = valueColorOverrides[customizingColorColumn]?.[entry.key];
                        const defaultColor = VALUE_COLORS[valueColorIndex(entry.value, paletteSeeds[customizingColorColumn] || 0)];
                        return <div className="value-color-row" key={entry.key}>
                          <input type="color" value={currentColor || defaultColor} onChange={event => setValueColor(customizingColorColumn, entry.key, event.target.value)} aria-label={`Color for ${entry.value}`} />
                          <span title={entry.value}>{entry.value}</span>
                          <small>{entry.count}</small>
                          {currentColor && <IconButton label={`Reset color for ${entry.value}`} onClick={() => resetValueColor(customizingColorColumn, entry.key)}><RotateCcw size={12} /></IconButton>}
                        </div>;
                      })}
                      {!customizableValues.length && <p>No values match.</p>}
                    </div>
                  </div>}
                  <div className="conditional-head">
                    <div><strong>Conditional formatting</strong><small>Rules override automatic colors</small></div>
                    <button onClick={addConditionalRule}><Plus size={13} /> Add rule</button>
                  </div>
                  {conditionalRules.length === 0 && <p className="conditional-empty">No formatting rules yet.</p>}
                  {conditionalRules.map(rule => {
                    const noValue = rule.operator === "blank" || rule.operator === "notBlank";
                    return <div className="conditional-rule" key={rule.id}>
                      <select value={rule.column} onChange={event => updateConditionalRule(rule.id, { column: event.target.value })}>
                        {columns.map(column => <option key={column} value={column}>{displayColumn(column)}</option>)}
                      </select>
                      <select value={rule.operator} onChange={event => updateConditionalRule(rule.id, { operator: event.target.value })}>
                        {OPERATORS.map(([value, label]) => <option key={value} value={value}>{label}</option>)}
                      </select>
                      <input className="conditional-value" disabled={noValue} value={rule.value} onChange={event => updateConditionalRule(rule.id, { value: event.target.value })} placeholder={noValue ? "No value" : "Value"} />
                      <input className="conditional-color" type="color" value={rule.color} onChange={event => updateConditionalRule(rule.id, { color: event.target.value })} aria-label="Rule color" />
                      <label className="conditional-row-scope">
                        <input type="checkbox" checked={Boolean(rule.applyToRow)} onChange={event => updateConditionalRule(rule.id, { applyToRow: event.target.checked })} />
                        <span className="check"><Check size={12} /></span>
                        Color entire row
                      </label>
                      <div className="conditional-actions">
                        <IconButton label="Duplicate formatting rule" onClick={() => duplicateConditionalRule(rule)}><Copy size={14} /></IconButton>
                        <IconButton label="Remove formatting rule" onClick={() => removeConditionalRule(rule.id)}><Trash2 size={14} /></IconButton>
                      </div>
                    </div>;
                  })}
                </div>}
              </div>
              <div className="menu-wrap" ref={columnsMenuRef}>
                <button className={`button ${columnsOpen ? "active" : ""}`} onClick={() => { setColumnsOpen(value => !value); setExportOpen(false); setColorMenuOpen(false); }}><Columns3 size={15} /> Columns <ChevronDown size={14} /></button>
                {columnsOpen && <div className="popover column-menu">
                  <div className="popover-head">
                    <strong>Columns</strong>
                    <div><button onClick={() => setHiddenColumns(new Set())}>Show all</button><button onClick={() => setColumnAliases(current => ({ ...current, [activeId]: {} }))}>Reset names</button></div>
                  </div>
                  <div className="column-size-actions">
                    <button onClick={fitVisibleColumnsToContent}>Fit visible to content</button>
                    <button onClick={() => setColumnWidths({})}>Reset widths</button>
                  </div>
                  {columns.map(column => (
                    <div className="column-menu-row" key={column}>
                      <label className="column-toggle" title={hiddenColumns.has(column) ? "Show column" : "Hide column"}>
                        <input type="checkbox" checked={!hiddenColumns.has(column)} onChange={() => toggleColumn(column)} />
                        <span className="check"><Check size={12} /></span>
                      </label>
                      {editingColumn === column ? (
                        <form className="column-rename-form" onSubmit={event => { event.preventDefault(); commitColumnRename(column); }}>
                          <input className="column-rename-input" value={renameDraft} onChange={event => setRenameDraft(event.target.value)} autoFocus />
                          <IconButton label="Save column name" type="submit"><Check size={14} /></IconButton>
                          <IconButton label="Cancel rename" type="button" onClick={() => setEditingColumn("")}><X size={14} /></IconButton>
                        </form>
                      ) : (
                        <>
                          <span className="column-name" title={column !== displayColumn(column) ? `Source name: ${column}` : column}>{displayColumn(column)}</span>
                          <IconButton label={`Rename ${displayColumn(column)}`} className="rename-column" onClick={() => beginColumnRename(column)}><Pencil size={13} /></IconButton>
                        </>
                      )}
                    </div>
                  ))}
                </div>}
              </div>
              <label className="select-wrap"><Group size={15} /><select value={groupBy} onChange={event => { setGroupBy(event.target.value); setPage(1); }}><option value="">No grouping</option>{columns.map(column => <option key={column} value={column}>Group by {displayColumn(column)}</option>)}</select><ChevronDown size={14} /></label>
              <span className="toolbar-spacer" />
              <span className={`match-label ${smartQuery.error ? "query-error" : ""}`}>
                {smartQuery.error || `${filteredRows.length.toLocaleString()} of ${selected.rows.length.toLocaleString()} rows`}
              </span>
            </section>

            {filterPanelOpen && (
              <section className="filter-panel">
                <div className="filter-panel-head"><div><SlidersHorizontal size={16} /><strong>Advanced filters</strong><span>All conditions must match</span></div><button className="link-button" onClick={clearFilters}>Clear all</button></div>
                {filters.length === 0 && <p className="filter-empty">Add conditions for precise filtering, or type directly below any column header.</p>}
                {filters.map(filter => {
                  const noValue = filter.operator === "blank" || filter.operator === "notBlank";
                  return <div className="filter-rule" key={filter.id}>
                    <select value={filter.column} onChange={event => updateFilter(filter.id, { column: event.target.value })}>{columns.map(column => <option key={column} value={column}>{displayColumn(column)}</option>)}</select>
                    <select value={filter.operator} onChange={event => updateFilter(filter.id, { operator: event.target.value })}>{OPERATORS.map(([value, label]) => <option key={value} value={value}>{label}</option>)}</select>
                    <input disabled={noValue} value={filter.value} onChange={event => updateFilter(filter.id, { value: event.target.value })} placeholder={noValue ? "No value needed" : "Value"} />
                    <IconButton label="Remove filter" onClick={() => setFilters(current => current.filter(item => item.id !== filter.id))}><Trash2 size={15} /></IconButton>
                  </div>;
                })}
                <button className="button add-filter" onClick={addFilter}><Plus size={15} /> Add condition</button>
              </section>
            )}

            <section className="table-card">
              <div className="table-scroll">
                <table style={{ width: `${tableWidth}px` }}>
                  <colgroup>
                    <col style={{ width: "64px" }} />
                    {visibleColumns.map(column => (
                      <col key={column} style={{ width: `${columnWidths[column] || defaultColumnWidth(column)}px` }} />
                    ))}
                  </colgroup>
                  <thead>
                    <tr className="header-row">
                      <th className="row-number">#</th>
                      {visibleColumns.map(column => (
                        <th
                          key={column}
                          onClick={() => toggleSort(column)}
                          onContextMenu={event => showContextMenu(event, "header", column)}
                        >
                          <span>{displayColumn(column)}</span>
                          {sort.column === column && (sort.direction === "asc" ? <ArrowUp size={14} /> : <ArrowDown size={14} />)}
                          <span
                            className="column-resizer"
                            role="separator"
                            aria-orientation="vertical"
                            aria-label={`Resize ${displayColumn(column)} column`}
                            title="Drag to resize; double-click to fit content"
                            onClick={event => event.stopPropagation()}
                            onPointerDown={event => startColumnResize(event, column)}
                            onDoubleClick={event => {
                              event.preventDefault();
                              event.stopPropagation();
                              fitColumnToContent(column);
                            }}
                          />
                        </th>
                      ))}
                    </tr>
                    {quickFiltersVisible && <tr className="quick-filter-row">
                      <th className="row-number">
                        <button
                          className="filter-row-toggle"
                          type="button"
                          title="Hide column filters"
                          aria-label="Hide column filters"
                          onClick={() => setQuickFiltersVisible(false)}
                        >
                          <Filter size={13} />
                        </button>
                      </th>
                      {visibleColumns.map(column => <th key={column}><input value={quickFilters[column] || ""} onChange={event => { setQuickFilters(current => ({ ...current, [column]: event.target.value })); setPage(1); }} placeholder="Filter…" aria-label={`Filter ${displayColumn(column)}`} /></th>)}
                    </tr>}
                  </thead>
                  <tbody>
                    {groupBy && groups ? groups.map(([groupName, rows]) => {
                      const collapsed = collapsedGroups.has(groupName);
                      return <React.Fragment key={groupName}>
                        <tr className="group-row" onClick={() => setCollapsedGroups(current => { const next = new Set(current); next.has(groupName) ? next.delete(groupName) : next.add(groupName); return next; })}>
                          <td colSpan={visibleColumns.length + 1}><ChevronDown className={collapsed ? "collapsed" : ""} size={15} /><strong>{groupName}</strong><span>{rows.length.toLocaleString()} rows</span></td>
                        </tr>
                        {!collapsed && renderRows(rows)}
                      </React.Fragment>;
                    }) : renderRows(pageRows)}
                    {!filteredRows.length && <tr><td className="empty-table" colSpan={visibleColumns.length + 1}>No rows match the current filters.</td></tr>}
                  </tbody>
                </table>
              </div>
              {!groupBy && <div className="pagination">
                <span>Rows per page</span>
                <select value={pageSize} onChange={event => { setPageSize(Number(event.target.value)); setPage(1); }}><option value={25}>25</option><option value={50}>50</option><option value={100}>100</option><option value={0}>All</option></select>
                <span className="page-summary">{filteredRows.length ? `${start + 1}–${Math.min(start + (pageSize || filteredRows.length), filteredRows.length)} of ${filteredRows.length.toLocaleString()}` : "0 rows"}</span>
                <IconButton label="Previous page" disabled={page <= 1} onClick={() => setPage(value => value - 1)}><ChevronLeft size={16} /></IconButton>
                <IconButton label="Next page" disabled={page >= totalPages} onClick={() => setPage(value => value + 1)}><ChevronRight size={16} /></IconButton>
              </div>}
            </section>
          </>
        )}
      </main>
      {contextMenu && <div
        className="context-menu"
        role="menu"
        aria-label={`${displayColumn(contextMenu.column)} actions`}
        style={{ left: contextMenu.x, top: contextMenu.y }}
      >
        <div className="context-menu-title">
          <strong>{displayColumn(contextMenu.column)}</strong>
          {contextMenu.type === "cell" && <span title={text(contextMenu.value)}>{text(contextMenu.value) || "(Blank)"}</span>}
        </div>
        {contextMenu.type === "cell" ? <>
          <button role="menuitem" onClick={copyContextValue}><Copy size={15} /><span>Copy cell value</span></button>
          <div className="context-divider" />
          <button role="menuitem" onClick={() => addContextFilter(isBlank(contextMenu.value) ? "blank" : "equals")}><Filter size={15} /><span>Filter to this value</span></button>
          {!isBlank(contextMenu.value) && <button role="menuitem" onClick={() => addContextFilter("notEquals")}><FilterX size={15} /><span>Exclude this value</span></button>}
          <button role="menuitem" onClick={colorContextValue}><PaintBucket size={15} /><span>Color matching cells</span></button>
          <button role="menuitem" onClick={colorColumnFromContext}><Palette size={15} /><span>Color this column</span></button>
          <button role="menuitem" onClick={() => clearColumnColors(contextMenu.column)}><RotateCcw size={15} /><span>Clear colors in column</span></button>
          <button role="menuitem" onClick={() => clearColumnFilters(contextMenu.column)}><RotateCcw size={15} /><span>Clear filters in column</span></button>
        </> : <>
          <button role="menuitem" onClick={() => { setSort({ column: contextMenu.column, direction: "asc" }); setContextMenu(null); }}><ArrowUp size={15} /><span>Sort ascending</span></button>
          <button role="menuitem" onClick={() => { setSort({ column: contextMenu.column, direction: "desc" }); setContextMenu(null); }}><ArrowDown size={15} /><span>Sort descending</span></button>
          <div className="context-divider" />
          <button role="menuitem" onClick={() => { fitColumnToContent(contextMenu.column); setContextMenu(null); }}><Columns3 size={15} /><span>Fit column to content</span></button>
          <button role="menuitem" onClick={colorColumnFromContext}><Palette size={15} /><span>Color this column</span></button>
          <button role="menuitem" onClick={() => clearColumnColors(contextMenu.column)}><RotateCcw size={15} /><span>Clear colors in column</span></button>
          <button role="menuitem" onClick={() => { setColumnsOpen(true); beginColumnRename(contextMenu.column); setContextMenu(null); }}><Pencil size={15} /><span>Rename column</span></button>
          <button role="menuitem" onClick={() => { toggleColumn(contextMenu.column); setContextMenu(null); }}><EyeOff size={15} /><span>Hide column</span></button>
        </>}
      </div>}
    </div>
  );
}

createRoot(document.getElementById("root")).render(<App />);
