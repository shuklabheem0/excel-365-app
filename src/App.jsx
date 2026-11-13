import React, { useState, useEffect, useRef, useMemo } from 'react';
import {
  Save, Undo2, Redo2, Bold, Italic, Underline, AlignLeft, AlignCenter, AlignRight,
  Plus, Trash2, Copy, Clipboard, Scissors, Download, Upload, Moon, Sun,
  Table, BarChart3, FunctionSquare, Freeze, Grid, Search, ArrowUpDown,
  Merge, DollarSign, Percent, Hash, FileSpreadsheet, X
} from 'lucide-react';

const COLS = 26;
const ROWS = 100;

const colToLetter = (c) => String.fromCharCode(65 + c);
const letterToCol = (l) => l.toUpperCase().charCodeAt(0) - 65;
const cellId = (r, c) => `${colToLetter(c)}${r + 1}`;
const parseCellId = (id) => {
  const m = id.replace(/\$/g, '').match(/^([A-Z]+)([0-9]+)$/i);
  if (!m) return null;
  return { c: letterToCol(m[1]), r: parseInt(m[2], 10) - 1 };
};

const createEmptySheet = (name) => ({
  name,
  cells: {}, // key: "A1", value: { raw: "", style: {} }
  colWidths: {},
  rowHeights: {},
  merges: [],
  freezeRow: false
});

const defaultSampleSheet = () => {
  const s = createEmptySheet('Sales_MIS');
  const data = {
    A1: { raw: 'Month', style: { bold: true, bg: '#107C41', color: '#ffffff', align: 'center' } },
    B1: { raw: 'Target Units', style: { bold: true, bg: '#107C41', color: '#ffffff', align: 'center' } },
    C1: { raw: 'Achieved Units', style: { bold: true, bg: '#107C41', color: '#ffffff', align: 'center' } },
    D1: { raw: 'Achievement %', style: { bold: true, bg: '#107C41', color: '#ffffff', align: 'center' } },
    E1: { raw: 'Incentive (INR)', style: { bold: true, bg: '#107C41', color: '#ffffff', align: 'center' } },
    A2: { raw: 'January' }, B2: { raw: '120' }, C2: { raw: '135' }, D2: { raw: '=ROUND((C2/B2)*100, 1)', style: { align: 'right' } }, E2: { raw: '=IF(C2>=B2, C2*500, 0)', style: { format: 'currency', align: 'right' } },
    A3: { raw: 'February' }, B3: { raw: '130' }, C3: { raw: '125' }, D3: { raw: '=ROUND((C3/B3)*100, 1)', style: { align: 'right' } }, E3: { raw: '=IF(C3>=B3, C3*500, 0)', style: { format: 'currency', align: 'right' } },
    A4: { raw: 'March' }, B4: { raw: '150' }, C4: { raw: '168' }, D4: { raw: '=ROUND((C4/B4)*100, 1)', style: { align: 'right' } }, E4: { raw: '=IF(C4>=B4, C4*500, 0)', style: { format: 'currency', align: 'right' } },
    A5: { raw: 'Total', style: { bold: true, bg: '#e8f5e9' } },
    B5: { raw: '=SUM(B2:B4)', style: { bold: true, bg: '#e8f5e9' } },
    C5: { raw: '=SUM(C2:C4)', style: { bold: true, bg: '#e8f5e9' } },
    D5: { raw: '=ROUND(AVERAGE(D2:D4), 1)', style: { bold: true, bg: '#e8f5e9', align: 'right' } },
    E5: { raw: '=SUM(E2:E4)', style: { bold: true, bg: '#e8f5e9', format: 'currency', align: 'right' } },
  };
  s.cells = data;
  s.colWidths = { 0: 130, 1: 120, 2: 135, 3: 135, 4: 150 };
  return s;
};

export default function App() {
  const [workbookName, setWorkbookName] = useState('Book1 - Excel 365 Offline');
  const [sheets, setSheets] = useState(() => {
    const saved = localStorage.getItem('excel365_offline_wb');
    if (saved) {
      try { return JSON.parse(saved); } catch (e) {}
    }
    return [defaultSampleSheet(), createEmptySheet('Sheet2')];
  });
  const [activeSheetIdx, setActiveSheetIdx] = useState(0);
  const [activeTab, setActiveTab] = useState('Home');
  const [darkMode, setDarkMode] = useState(false);

  const [selected, setSelected] = useState({ r: 0, c: 0 });
  const [selEnd, setSelEnd] = useState({ r: 0, c: 0 });
  const [isSelecting, setIsSelecting] = useState(false);
  const [editingCell, setEditingCell] = useState(null);
  const [editValue, setEditValue] = useState('');
  const [clipboard, setClipboard] = useState(null);
  const [history, setHistory] = useState([]);
  const [redoStack, setRedoStack] = useState([]);
  const [showChartModal, setShowChartModal] = useState(false);
  const [chartType, setChartType] = useState('bar');
  const [searchQuery, setSearchQuery] = useState('');

  const inputRef = useRef(null);
  const fileInputRef = useRef(null);

  const currentSheet = sheets[activeSheetIdx] || sheets[0];

  // Auto-save locally for offline persistence
  useEffect(() => {
    localStorage.setItem('excel365_offline_wb', JSON.stringify(sheets));
  }, [sheets]);

  const pushHistory = (prevSheets) => {
    setHistory((h) => [...h.slice(-29), JSON.stringify(prevSheets)]);
    setRedoStack([]);
  };

  const undo = () => {
    if (history.length === 0) return;
    const last = history[history.length - 1];
    setRedoStack((r) => [...r, JSON.stringify(sheets)]);
    setHistory((h) => h.slice(0, -1));
    setSheets(JSON.parse(last));
  };

  const redo = () => {
    if (redoStack.length === 0) return;
    const next = redoStack[redoStack.length - 1];
    setHistory((h) => [...h, JSON.stringify(sheets)]);
    setRedoStack((r) => r.slice(0, -1));
    setSheets(JSON.parse(next));
  };

  // Formula Evaluator
  const evalFormula = (raw, sheetObj, visited = new Set()) => {
    if (typeof raw !== 'string' || !raw.startsWith('=')) return raw;
    const expr = raw.slice(1).trim();

    const expandRange = (rangeStr, targetSheet) => {
      const [startStr, endStr] = rangeStr.split(':');
      const s = parseCellId(startStr);
      const e = parseCellId(endStr);
      if (!s || !e) return [];
      const vals = [];
      for (let r = Math.min(s.r, e.r); r <= Math.max(s.r, e.r); r++) {
        for (let c = Math.min(s.c, e.c); c <= Math.max(s.c, e.c); c++) {
          const id = cellId(r, c);
          if (visited.has(`${targetSheet.name}!${id}`)) continue;
          const nextVisited = new Set(visited);
          nextVisited.add(`${targetSheet.name}!${id}`);
          const cellRaw = targetSheet.cells[id]?.raw ?? '';
          const v = evalFormula(cellRaw, targetSheet, nextVisited);
          const num = parseFloat(v);
          if (!isNaN(num)) vals.push(num);
        }
      }
      return vals;
    };

    try {
      // Handle Excel functions
      let processed = expr.replace(/([A-Z]+)\(([^()]+)\)/gi, (_, fn, argsStr) => {
        const funcName = fn.toUpperCase();
        if (funcName === 'IF') {
          const parts = argsStr.split(',').map((p) => p.trim());
          if (parts.length < 2) return '0';
          const condEval = evalFormula('=' + parts[0], sheetObj, visited);
          return condEval ? evalFormula('=' + parts[1], sheetObj, visited) : (parts[2] ? evalFormula('=' + parts[2], sheetObj, visited) : 0);
        }
        if (funcName === 'ROUND') {
          const [valPart, decPart] = argsStr.split(',');
          const val = parseFloat(evalFormula('=' + valPart, sheetObj, visited)) || 0;
          const dec = parseInt(decPart || '0', 10);
          return Number(Math.round(val + 'e' + dec) + 'e-' + dec);
        }
        if (funcName === 'TODAY') {
          return `"${new Date().toLocaleDateString('en-IN')}"`;
        }
        if (funcName === 'CONCAT') {
          return `"${argsStr.split(',').map(a => String(evalFormula('=' + a.trim(), sheetObj, visited)).replace(/^"|"$/g, '')).join('')}"`;
        }

        // Range aggregate functions
        const nums = [];
        argsStr.split(',').forEach((arg) => {
          const trimmed = arg.trim();
          if (trimmed.includes(':')) {
            nums.push(...expandRange(trimmed, sheetObj));
          } else {
            const v = parseFloat(evalFormula('=' + trimmed, sheetObj, visited));
            if (!isNaN(v)) nums.push(v);
          }
        });

        if (funcName === 'SUM') return nums.reduce((a, b) => a + b, 0);
        if (funcName === 'AVERAGE') return nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
        if (funcName === 'COUNT') return nums.length;
        if (funcName === 'MAX') return nums.length ? Math.max(...nums) : 0;
        if (funcName === 'MIN') return nums.length ? Math.min(...nums) : 0;
        return 0;
      });

      // Replace cross-sheet refs like Sheet2!A1
      processed = processed.replace(/([A-Za-z0-9_]+)!\$?([A-Z]+)\$?([0-9]+)/g, (_, shName, colStr, rowStr) => {
        const target = sheets.find((s) => s.name.toLowerCase() === shName.toLowerCase());
        if (!target) return 0;
        const id = `${colStr.toUpperCase()}${rowStr}`;
        const v = evalFormula(target.cells[id]?.raw ?? '', target, visited);
        return isNaN(parseFloat(v)) ? `"${v}"` : parseFloat(v);
      });

      // Replace same-sheet cell refs
      processed = processed.replace(/\$?([A-Z]+)\$?([0-9]+)/g, (match, colStr, rowStr) => {
        const id = `${colStr.toUpperCase()}${rowStr}`;
        if (visited.has(`${sheetObj.name}!${id}`)) return 0;
        const nextVisited = new Set(visited);
        nextVisited.add(`${sheetObj.name}!${id}`);
        const rawVal = sheetObj.cells[id]?.raw ?? '';
        if (rawVal === '') return 0;
        const evaluated = evalFormula(rawVal, sheetObj, nextVisited);
        return isNaN(parseFloat(evaluated)) ? `"${evaluated}"` : parseFloat(evaluated);
      });

      // Convert Excel = comparison to JS ==
      processed = processed.replace(/([^<>!])=([^=])/g, '$1==$2');

      // eslint-disable-next-line no-new-func
      const res = Function(`"use strict"; return (${processed})`)();
      return res;
    } catch (e) {
      return '#ERROR!';
    }
  };

  const getDisplayValue = (id, sheetObj = currentSheet) => {
    const cell = sheetObj.cells[id];
    if (!cell || cell.raw === undefined || cell.raw === '') return '';
    const val = evalFormula(cell.raw, sheetObj);
    if (typeof val === 'number' && !isNaN(val)) {
      if (cell.style?.format === 'currency') {
        return '₹' + val.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      }
      if (cell.style?.format === 'percent') {
        return (val * 100).toFixed(1) + '%';
      }
    }
    return val;
  };

  const updateCell = (id, raw, stylePatch = null) => {
    pushHistory(sheets);
    setSheets((prev) =>
      prev.map((sh, idx) => {
        if (idx !== activeSheetIdx) return sh;
        const existing = sh.cells[id] || { raw: '', style: {} };
        return {
          ...sh,
          cells: {
            ...sh.cells,
            [id]: {
              raw: raw !== undefined ? raw : existing.raw,
              style: stylePatch ? { ...existing.style, ...stylePatch } : existing.style
            }
          }
        };
      })
    );
  };

  const getSelectedRange = () => {
    const r1 = Math.min(selected.r, selEnd.r);
    const r2 = Math.max(selected.r, selEnd.r);
    const c1 = Math.min(selected.c, selEnd.c);
    const c2 = Math.max(selected.c, selEnd.c);
    return { r1, r2, c1, c2 };
  };

  const applyStyleToSelection = (stylePatch) => {
    pushHistory(sheets);
    const { r1, r2, c1, c2 } = getSelectedRange();
    setSheets((prev) =>
      prev.map((sh, idx) => {
        if (idx !== activeSheetIdx) return sh;
        const nextCells = { ...sh.cells };
        for (let r = r1; r <= r2; r++) {
          for (let c = c1; c <= c2; c++) {
            const id = cellId(r, c);
            const cur = nextCells[id] || { raw: '', style: {} };
            nextCells[id] = { ...cur, style: { ...cur.style, ...stylePatch } };
          }
        }
        return { ...sh, cells: nextCells };
      })
    );
  };

  const activeCellId = cellId(selected.r, selected.c);
  const activeCellObj = currentSheet.cells[activeCellId] || { raw: '', style: {} };

  // Status bar live aggregations
  const selectionStats = useMemo(() => {
    const { r1, r2, c1, c2 } = getSelectedRange();
    const nums = [];
    let count = 0;
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        const val = evalFormula(currentSheet.cells[cellId(r, c)]?.raw ?? '', currentSheet);
        if (val !== '' && val !== undefined) {
          count++;
          const n = parseFloat(val);
          if (!isNaN(n)) nums.push(n);
        }
      }
    }
    const sum = nums.reduce((a, b) => a + b, 0);
    const avg = nums.length ? sum / nums.length : 0;
    return { count, sum, avg, hasNums: nums.length > 0 };
  }, [selected, selEnd, currentSheet]);

  // Keyboard navigation
  const handleKeyDown = (e) => {
    if (editingCell) {
      if (e.key === 'Enter') {
        updateCell(editingCell, editValue);
        setEditingCell(null);
        setSelected((s) => ({ r: Math.min(ROWS - 1, s.r + 1), c: s.c }));
        setSelEnd((s) => ({ r: Math.min(ROWS - 1, s.r + 1), c: s.c }));
      } else if (e.key === 'Escape') {
        setEditingCell(null);
      }
      return;
    }

    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') {
      e.preventDefault();
      undo();
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') {
      e.preventDefault();
      redo();
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'b') {
      e.preventDefault();
      applyStyleToSelection({ bold: !activeCellObj.style?.bold });
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'c') {
      setClipboard({ ...activeCellObj });
      return;
    }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'v' && clipboard) {
      updateCell(activeCellId, clipboard.raw, clipboard.style);
      return;
    }

    if (e.key === 'ArrowUp') {
      e.preventDefault();
      const nr = Math.max(0, selected.r - 1);
      setSelected({ r: nr, c: selected.c });
      setSelEnd({ r: nr, c: selected.c });
    } else if (e.key === 'ArrowDown') {
      e.preventDefault();
      const nr = Math.min(ROWS - 1, selected.r + 1);
      setSelected({ r: nr, c: selected.c });
      setSelEnd({ r: nr, c: selected.c });
    } else if (e.key === 'ArrowLeft') {
      e.preventDefault();
      const nc = Math.max(0, selected.c - 1);
      setSelected({ r: selected.r, c: nc });
      setSelEnd({ r: selected.r, c: nc });
    } else if (e.key === 'ArrowRight' || e.key === 'Tab') {
      e.preventDefault();
      const nc = Math.min(COLS - 1, selected.c + 1);
      setSelected({ r: selected.r, c: nc });
      setSelEnd({ r: selected.r, c: nc });
    } else if (e.key === 'Delete' || e.key === 'Backspace') {
      const { r1, r2, c1, c2 } = getSelectedRange();
      pushHistory(sheets);
      setSheets((prev) =>
        prev.map((sh, idx) => {
          if (idx !== activeSheetIdx) return sh;
          const next = { ...sh.cells };
          for (let r = r1; r <= r2; r++) {
            for (let c = c1; c <= c2; c++) {
              delete next[cellId(r, c)];
            }
          }
          return { ...sh, cells: next };
        })
      );
    } else if (e.key === 'Enter' || e.key === 'F2') {
      e.preventDefault();
      setEditingCell(activeCellId);
      setEditValue(activeCellObj.raw || '');
    } else if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
      setEditingCell(activeCellId);
      setEditValue(e.key);
    }
  };

  // CSV Export & Import
  const exportCSV = () => {
    let csv = '';
    for (let r = 0; r < 50; r++) {
      const rowArr = [];
      let hasContent = false;
      for (let c = 0; c < COLS; c++) {
        const val = String(getDisplayValue(cellId(r, c)) ?? '');
        if (val !== '') hasContent = true;
        rowArr.push(`"${val.replace(/"/g, '""')}"`);
      }
      if (hasContent || r < 10) csv += rowArr.join(',') + '\n';
    }
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${currentSheet.name}.csv`;
    a.click();
  };

  const importCSV = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const text = evt.target.result;
      const lines = text.split(/\r?\n/);
      const newCells = {};
      lines.forEach((line, r) => {
        if (r >= ROWS || !line.trim()) return;
        const cols = line.split(',');
        cols.forEach((val, c) => {
          if (c >= COLS) return;
          const clean = val.replace(/^"|"$/g, '').trim();
          if (clean) newCells[cellId(r, c)] = { raw: clean, style: {} };
        });
      });
      pushHistory(sheets);
      setSheets((prev) =>
        prev.map((sh, idx) => (idx === activeSheetIdx ? { ...sh, cells: newCells } : sh))
      );
    };
    reader.readAsText(file);
  };

  // Sort active column
  const sortActiveColumn = (asc = true) => {
    pushHistory(sheets);
    const col = selected.c;
    const rowsData = [];
    for (let r = 1; r < ROWS; r++) {
      const val = getDisplayValue(cellId(r, col));
      if (val !== '') {
        const rowCells = {};
        for (let c = 0; c < COLS; c++) {
          if (currentSheet.cells[cellId(r, c)]) {
            rowCells[c] = currentSheet.cells[cellId(r, c)];
          }
        }
        rowsData.push({ sortKey: val, rowCells });
      }
    }
    rowsData.sort((a, b) => {
      const na = parseFloat(a.sortKey);
      const nb = parseFloat(b.sortKey);
      if (!isNaN(na) && !isNaN(nb)) return asc ? na - nb : nb - na;
      return asc
        ? String(a.sortKey).localeCompare(String(b.sortKey))
        : String(b.sortKey).localeCompare(String(a.sortKey));
    });

    setSheets((prev) =>
      prev.map((sh, idx) => {
        if (idx !== activeSheetIdx) return sh;
        const nextCells = { ...sh.cells };
        // Clear old rows 1..rowsData.length
        for (let r = 1; r <= rowsData.length; r++) {
          for (let c = 0; c < COLS; c++) delete nextCells[cellId(r, c)];
        }
        rowsData.forEach((item, i) => {
          Object.entries(item.rowCells).forEach(([c, cellObj]) => {
            nextCells[cellId(i + 1, Number(c))] = cellObj;
          });
        });
        return { ...sh, cells: nextCells };
      })
    );
  };

  // Chart Data from Selection
  const chartData = useMemo(() => {
    const { r1, r2, c1, c2 } = getSelectedRange();
    const items = [];
    for (let r = r1; r <= r2; r++) {
      const label = String(getDisplayValue(cellId(r, c1)) || `Row ${r + 1}`);
      const valCol = c2 > c1 ? c2 : c1;
      const val = parseFloat(evalFormula(currentSheet.cells[cellId(r, valCol)]?.raw ?? '', currentSheet)) || 0;
      items.push({ label, val });
    }
    return items.slice(0, 15);
  }, [selected, selEnd, currentSheet, showChartModal]);

  const { r1, r2, c1, c2 } = getSelectedRange();

  return (
    <div
      tabIndex={0}
      onKeyDown={handleKeyDown}
      onMouseUp={() => setIsSelecting(false)}
      className={`flex flex-col h-screen w-screen outline-none ${darkMode ? 'dark bg-zinc-900 text-zinc-100' : 'bg-white text-zinc-900'}`}
    >
      {/* Excel 365 Top Title Bar */}
      <div className="flex items-center justify-between px-3 py-1.5 bg-[#107C41] text-white text-xs select-none">
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-1.5 font-semibold text-sm">
            <FileSpreadsheet className="w-4 h-4" />
            <span>Excel 365</span>
          </div>
          <div className="flex items-center gap-1 border-l border-emerald-500/40 pl-2">
            <button onClick={() => localStorage.setItem('excel365_offline_wb', JSON.stringify(sheets))} title="Save Locally" className="p-1 hover:bg-emerald-700 rounded">
              <Save className="w-3.5 h-3.5" />
            </button>
            <button onClick={undo} title="Undo (Ctrl+Z)" className="p-1 hover:bg-emerald-700 rounded">
              <Undo2 className="w-3.5 h-3.5" />
            </button>
            <button onClick={redo} title="Redo (Ctrl+Y)" className="p-1 hover:bg-emerald-700 rounded">
              <Redo2 className="w-3.5 h-3.5" />
            </button>
          </div>
        </div>

        <input
          value={workbookName}
          onChange={(e) => setWorkbookName(e.target.value)}
          className="bg-emerald-800/60 text-white text-center px-3 py-0.5 rounded text-xs font-medium focus:outline-none focus:ring-1 focus:ring-white w-64"
        />

        <div className="flex items-center gap-2">
          <div className="relative flex items-center">
            <Search className="w-3.5 h-3.5 absolute left-2 text-emerald-200" />
            <input
              placeholder="Search in sheet..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="bg-emerald-800/70 text-white placeholder-emerald-200 pl-7 pr-2 py-0.5 rounded text-xs focus:outline-none w-40"
            />
          </div>
          <button
            onClick={() => setDarkMode(!darkMode)}
            className="flex items-center gap-1 px-2 py-0.5 bg-emerald-800 hover:bg-emerald-700 rounded text-xs"
          >
            {darkMode ? <Sun className="w-3.5 h-3.5" /> : <Moon className="w-3.5 h-3.5" />}
            <span>{darkMode ? 'Light' : 'Dark'}</span>
          </button>
        </div>
      </div>

      {/* Ribbon Tabs */}
      <div className="flex items-center gap-1 px-3 pt-1 bg-[#107C41] text-white text-xs">
        {['Home', 'Insert', 'Formulas', 'Data', 'View'].map((tab) => (
          <button
            key={tab}
            onClick={() => setActiveTab(tab)}
            className={`px-3 py-1 rounded-t font-medium transition ${
              activeTab === tab
                ? darkMode
                  ? 'bg-zinc-800 text-emerald-400'
                  : 'bg-zinc-100 text-[#107C41]'
                : 'hover:bg-emerald-700/80 text-white'
            }`}
          >
            {tab}
          </button>
        ))}
      </div>

      {/* Ribbon Toolbar */}
      <div className={`flex items-center gap-3 px-3 py-2 border-b text-xs ${darkMode ? 'bg-zinc-800 border-zinc-700' : 'bg-zinc-100 border-zinc-300'}`}>
        {activeTab === 'Home' && (
          <>
            <div className="flex items-center gap-1 pr-3 border-r border-zinc-300 dark:border-zinc-700">
              <button onClick={() => setClipboard({ ...activeCellObj })} className="flex flex-col items-center p-1.5 hover:bg-zinc-200 dark:hover:bg-zinc-700 rounded">
                <Copy className="w-4 h-4 text-[#107C41]" />
                <span className="text-[10px]">Copy</span>
              </button>
              <button onClick={() => clipboard && updateCell(activeCellId, clipboard.raw, clipboard.style)} className="flex flex-col items-center p-1.5 hover:bg-zinc-200 dark:hover:bg-zinc-700 rounded">
                <Clipboard className="w-4 h-4 text-[#107C41]" />
                <span className="text-[10px]">Paste</span>
              </button>
            </div>

            <div className="flex items-center gap-1 pr-3 border-r border-zinc-300 dark:border-zinc-700">
              <button
                onClick={() => applyStyleToSelection({ bold: !activeCellObj.style?.bold })}
                className={`p-1.5 rounded ${activeCellObj.style?.bold ? 'bg-emerald-200 dark:bg-emerald-900 text-[#107C41]' : 'hover:bg-zinc-200 dark:hover:bg-zinc-700'}`}
              >
                <Bold className="w-4 h-4" />
              </button>
              <button
                onClick={() => applyStyleToSelection({ italic: !activeCellObj.style?.italic })}
                className={`p-1.5 rounded ${activeCellObj.style?.italic ? 'bg-emerald-200 dark:bg-emerald-900 text-[#107C41]' : 'hover:bg-zinc-200 dark:hover:bg-zinc-700'}`}
              >
                <Italic className="w-4 h-4" />
              </button>
              <button
                onClick={() => applyStyleToSelection({ underline: !activeCellObj.style?.underline })}
                className={`p-1.5 rounded ${activeCellObj.style?.underline ? 'bg-emerald-200 dark:bg-emerald-900 text-[#107C41]' : 'hover:bg-zinc-200 dark:hover:bg-zinc-700'}`}
              >
                <Underline className="w-4 h-4" />
              </button>
              <input
                type="color"
                title="Cell Fill Color"
                onChange={(e) => applyStyleToSelection({ bg: e.target.value })}
                className="w-6 h-6 p-0 border-0 bg-transparent cursor-pointer"
              />
              <input
                type="color"
                title="Font Color"
                onChange={(e) => applyStyleToSelection({ color: e.target.value })}
                className="w-6 h-6 p-0 border-0 bg-transparent cursor-pointer"
              />
            </div>

            <div className="flex items-center gap-1 pr-3 border-r border-zinc-300 dark:border-zinc-700">
              <button onClick={() => applyStyleToSelection({ align: 'left' })} className="p-1.5 hover:bg-zinc-200 dark:hover:bg-zinc-700 rounded">
                <AlignLeft className="w-4 h-4" />
              </button>
              <button onClick={() => applyStyleToSelection({ align: 'center' })} className="p-1.5 hover:bg-zinc-200 dark:hover:bg-zinc-700 rounded">
                <AlignCenter className="w-4 h-4" />
              </button>
              <button onClick={() => applyStyleToSelection({ align: 'right' })} className="p-1.5 hover:bg-zinc-200 dark:hover:bg-zinc-700 rounded">
                <AlignRight className="w-4 h-4" />
              </button>
            </div>

            <div className="flex items-center gap-1 pr-3 border-r border-zinc-300 dark:border-zinc-700">
              <button onClick={() => applyStyleToSelection({ format: 'currency' })} title="INR Currency Format" className="p-1.5 hover:bg-zinc-200 dark:hover:bg-zinc-700 rounded flex items-center gap-1">
                <DollarSign className="w-4 h-4 text-[#107C41]" />
                <span>INR</span>
              </button>
              <button onClick={() => applyStyleToSelection({ format: 'percent' })} title="Percentage Format" className="p-1.5 hover:bg-zinc-200 dark:hover:bg-zinc-700 rounded">
                <Percent className="w-4 h-4" />
              </button>
              <button onClick={() => applyStyleToSelection({ format: 'general' })} title="General Format" className="p-1.5 hover:bg-zinc-200 dark:hover:bg-zinc-700 rounded">
                <Hash className="w-4 h-4" />
              </button>
            </div>

            <button
              onClick={() => updateCell(activeCellId, `=SUM(${colToLetter(selected.c)}1:${colToLetter(selected.c)}${selected.r})`)}
              className="flex items-center gap-1.5 px-2.5 py-1.5 bg-emerald-600 text-white rounded hover:bg-emerald-700 font-medium"
            >
              <FunctionSquare className="w-4 h-4" />
              <span>AutoSum</span>
            </button>
          </>
        )}

        {activeTab === 'Insert' && (
          <div className="flex items-center gap-2">
            <button onClick={() => setShowChartModal(true)} className="flex items-center gap-1.5 px-3 py-1.5 bg-[#107C41] text-white rounded hover:bg-emerald-700">
              <BarChart3 className="w-4 h-4" />
              <span>Insert Chart from Selection</span>
            </button>
            <button
              onClick={() => applyStyleToSelection({ bold: true, bg: '#107C41', color: '#ffffff' })}
              className="flex items-center gap-1.5 px-3 py-1.5 bg-zinc-200 dark:bg-zinc-700 rounded hover:opacity-90"
            >
              <Table className="w-4 h-4" />
              <span>Format as Table Header</span>
            </button>
          </div>
        )}

        {activeTab === 'Formulas' && (
          <div className="flex items-center gap-2 flex-wrap">
            {[
              { label: 'SUM', formula: '=SUM(B2:B4)' },
              { label: 'AVERAGE', formula: '=AVERAGE(B2:B4)' },
              { label: 'COUNT', formula: '=COUNT(B2:B4)' },
              { label: 'MAX', formula: '=MAX(B2:B4)' },
              { label: 'MIN', formula: '=MIN(B2:B4)' },
              { label: 'IF', formula: '=IF(C2>=B2, "Pass", "Check")' },
              { label: 'ROUND', formula: '=ROUND(C2/B2, 2)' },
              { label: 'TODAY', formula: '=TODAY()' }
            ].map((f) => (
              <button
                key={f.label}
                onClick={() => updateCell(activeCellId, f.formula)}
                className="px-2.5 py-1 bg-white dark:bg-zinc-700 border border-zinc-300 dark:border-zinc-600 rounded hover:border-[#107C41] font-mono text-xs"
              >
                {f.label}
              </button>
            ))}
          </div>
        )}

        {activeTab === 'Data' && (
          <div className="flex items-center gap-2">
            <button onClick={() => sortActiveColumn(true)} className="flex items-center gap-1 px-2.5 py-1.5 bg-white dark:bg-zinc-700 border border-zinc-300 dark:border-zinc-600 rounded">
              <ArrowUpDown className="w-3.5 h-3.5 text-[#107C41]" />
              <span>Sort A → Z</span>
            </button>
            <button onClick={() => sortActiveColumn(false)} className="flex items-center gap-1 px-2.5 py-1.5 bg-white dark:bg-zinc-700 border border-zinc-300 dark:border-zinc-600 rounded">
              <ArrowUpDown className="w-3.5 h-3.5 text-[#107C41]" />
              <span>Sort Z → A</span>
            </button>
            <button onClick={exportCSV} className="flex items-center gap-1 px-2.5 py-1.5 bg-[#107C41] text-white rounded">
              <Download className="w-3.5 h-3.5" />
              <span>Export CSV</span>
            </button>
            <button onClick={() => fileInputRef.current?.click()} className="flex items-center gap-1 px-2.5 py-1.5 bg-zinc-200 dark:bg-zinc-700 rounded">
              <Upload className="w-3.5 h-3.5" />
              <span>Import CSV</span>
            </button>
            <input ref={fileInputRef} type="file" accept=".csv" onChange={importCSV} className="hidden" />
          </div>
        )}

        {activeTab === 'View' && (
          <div className="flex items-center gap-2">
            <button
              onClick={() =>
                setSheets((prev) =>
                  prev.map((sh, i) => (i === activeSheetIdx ? { ...sh, freezeRow: !sh.freezeRow } : sh))
                )
              }
              className="flex items-center gap-1.5 px-3 py-1.5 bg-white dark:bg-zinc-700 border border-zinc-300 dark:border-zinc-600 rounded"
            >
              <Grid className="w-4 h-4 text-[#107C41]" />
              <span>{currentSheet.freezeRow ? 'Unfreeze Top Row' : 'Freeze Top Row'}</span>
            </button>
            <button
              onClick={() => {
                localStorage.removeItem('excel365_offline_wb');
                setSheets([defaultSampleSheet(), createEmptySheet('Sheet2')]);
              }}
              className="px-3 py-1.5 bg-red-600 text-white rounded"
            >
              Reset Workbook
            </button>
          </div>
        )}
      </div>

      {/* Formula Bar */}
      <div className={`flex items-center border-b text-xs ${darkMode ? 'bg-zinc-900 border-zinc-700' : 'bg-white border-zinc-300'}`}>
        <div className="w-16 text-center font-semibold py-1.5 border-r border-zinc-300 dark:border-zinc-700 text-[#107C41]">
          {activeCellId}
        </div>
        <div className="px-2.5 py-1.5 font-serif italic font-bold text-zinc-400 border-r border-zinc-300 dark:border-zinc-700">
          fx
        </div>
        <input
          value={editingCell === activeCellId ? editValue : activeCellObj.raw || ''}
          onChange={(e) => {
            if (editingCell !== activeCellId) setEditingCell(activeCellId);
            setEditValue(e.target.value);
          }}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              updateCell(activeCellId, editValue);
              setEditingCell(null);
            }
          }}
          placeholder="Enter value or formula (e.g., =SUM(B2:B4))"
          className="flex-1 px-3 py-1.5 bg-transparent focus:outline-none font-mono"
        />
      </div>

      {/* Main Spreadsheet Grid */}
      <div className="flex-1 overflow-auto relative">
        <table className="border-collapse table-fixed text-xs">
          <thead className="sticky top-0 z-20">
            <tr>
              <th className={`w-10 border ${darkMode ? 'bg-zinc-800 border-zinc-700' : 'bg-zinc-100 border-zinc-300'}`} />
              {Array.from({ length: COLS }).map((_, c) => (
                <th
                  key={c}
                  style={{ width: currentSheet.colWidths[c] || 100 }}
                  className={`py-1 font-semibold border select-none ${
                    c >= c1 && c <= c2
                      ? 'bg-emerald-100 dark:bg-emerald-950 text-[#107C41] border-b-2 border-b-[#107C41]'
                      : darkMode
                      ? 'bg-zinc-800 border-zinc-700 text-zinc-300'
                      : 'bg-zinc-100 border-zinc-300 text-zinc-700'
                  }`}
                >
                  {colToLetter(c)}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {Array.from({ length: ROWS }).map((_, r) => (
              <tr
                key={r}
                className={r === 0 && currentSheet.freezeRow ? 'sticky top-6 z-10 shadow-sm' : ''}
              >
                <td
                  className={`sticky left-0 z-10 text-center font-medium border select-none ${
                    r >= r1 && r <= r2
                      ? 'bg-emerald-100 dark:bg-emerald-950 text-[#107C41] border-r-2 border-r-[#107C41]'
                      : darkMode
                      ? 'bg-zinc-800 border-zinc-700 text-zinc-400'
                      : 'bg-zinc-100 border-zinc-300 text-zinc-600'
                  }`}
                >
                  {r + 1}
                </td>
                {Array.from({ length: COLS }).map((__, c) => {
                  const id = cellId(r, c);
                  const cell = currentSheet.cells[id];
                  const isActive = selected.r === r && selected.c === c;
                  const isInRange = r >= r1 && r <= r2 && c >= c1 && c <= c2;
                  const displayVal = getDisplayValue(id);
                  const isMatch =
                    searchQuery &&
                    String(displayVal).toLowerCase().includes(searchQuery.toLowerCase());

                  return (
                    <td
                      key={c}
                      onMouseDown={() => {
                        if (editingCell && editingCell !== id) {
                          updateCell(editingCell, editValue);
                          setEditingCell(null);
                        }
                        setSelected({ r, c });
                        setSelEnd({ r, c });
                        setIsSelecting(true);
                      }}
                      onMouseEnter={() => {
                        if (isSelecting) setSelEnd({ r, c });
                      }}
                      onDoubleClick={() => {
                        setEditingCell(id);
                        setEditValue(cell?.raw || '');
                      }}
                      style={{
                        backgroundColor: isMatch
                          ? '#fef08a'
                          : cell?.style?.bg || (isInRange && !isActive ? (darkMode ? '#064e3b44' : '#ecfdf5') : undefined),
                        color: cell?.style?.color || undefined,
                        fontWeight: cell?.style?.bold ? 'bold' : 'normal',
                        fontStyle: cell?.style?.italic ? 'italic' : 'normal',
                        textDecoration: cell?.style?.underline ? 'underline' : 'none',
                        textAlign: cell?.style?.align || (typeof displayVal === 'number' ? 'right' : 'left')
                      }}
                      className={`px-1.5 py-1 border truncate cursor-cell ${
                        isActive
                          ? 'ring-2 ring-[#107C41] z-10 relative'
                          : darkMode
                          ? 'border-zinc-800'
                          : 'border-zinc-200'
                      }`}
                    >
                      {editingCell === id ? (
                        <input
                          ref={inputRef}
                          autoFocus
                          value={editValue}
                          onChange={(e) => setEditValue(e.target.value)}
                          className="w-full h-full bg-transparent outline-none font-mono"
                        />
                      ) : (
                        displayVal
                      )}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* Bottom Sheet Tabs & Status Bar */}
      <div className={`flex items-center justify-between px-3 py-1 border-t text-xs select-none ${darkMode ? 'bg-zinc-800 border-zinc-700' : 'bg-zinc-100 border-zinc-300'}`}>
        <div className="flex items-center gap-1">
          {sheets.map((sh, idx) => (
            <div
              key={sh.name + idx}
              onClick={() => setActiveSheetIdx(idx)}
              className={`flex items-center gap-1.5 px-3 py-1 rounded-t cursor-pointer border-b-2 ${
                idx === activeSheetIdx
                  ? 'bg-white dark:bg-zinc-900 text-[#107C41] font-semibold border-[#107C41]'
                  : 'border-transparent hover:bg-zinc-200 dark:hover:bg-zinc-700'
              }`}
            >
              <span>{sh.name}</span>
              {sheets.length > 1 && (
                <X
                  onClick={(e) => {
                    e.stopPropagation();
                    setSheets((prev) => prev.filter((_, i) => i !== idx));
                    setActiveSheetIdx(0);
                  }}
                  className="w-3 h-3 opacity-60 hover:opacity-100"
                />
              )}
            </div>
          ))}
          <button
            onClick={() => {
              const name = `Sheet${sheets.length + 1}`;
              setSheets((prev) => [...prev, createEmptySheet(name)]);
              setActiveSheetIdx(sheets.length);
            }}
            title="Add New Sheet"
            className="p-1 hover:bg-zinc-200 dark:hover:bg-zinc-700 rounded-full ml-1"
          >
            <Plus className="w-3.5 h-3.5 text-[#107C41]" />
          </button>
        </div>

        <div className="flex items-center gap-4 text-[11px] font-medium text-zinc-600 dark:text-zinc-300">
          <span>Ready (Offline Mode)</span>
          {selectionStats.count > 1 && (
            <>
              <span>Count: {selectionStats.count}</span>
              {selectionStats.hasNums && (
                <>
                  <span>Average: {selectionStats.avg.toFixed(2)}</span>
                  <span>Sum: {selectionStats.sum.toLocaleString('en-IN')}</span>
                </>
              )}
            </>
          )}
        </div>
      </div>

      {/* Chart Modal */}
      {showChartModal && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50">
          <div className={`w-[520px] p-5 rounded-lg shadow-xl ${darkMode ? 'bg-zinc-800 text-white' : 'bg-white text-zinc-900'}`}>
            <div className="flex justify-between items-center mb-4">
              <h3 className="font-bold text-sm text-[#107C41]">Excel 365 Chart Visualizer</h3>
              <button onClick={() => setShowChartModal(false)}>
                <X className="w-4 h-4" />
              </button>
            </div>
            <div className="flex gap-2 mb-4 text-xs">
              {['bar', 'horizontal'].map((t) => (
                <button
                  key={t}
                  onClick={() => setChartType(t)}
                  className={`px-3 py-1 rounded capitalize ${chartType === t ? 'bg-[#107C41] text-white' : 'bg-zinc-200 dark:bg-zinc-700'}`}
                >
                  {t} Chart
                </button>
              ))}
            </div>
            <div className="space-y-2 max-h-64 overflow-y-auto p-3 border rounded border-zinc-200 dark:border-zinc-700">
              {chartData.map((d, i) => {
                const maxVal = Math.max(...chartData.map((x) => x.val), 1);
                const pct = Math.max(4, Math.round((d.val / maxVal) * 100));
                return (
                  <div key={i} className="flex items-center gap-2 text-xs">
                    <span className="w-24 truncate font-medium">{d.label}</span>
                    <div className="flex-1 bg-zinc-200 dark:bg-zinc-700 h-5 rounded overflow-hidden">
                      <div style={{ width: `${pct}%` }} className="bg-[#107C41] h-full transition-all" />
                    </div>
                    <span className="w-16 text-right font-mono">{d.val}</span>
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
