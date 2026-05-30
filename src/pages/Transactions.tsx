import { useState, useRef, useCallback, useEffect } from 'react';
import * as XLSX from 'xlsx';
import type { Transaction, DropdownSettings } from '../types';
import { generateId } from '../store';

interface Props {
  transactions: Transaction[];
  settings: DropdownSettings;
  onChange: (transactions: Transaction[]) => void;
}

const MONTHS = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];

function r2(n: number): number {
  return Math.round(n * 100) / 100;
}

function fmt(n: number | null | undefined): string {
  if (n === null || n === undefined) return '';
  return n.toFixed(2);
}

function parseNum(s: string): number | null {
  const v = parseFloat(s.replace(',', '.'));
  return isNaN(v) ? null : v;
}

// Always display and accept DD/MM/YYYY regardless of OS locale
function isoToDisplay(iso: string): string {
  if (!iso || iso.length < 10) return iso;
  const [y, m, d] = iso.split('-');
  return `${d}/${m}/${y}`;
}

function displayToIso(display: string): string {
  const parts = display.split('/');
  if (parts.length === 3) {
    const [d, m, y] = parts;
    if (d.length <= 2 && m.length <= 2 && y.length === 4) {
      return `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
    }
  }
  return '';
}

function DateInput({ value, onChange }: { value: string; onChange: (v: string) => void }) {
  const [raw, setRaw] = useState(() => isoToDisplay(value));

  useEffect(() => { setRaw(isoToDisplay(value)); }, [value]);

  function handleChange(v: string) {
    const digits = v.replace(/\D/g, '').slice(0, 8);
    let formatted = digits;
    if (digits.length > 4) formatted = `${digits.slice(0,2)}/${digits.slice(2,4)}/${digits.slice(4)}`;
    else if (digits.length > 2) formatted = `${digits.slice(0,2)}/${digits.slice(2)}`;
    setRaw(formatted);
  }

  function commit() {
    const iso = displayToIso(raw);
    if (iso) onChange(iso);
    else setRaw(isoToDisplay(value));
  }

  return (
    <input
      className="table-cell-input"
      value={raw}
      placeholder="DD/MM/YYYY"
      onChange={e => handleChange(e.target.value)}
      onBlur={commit}
      onKeyDown={e => { if (e.key === 'Enter' || e.key === 'Tab') commit(); }}
    />
  );
}

// Recalculate End Balance = Start + Amount for every row in the month,
// cascading Start Balance from the previous row's End Balance (except row 0).
function recalcInMonth(all: Transaction[], year: number, month: number): Transaction[] {
  const inMonth = all
    .filter(t => {
      if (!t.date) return false;
      const d = new Date(t.date);
      return d.getFullYear() === year && d.getMonth() + 1 === month;
    })
    .sort((a, b) => a.date.localeCompare(b.date));

  const recalced = inMonth.reduce<Transaction[]>((acc, t, i) => {
    if (i === 0) {
      const end = t.startBalance !== null && t.amount !== null
        ? r2(t.startBalance + t.amount)
        : null;
      return [...acc, { ...t, endBalance: end }];
    }
    const prev = acc[i - 1];
    const start = prev.endBalance;
    const end = start !== null && t.amount !== null ? r2(start + t.amount) : null;
    return [...acc, { ...t, startBalance: start, endBalance: end }];
  }, []);

  const map = new Map(recalced.map(t => [t.id, t]));
  return all.map(t => map.get(t.id) ?? t);
}

const DEFAULT_WIDTHS: Record<string, number> = {
  date: 120, startBalance: 90, endBalance: 90, amount: 90,
  type: 125, category: 140, subCategory: 140,
  paidTo: 140, comment: 140, bankText: 260, budgeted: 90,
};

export default function Transactions({ transactions, settings, onChange }: Props) {
  const now = new Date();
  const [year, setYear] = useState(now.getFullYear());
  const [month, setMonth] = useState(now.getMonth() + 1);
  const fileRef = useRef<HTMLInputElement>(null);
  const [colWidths, setColWidths] = useState<Record<string, number>>(DEFAULT_WIDTHS);
  const focusRowId = useRef<string | null>(null);

  function onResizeStart(key: string, e: React.MouseEvent) {
    e.preventDefault();
    const startX = e.clientX;
    const startW = colWidths[key];
    function onMove(ev: MouseEvent) {
      setColWidths(w => ({ ...w, [key]: Math.max(50, startW + ev.clientX - startX) }));
    }
    function onUp() {
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
    }
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  }

  const years = Array.from(
    new Set([now.getFullYear() - 1, now.getFullYear(), now.getFullYear() + 1,
      ...transactions.map(t => t.date ? new Date(t.date).getFullYear() : now.getFullYear())])
  ).sort();

  const filtered = transactions
    .filter(t => {
      if (!t.date) return false;
      const d = new Date(t.date);
      return d.getFullYear() === year && d.getMonth() + 1 === month;
    })
    .sort((a, b) => a.date.localeCompare(b.date));

  function updateField(id: string, field: keyof Transaction, value: string | number | null) {
    let updated = transactions.map(t =>
      t.id !== id ? t : { ...t, [field]: value } as Transaction
    );
    // Recalculate balance chain whenever amount or first-row startBalance changes
    if (field === 'amount' || field === 'startBalance' || field === 'date') {
      updated = recalcInMonth(updated, year, month);
    }
    onChange(updated);
  }

  function addRow() {
    const lastInMonth = filtered[filtered.length - 1];
    const dateStr = lastInMonth?.date ?? `${year}-${String(month).padStart(2, '0')}-01`;
    const newRow: Transaction = {
      id: generateId(),
      date: dateStr,
      startBalance: lastInMonth?.endBalance ?? null,
      endBalance: null,
      amount: null,
      type: '',
      category: '',
      subCategory: '',
      paidTo: '',
      comment: '',
      bankText: '',
      budgeted: '',
    };
    focusRowId.current = newRow.id;
    const updated = recalcInMonth([...transactions, newRow], year, month);
    onChange(updated);
  }

  function deleteRow(id: string) {
    const updated = recalcInMonth(transactions.filter(t => t.id !== id), year, month);
    onChange(updated);
  }

  function handleImport(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const data = ev.target?.result;
      const wb = XLSX.read(data, { type: 'binary' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, { raw: true });

      const imported: Transaction[] = rows.map(row => {
        const rawDate = row['transactiondate'] as number;
        const ds = String(rawDate);
        const dateStr = `${ds.slice(0, 4)}-${ds.slice(4, 6)}-${ds.slice(6, 8)}`;
        return {
          id: generateId(),
          date: dateStr,
          startBalance: (row['startsaldo'] as number) ?? null,
          endBalance: null,
          amount: (row['amount'] as number) ?? null,
          type: '',
          category: '',
          subCategory: '',
          paidTo: '',
          comment: '',
          bankText: String(row['description'] ?? ''),
          budgeted: '',
        };
      });

      const kept = transactions.filter(t => {
        const d = new Date(t.date);
        const sameMonth = d.getFullYear() === year && d.getMonth() + 1 === month;
        return !sameMonth || t.bankText === '';
      });

      const filteredImport = imported.filter(t => {
        const d = new Date(t.date);
        return d.getFullYear() === year && d.getMonth() + 1 === month;
      });

      const merged = recalcInMonth([...kept, ...filteredImport], year, month);
      onChange(merged);
    };
    reader.readAsBinaryString(file);
    e.target.value = '';
  }

  const getOptions = useCallback((colType: string) => {
    if (colType === 'select-budgeted') return settings.budgetedOptions;
    if (colType === 'select-cat') return settings.categories;
    if (colType === 'select-sub') return settings.subCategories;
    if (colType === 'select-type') return settings.types;
    return [];
  }, [settings]);

  return (
    <div className="p-4">
      {/* Filters & Actions */}
      <div className="flex items-center gap-3 mb-4 flex-wrap">
        <select
          value={year}
          onChange={e => setYear(Number(e.target.value))}
          className="border border-gray-300 rounded px-3 py-1.5 text-sm bg-white"
        >
          {years.map(y => <option key={y} value={y}>{y}</option>)}
        </select>
        <select
          value={month}
          onChange={e => setMonth(Number(e.target.value))}
          className="border border-gray-300 rounded px-3 py-1.5 text-sm bg-white"
        >
          {MONTHS.map((m, i) => <option key={i + 1} value={i + 1}>{m}</option>)}
        </select>
        <button
          onClick={addRow}
          className="bg-blue-600 text-white px-3 py-1.5 rounded text-sm hover:bg-blue-700"
        >
          + Add Row
        </button>
        <button
          onClick={() => fileRef.current?.click()}
          className="bg-green-600 text-white px-3 py-1.5 rounded text-sm hover:bg-green-700"
        >
          Import Bank File
        </button>
        <input ref={fileRef} type="file" accept=".xls,.xlsx" className="hidden" onChange={handleImport} />
        <span className="text-sm text-gray-500 ml-auto">
          {filtered.length} transactions
          {filtered.length > 0 && (() => {
            const total = filtered.reduce((s, t) => s + (t.amount ?? 0), 0);
            return ` | Net: €${total.toFixed(2)}`;
          })()}
        </span>
      </div>

      {/* Table */}
      <div className="overflow-x-auto rounded-lg border border-gray-200 bg-white">
        <table className="text-sm border-collapse min-w-full">
          <thead>
            <tr className="bg-gray-50 border-b border-gray-200 select-none">
              {([
                ['date',         'Date'],
                ['startBalance', 'Start Bal'],
                ['endBalance',   'End Bal'],
                ['amount',       'Amount'],
                ['type',         'Type'],
                ['category',     'Category'],
                ['subCategory',  'Sub Category'],
                ['paidTo',       'Paid To'],
                ['comment',      'Comment'],
                ['bankText',     'Bank Text'],
                ['budgeted',     'Budgeted'],
              ] as [string, string][]).map(([key, label]) => (
                <th
                  key={key}
                  style={{ width: colWidths[key], minWidth: colWidths[key] }}
                  className="relative px-2 py-2 text-left text-xs font-semibold text-gray-600 border-r border-gray-200 overflow-hidden"
                >
                  <span className="truncate block pr-2">{label}</span>
                  <div
                    className="absolute right-0 top-0 h-full w-1.5 cursor-col-resize hover:bg-blue-400 active:bg-blue-500"
                    onMouseDown={e => onResizeStart(key, e)}
                  />
                </th>
              ))}
              <th className="w-8 border-gray-200"></th>
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 && (
              <tr>
                <td colSpan={12} className="px-4 py-8 text-center text-gray-400">
                  No transactions for {MONTHS[month - 1]} {year}. Add a row or import a bank file.
                </td>
              </tr>
            )}
            {filtered.map((t, rowIdx) => (
              <tr
                key={t.id}
                className={`border-b border-gray-100 hover:bg-blue-50 transition-colors ${
                  rowIdx % 2 === 0 ? 'bg-white' : 'bg-gray-50/50'
                }`}
              >
                {/* Date */}
                <td style={{ width: colWidths.date, minWidth: colWidths.date }} className="px-1 py-0.5 border-r border-gray-100 overflow-hidden">
                  <DateInput value={t.date} onChange={v => updateField(t.id, 'date', v)} />
                </td>

                {/* Start Balance — editable only on first row */}
                <td style={{ width: colWidths.startBalance, minWidth: colWidths.startBalance }} className="px-1 py-0.5 border-r border-gray-100 overflow-hidden">
                  {rowIdx === 0 ? (
                    <input
                      className="table-cell-input text-right"
                      value={fmt(t.startBalance)}
                      onChange={e => updateField(t.id, 'startBalance', parseNum(e.target.value))}
                    />
                  ) : (
                    <span className="block text-right px-1 py-0.5 text-gray-500 text-xs bg-gray-50 rounded">
                      {fmt(t.startBalance)}
                    </span>
                  )}
                </td>

                {/* End Balance — always calculated */}
                <td style={{ width: colWidths.endBalance, minWidth: colWidths.endBalance }} className="px-1 py-0.5 border-r border-gray-100 overflow-hidden">
                  <span className={`block text-right px-1 py-0.5 text-xs bg-gray-50 rounded ${
                    t.endBalance !== null && t.endBalance < 0 ? 'text-red-500' : 'text-gray-500'
                  }`}>
                    {fmt(t.endBalance)}
                  </span>
                </td>

                {/* Amount */}
                <td style={{ width: colWidths.amount, minWidth: colWidths.amount }} className="px-1 py-0.5 border-r border-gray-100 overflow-hidden">
                  <input
                    className={`table-cell-input text-right ${
                      t.amount !== null && t.amount < 0 ? 'text-red-600' :
                      t.amount !== null && t.amount > 0 ? 'text-green-700' : ''
                    }`}
                    value={fmt(t.amount)}
                    onChange={e => updateField(t.id, 'amount', parseNum(e.target.value))}
                    ref={el => {
                      if (el && focusRowId.current === t.id) {
                        el.focus();
                        el.select();
                        focusRowId.current = null;
                      }
                    }}
                  />
                </td>

                {/* Type */}
                <td style={{ width: colWidths.type, minWidth: colWidths.type }} className="px-1 py-0.5 border-r border-gray-100 overflow-hidden">
                  <select className="table-cell-select" value={t.type}
                    onChange={e => updateField(t.id, 'type', e.target.value)}>
                    <option value=""></option>
                    {getOptions('select-type').map(o => <option key={o} value={o}>{o}</option>)}
                  </select>
                </td>

                {/* Category */}
                <td style={{ width: colWidths.category, minWidth: colWidths.category }} className="px-1 py-0.5 border-r border-gray-100 overflow-hidden">
                  <select className="table-cell-select" value={t.category}
                    onChange={e => updateField(t.id, 'category', e.target.value)}>
                    <option value=""></option>
                    {getOptions('select-cat').map(o => <option key={o} value={o}>{o}</option>)}
                  </select>
                </td>

                {/* Sub Category */}
                <td style={{ width: colWidths.subCategory, minWidth: colWidths.subCategory }} className="px-1 py-0.5 border-r border-gray-100 overflow-hidden">
                  <select className="table-cell-select" value={t.subCategory}
                    onChange={e => updateField(t.id, 'subCategory', e.target.value)}>
                    <option value=""></option>
                    {getOptions('select-sub').map(o => <option key={o} value={o}>{o}</option>)}
                  </select>
                </td>

                {/* Paid To */}
                <td style={{ width: colWidths.paidTo, minWidth: colWidths.paidTo }} className="px-1 py-0.5 border-r border-gray-100 overflow-hidden">
                  <input className="table-cell-input" value={t.paidTo}
                    onChange={e => updateField(t.id, 'paidTo', e.target.value)} />
                </td>

                {/* Comment */}
                <td style={{ width: colWidths.comment, minWidth: colWidths.comment }} className="px-1 py-0.5 border-r border-gray-100 overflow-hidden">
                  <input className="table-cell-input" value={t.comment}
                    onChange={e => updateField(t.id, 'comment', e.target.value)} />
                </td>

                {/* Bank Text */}
                <td style={{ width: colWidths.bankText, minWidth: colWidths.bankText }} className="px-1 py-0.5 border-r border-gray-100 overflow-hidden">
                  <input className="table-cell-input" value={t.bankText}
                    title={t.bankText}
                    onChange={e => updateField(t.id, 'bankText', e.target.value)} />
                </td>

                {/* Budgeted */}
                <td style={{ width: colWidths.budgeted, minWidth: colWidths.budgeted }} className="px-1 py-0.5 border-r border-gray-100 overflow-hidden">
                  <select className="table-cell-select" value={t.budgeted}
                    onChange={e => updateField(t.id, 'budgeted', e.target.value)}>
                    <option value=""></option>
                    {getOptions('select-budgeted').map(o => <option key={o} value={o}>{o}</option>)}
                  </select>
                </td>

                {/* Delete */}
                <td className="px-1 py-0.5">
                  <button onClick={() => deleteRow(t.id)}
                    className="text-gray-300 hover:text-red-500 text-xs px-1" title="Delete row">
                    ✕
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
