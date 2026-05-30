import { useState, useRef, useCallback } from 'react';
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

export default function Transactions({ transactions, settings, onChange }: Props) {
  const now = new Date();
  const [year, setYear] = useState(now.getFullYear());
  const [month, setMonth] = useState(now.getMonth() + 1);
  const fileRef = useRef<HTMLInputElement>(null);

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
            <tr className="bg-gray-50 border-b border-gray-200">
              <th className="w-28 px-2 py-2 text-left text-xs font-semibold text-gray-600 border-r border-gray-200">Date</th>
              <th className="w-24 px-2 py-2 text-left text-xs font-semibold text-gray-600 border-r border-gray-200">Start Bal</th>
              <th className="w-24 px-2 py-2 text-left text-xs font-semibold text-gray-600 border-r border-gray-200">End Bal <span className="font-normal text-gray-400">(calc)</span></th>
              <th className="w-24 px-2 py-2 text-left text-xs font-semibold text-gray-600 border-r border-gray-200">Amount</th>
              <th className="w-32 px-2 py-2 text-left text-xs font-semibold text-gray-600 border-r border-gray-200">Type</th>
              <th className="w-36 px-2 py-2 text-left text-xs font-semibold text-gray-600 border-r border-gray-200">Category</th>
              <th className="w-36 px-2 py-2 text-left text-xs font-semibold text-gray-600 border-r border-gray-200">Sub Category</th>
              <th className="w-36 px-2 py-2 text-left text-xs font-semibold text-gray-600 border-r border-gray-200">Paid To</th>
              <th className="w-36 px-2 py-2 text-left text-xs font-semibold text-gray-600 border-r border-gray-200">Comment</th>
              <th className="w-64 px-2 py-2 text-left text-xs font-semibold text-gray-600 border-r border-gray-200">Bank Text</th>
              <th className="w-24 px-2 py-2 text-left text-xs font-semibold text-gray-600">Budgeted</th>
              <th className="w-8"></th>
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
                <td className="w-28 px-1 py-0.5 border-r border-gray-100">
                  <input
                    className="table-cell-input"
                    type="date"
                    value={t.date}
                    onChange={e => updateField(t.id, 'date', e.target.value)}
                  />
                </td>

                {/* Start Balance — editable only on first row */}
                <td className="w-24 px-1 py-0.5 border-r border-gray-100">
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
                <td className="w-24 px-1 py-0.5 border-r border-gray-100">
                  <span className={`block text-right px-1 py-0.5 text-xs bg-gray-50 rounded ${
                    t.endBalance !== null && t.endBalance < 0 ? 'text-red-500' : 'text-gray-500'
                  }`}>
                    {fmt(t.endBalance)}
                  </span>
                </td>

                {/* Amount */}
                <td className="w-24 px-1 py-0.5 border-r border-gray-100">
                  <input
                    className={`table-cell-input text-right ${
                      t.amount !== null && t.amount < 0 ? 'text-red-600' : 'text-green-700'
                    }`}
                    value={fmt(t.amount)}
                    onChange={e => updateField(t.id, 'amount', parseNum(e.target.value))}
                  />
                </td>

                {/* Type */}
                <td className="w-32 px-1 py-0.5 border-r border-gray-100">
                  <select className="table-cell-select" value={t.type}
                    onChange={e => updateField(t.id, 'type', e.target.value)}>
                    <option value=""></option>
                    {getOptions('select-type').map(o => <option key={o} value={o}>{o}</option>)}
                  </select>
                </td>

                {/* Category */}
                <td className="w-36 px-1 py-0.5 border-r border-gray-100">
                  <select className="table-cell-select" value={t.category}
                    onChange={e => updateField(t.id, 'category', e.target.value)}>
                    <option value=""></option>
                    {getOptions('select-cat').map(o => <option key={o} value={o}>{o}</option>)}
                  </select>
                </td>

                {/* Sub Category */}
                <td className="w-36 px-1 py-0.5 border-r border-gray-100">
                  <select className="table-cell-select" value={t.subCategory}
                    onChange={e => updateField(t.id, 'subCategory', e.target.value)}>
                    <option value=""></option>
                    {getOptions('select-sub').map(o => <option key={o} value={o}>{o}</option>)}
                  </select>
                </td>

                {/* Paid To */}
                <td className="w-36 px-1 py-0.5 border-r border-gray-100">
                  <input className="table-cell-input" value={t.paidTo}
                    onChange={e => updateField(t.id, 'paidTo', e.target.value)} />
                </td>

                {/* Comment */}
                <td className="w-36 px-1 py-0.5 border-r border-gray-100">
                  <input className="table-cell-input" value={t.comment}
                    onChange={e => updateField(t.id, 'comment', e.target.value)} />
                </td>

                {/* Bank Text */}
                <td className="w-64 px-1 py-0.5 border-r border-gray-100">
                  <input className="table-cell-input" value={t.bankText}
                    title={t.bankText}
                    onChange={e => updateField(t.id, 'bankText', e.target.value)} />
                </td>

                {/* Budgeted */}
                <td className="w-24 px-1 py-0.5 border-r border-gray-100">
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
