import { useState } from 'react';
import type { Transaction } from '../types';

interface Props {
  transactions: Transaction[];
}

const MONTHS = [
  'January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December',
];

type PivotRow = { label: string; children?: PivotRow[]; amount: number };

function groupBy<T>(arr: T[], key: (item: T) => string): Map<string, T[]> {
  const map = new Map<string, T[]>();
  for (const item of arr) {
    const k = key(item);
    if (!map.has(k)) map.set(k, []);
    map.get(k)!.push(item);
  }
  return map;
}

function buildPivot(
  rows: Transaction[],
  levels: ((t: Transaction) => string)[]
): PivotRow[] {
  if (levels.length === 0) return [];
  const [first, ...rest] = levels;
  const grouped = groupBy(rows, first);
  const result: PivotRow[] = [];
  for (const [label, items] of grouped) {
    const amount = items.reduce((s, t) => s + (t.amount ?? 0), 0);
    const children = rest.length > 0 ? buildPivot(items, rest) : undefined;
    result.push({ label, amount, children });
  }
  return result.sort((a, b) => a.label.localeCompare(b.label));
}

function PivotTable({ title, rows }: { title: string; rows: PivotRow[] }) {
  const total = rows.reduce((s, r) => s + r.amount, 0);

  function renderRows(items: PivotRow[], depth = 0): React.ReactNode {
    return items.map((item, i) => (
      <>
        <tr key={`${depth}-${i}`} className={depth === 0 ? 'bg-gray-50 font-medium' : ''}>
          <td
            className="py-1 px-2 text-sm border-b border-gray-100"
            style={{ paddingLeft: `${8 + depth * 16}px` }}
          >
            {item.label || '(blank)'}
          </td>
          <td className={`py-1 px-3 text-sm text-right border-b border-gray-100 ${item.amount < 0 ? 'text-red-600' : 'text-green-700'}`}>
            €{item.amount.toFixed(2)}
          </td>
        </tr>
        {item.children && renderRows(item.children, depth + 1)}
      </>
    ));
  }

  return (
    <div className="bg-white rounded-lg border border-gray-200 overflow-hidden">
      <div className="bg-blue-600 text-white px-4 py-2 text-sm font-semibold">{title}</div>
      <table className="w-full">
        <thead>
          <tr className="bg-gray-100">
            <th className="py-1.5 px-2 text-xs font-semibold text-gray-600 text-left">Category / Detail</th>
            <th className="py-1.5 px-3 text-xs font-semibold text-gray-600 text-right">Amount</th>
          </tr>
        </thead>
        <tbody>
          {rows.length === 0 ? (
            <tr><td colSpan={2} className="py-4 text-center text-gray-400 text-sm">No data</td></tr>
          ) : (
            renderRows(rows)
          )}
        </tbody>
        {rows.length > 0 && (
          <tfoot>
            <tr className="border-t-2 border-gray-300">
              <td className="py-1.5 px-2 text-sm font-bold">Grand Total</td>
              <td className={`py-1.5 px-3 text-sm font-bold text-right ${total < 0 ? 'text-red-600' : 'text-green-700'}`}>
                €{total.toFixed(2)}
              </td>
            </tr>
          </tfoot>
        )}
      </table>
    </div>
  );
}

export default function Analytics({ transactions }: Props) {
  const now = new Date();
  const [year, setYear] = useState(now.getFullYear());
  const [month, setMonth] = useState(now.getMonth() + 1);

  const years = Array.from(
    new Set([now.getFullYear() - 1, now.getFullYear(),
      ...transactions.map(t => t.date ? new Date(t.date).getFullYear() : now.getFullYear())])
  ).sort();

  const filtered = transactions.filter(t => {
    if (!t.date) return false;
    const d = new Date(t.date);
    return d.getFullYear() === year && d.getMonth() + 1 === month;
  });

  const oneOffBudgeted = filtered.filter(t => t.type === 'One off' && t.budgeted === 'Yes');
  const recurringBudgeted = filtered.filter(t => t.type === 'Reccuring' && t.budgeted === 'Yes');
  const notBudgeted = filtered.filter(t => t.budgeted === 'No');
  const moneyIn = filtered.filter(t => t.category === 'Money In');
  const writeOff = filtered.filter(t => t.budgeted === 'WO');

  return (
    <div className="p-4">
      <div className="flex items-center gap-3 mb-5">
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
        <span className="text-sm text-gray-500">{filtered.length} transactions in {MONTHS[month - 1]} {year}</span>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4">
        <PivotTable
          title="One-off Budgeted"
          rows={buildPivot(oneOffBudgeted, [t => t.category, t => t.subCategory])}
        />
        <PivotTable
          title="Recurring Budgeted"
          rows={buildPivot(recurringBudgeted, [t => t.category, t => t.subCategory])}
        />
        <PivotTable
          title="Not Budgeted"
          rows={buildPivot(notBudgeted, [t => t.category, t => t.paidTo])}
        />
        <PivotTable
          title="Money In"
          rows={buildPivot(moneyIn, [t => t.category, t => t.subCategory, t => t.paidTo])}
        />
        <PivotTable
          title="Write-off"
          rows={buildPivot(writeOff, [t => t.category, t => t.subCategory, t => t.paidTo])}
        />
      </div>
    </div>
  );
}
