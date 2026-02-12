
import React from 'react';
import { FileText, Download, Filter, ArrowUpRight, HandCoins, ReceiptText, Calculator } from 'lucide-react';
import { MOCK_INVOICES, MOCK_LOANS } from '../mockData';

export const GstReports: React.FC = () => {
  const summaryData = {
    period: 'May 2024',
    taxableValue: 1250000,
    cgst: 18750,
    sgst: 18750,
    totalGst: 37500,
    invoiceCount: 42
  };

  const loanStats = {
    totalPrincipal: MOCK_LOANS.reduce((acc, curr) => acc + curr.principalAmount, 0),
    avgInterest: MOCK_LOANS.reduce((acc, curr) => acc + curr.interestRate, 0) / MOCK_LOANS.length,
    activeWeight: MOCK_LOANS.reduce((acc, curr) => acc + curr.weight, 0)
  };

  return (
    <div className="space-y-8 animate-in fade-in duration-700">
      <div className="flex justify-between items-center">
        <div>
          <h1 className="text-3xl font-black text-slate-900">Vyapar Dashboard</h1>
          <p className="text-slate-500 italic">Unified Sales & Loan Ledger Reporting</p>
        </div>
        <div className="flex gap-3">
          <button className="flex items-center gap-2 px-5 py-3 bg-white border-2 border-slate-100 rounded-2xl hover:bg-slate-50 transition-all font-bold">
            <Filter size={18} /> Filter Period
          </button>
          <button className="flex items-center gap-2 px-6 py-3 bg-slate-900 text-white font-black rounded-2xl hover:bg-slate-800 transition-all shadow-xl">
            <Download size={18} /> Export Unified Excel
          </button>
        </div>
      </div>

      {/* Top Level Unified Stats */}
      <div className="grid grid-cols-1 md:grid-cols-4 gap-6">
        <div className="bg-white p-6 rounded-[2rem] shadow-sm border border-slate-100">
          <div className="bg-blue-50 w-10 h-10 rounded-xl flex items-center justify-center text-blue-600 mb-4">
            <ReceiptText size={20} />
          </div>
          <p className="text-[10px] font-black text-slate-400 uppercase mb-1">Total Sales</p>
          <h2 className="text-2xl font-black text-slate-900">₹{summaryData.taxableValue.toLocaleString()}</h2>
        </div>
        <div className="bg-white p-6 rounded-[2rem] shadow-sm border border-slate-100">
          <div className="bg-amber-50 w-10 h-10 rounded-xl flex items-center justify-center text-amber-600 mb-4">
            <HandCoins size={20} />
          </div>
          <p className="text-[10px] font-black text-slate-400 uppercase mb-1">Loans Issued</p>
          <h2 className="text-2xl font-black text-slate-900">₹{loanStats.totalPrincipal.toLocaleString()}</h2>
        </div>
        <div className="bg-white p-6 rounded-[2rem] shadow-sm border border-slate-100">
          <div className="bg-green-50 w-10 h-10 rounded-xl flex items-center justify-center text-green-600 mb-4">
            <Calculator size={20} />
          </div>
          <p className="text-[10px] font-black text-slate-400 uppercase mb-1">Total GST</p>
          <h2 className="text-2xl font-black text-green-700">₹{summaryData.totalGst.toLocaleString()}</h2>
        </div>
        <div className="bg-slate-900 p-6 rounded-[2rem] shadow-xl text-white">
          <div className="bg-slate-800 w-10 h-10 rounded-xl flex items-center justify-center text-amber-500 mb-4">
            <ArrowUpRight size={20} />
          </div>
          <p className="text-[10px] font-black text-slate-500 uppercase mb-1 tracking-widest">Net Cash Outflow</p>
          <h2 className="text-2xl font-black text-white">₹{(summaryData.taxableValue + loanStats.totalPrincipal).toLocaleString()}</h2>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
        {/* Sales Table */}
        <div className="bg-white rounded-[2.5rem] shadow-xl border border-slate-100 overflow-hidden">
          <div className="p-8 border-b bg-slate-50/50">
            <h3 className="font-black text-slate-900 tracking-tight flex items-center gap-2">
              <ReceiptText size={18} className="text-blue-500" /> Recent Sales Journal
            </h3>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs">
              <thead className="bg-slate-50">
                <tr className="text-[9px] font-black text-slate-400 uppercase tracking-widest">
                  <th className="px-6 py-4">ID</th>
                  <th className="px-6 py-4">Date</th>
                  <th className="px-6 py-4">Amt</th>
                  <th className="px-6 py-4">GST</th>
                </tr>
              </thead>
              <tbody className="divide-y">
                {MOCK_INVOICES.map(inv => (
                  <tr key={inv.id}>
                    <td className="px-6 py-4 font-bold">{inv.id}</td>
                    <td className="px-6 py-4 text-slate-500">{inv.date}</td>
                    <td className="px-6 py-4 font-black">₹{inv.subTotal.toLocaleString()}</td>
                    <td className="px-6 py-4 text-green-600 font-bold">₹{inv.gstAmount.toLocaleString()}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {/* Loan Table */}
        <div className="bg-white rounded-[2.5rem] shadow-xl border border-slate-100 overflow-hidden">
          <div className="p-8 border-b bg-slate-50/50">
            <h3 className="font-black text-slate-900 tracking-tight flex items-center gap-2">
              <HandCoins size={18} className="text-amber-500" /> Recent Loan Journal
            </h3>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-left text-xs">
              <thead className="bg-slate-50">
                <tr className="text-[9px] font-black text-slate-400 uppercase tracking-widest">
                  <th className="px-6 py-4">ID</th>
                  <th className="px-6 py-4">Date</th>
                  <th className="px-6 py-4">Principal</th>
                  <th className="px-6 py-4">Int %</th>
                </tr>
              </thead>
              <tbody className="divide-y">
                {MOCK_LOANS.map(loan => (
                  <tr key={loan.id}>
                    <td className="px-6 py-4 font-bold">{loan.id}</td>
                    <td className="px-6 py-4 text-slate-500">{loan.startDate}</td>
                    <td className="px-6 py-4 font-black">₹{loan.principalAmount.toLocaleString()}</td>
                    <td className="px-6 py-4 text-amber-600 font-bold">{loan.interestRate}%</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
};
