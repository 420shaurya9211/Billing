
import React, { useMemo } from 'react';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  LineChart, Line, AreaChart, Area
} from 'recharts';
import { TrendingUp, Users, ShoppingBag, Wallet, AlertCircle } from 'lucide-react';
import { MOCK_INVOICES, MOCK_LOANS, MOCK_CUSTOMERS } from '../mockData';

const StatCard = ({ title, value, icon, color, subtitle }: any) => (
  <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100 flex items-start justify-between">
    <div>
      <p className="text-sm text-slate-500 font-medium mb-1">{title}</p>
      <h3 className="text-2xl font-bold text-slate-900">{value}</h3>
      {subtitle && <p className="text-xs text-slate-400 mt-2">{subtitle}</p>}
    </div>
    <div className={`${color} p-3 rounded-xl text-white`}>
      {icon}
    </div>
  </div>
);

export const Dashboard: React.FC = () => {
  const chartData = [
    { name: 'Mon', sales: 4000, loans: 2400 },
    { name: 'Tue', sales: 3000, loans: 1398 },
    { name: 'Wed', sales: 2000, loans: 9800 },
    { name: 'Thu', sales: 2780, loans: 3908 },
    { name: 'Fri', sales: 1890, loans: 4800 },
    { name: 'Sat', sales: 2390, loans: 3800 },
    { name: 'Sun', sales: 3490, loans: 4300 },
  ];

  return (
    <div className="space-y-8">
      <div className="flex justify-between items-center">
        <div>
          <h1 className="text-2xl font-bold text-slate-900">Vyapar Summary</h1>
          <p className="text-slate-500">Quick look at your shop's performance</p>
        </div>
        <button className="bg-amber-500 text-slate-900 px-4 py-2 rounded-lg font-semibold hover:bg-amber-600 transition-colors">
          Generate Monthly Report
        </button>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
        <StatCard 
          title="Total Sales (May)" 
          value="₹4,25,000" 
          icon={<ShoppingBag size={20} />} 
          color="bg-blue-500"
          subtitle="+12% from last month"
        />
        <StatCard 
          title="Active Loans" 
          value="14" 
          icon={<Wallet size={20} />} 
          color="bg-green-500"
          subtitle="Total Portfolio: ₹12.5L"
        />
        <StatCard 
          title="New Customers" 
          value="24" 
          icon={<Users size={20} />} 
          color="bg-purple-500"
          subtitle="4 conversion today"
        />
        <StatCard 
          title="Overdue Risk" 
          value="3 Loans" 
          icon={<AlertCircle size={20} />} 
          color="bg-rose-500"
          subtitle="Action required"
        />
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
        <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100">
          <div className="flex items-center justify-between mb-6">
            <h3 className="font-bold text-slate-800">Weekly Performance</h3>
            <div className="flex gap-2">
              <span className="flex items-center gap-1 text-xs text-slate-500">
                <div className="w-2 h-2 rounded-full bg-blue-500" /> Sales
              </span>
              <span className="flex items-center gap-1 text-xs text-slate-500">
                <div className="w-2 h-2 rounded-full bg-green-500" /> Loans
              </span>
            </div>
          </div>
          <div className="h-64">
            <ResponsiveContainer width="100%" height="100%">
              <AreaChart data={chartData}>
                <defs>
                  <linearGradient id="colorSales" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="5%" stopColor="#3b82f6" stopOpacity={0.1}/>
                    <stop offset="95%" stopColor="#3b82f6" stopOpacity={0}/>
                  </linearGradient>
                </defs>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                <XAxis dataKey="name" axisLine={false} tickLine={false} tick={{fill: '#94a3b8', fontSize: 12}} />
                <YAxis axisLine={false} tickLine={false} tick={{fill: '#94a3b8', fontSize: 12}} />
                <Tooltip 
                  contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                />
                <Area type="monotone" dataKey="sales" stroke="#3b82f6" fillOpacity={1} fill="url(#colorSales)" />
                <Area type="monotone" dataKey="loans" stroke="#22c55e" fillOpacity={0} />
              </AreaChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100">
          <h3 className="font-bold text-slate-800 mb-6">Recent Transactions</h3>
          <div className="space-y-4">
            {MOCK_INVOICES.map(inv => (
              <div key={inv.id} className="flex items-center justify-between p-4 bg-slate-50 rounded-xl hover:bg-slate-100 transition-colors">
                <div className="flex items-center gap-3">
                  <div className="bg-white p-2 rounded-lg shadow-sm">
                    <TrendingUp className="text-blue-500" size={18} />
                  </div>
                  <div>
                    <p className="text-sm font-semibold text-slate-900">Invoice {inv.id}</p>
                    <p className="text-xs text-slate-500">{inv.date}</p>
                  </div>
                </div>
                <div className="text-right">
                  <p className="text-sm font-bold text-slate-900">₹{inv.totalAmount.toLocaleString()}</p>
                  <p className="text-[10px] text-green-600 font-bold uppercase">Paid</p>
                </div>
              </div>
            ))}
            {MOCK_LOANS.filter(l => l.status === 'OVERDUE').map(loan => (
              <div key={loan.id} className="flex items-center justify-between p-4 bg-rose-50 rounded-xl border border-rose-100">
                <div className="flex items-center gap-3">
                  <div className="bg-white p-2 rounded-lg shadow-sm">
                    <AlertCircle className="text-rose-500" size={18} />
                  </div>
                  <div>
                    <p className="text-sm font-semibold text-slate-900">Gold Loan {loan.id}</p>
                    <p className="text-xs text-slate-500">Due: {loan.dueDate}</p>
                  </div>
                </div>
                <div className="text-right">
                  <p className="text-sm font-bold text-slate-900">₹{loan.principalAmount.toLocaleString()}</p>
                  <p className="text-[10px] text-rose-600 font-bold uppercase">Overdue</p>
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
};
