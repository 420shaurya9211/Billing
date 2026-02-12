
import React, { useState, useEffect } from 'react';
import { NavLink } from 'react-router-dom';
import { 
  LayoutDashboard, 
  Users, 
  Receipt, 
  HandCoins, 
  Sparkles, 
  Settings,
  Store,
  FileText,
  RefreshCw,
  Clock
} from 'lucide-react';
import { munimAI } from '../geminiService';

export const Layout: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [rates, setRates] = useState({ gold24k: 7250, gold22k: 6850, silver: 94500, timestamp: '' });
  const [loadingRates, setLoadingRates] = useState(false);

  const fetchRates = async () => {
    setLoadingRates(true);
    try {
      const liveRates = await munimAI.getLiveMetalRates();
      if (liveRates) {
        setRates({
          gold24k: liveRates.gold24k,
          gold22k: liveRates.gold22k,
          silver: liveRates.silver,
          timestamp: new Date().toLocaleTimeString()
        });
      }
    } catch (error) {
      console.error("Failed to fetch rates:", error);
    } finally {
      setLoadingRates(false);
    }
  };

  useEffect(() => {
    fetchRates();
  }, []);

  const navItems = [
    { icon: <LayoutDashboard size={20} />, label: 'Dashboard', path: '/' },
    { icon: <Users size={20} />, label: 'Customers', path: '/customers' },
    { icon: <Receipt size={20} />, label: 'Billing', path: '/billing' },
    { icon: <FileText size={20} />, label: 'GST Reports', path: '/gst' },
    { icon: <HandCoins size={20} />, label: 'Gold Loans', path: '/loans' },
    { icon: <Sparkles size={20} />, label: 'Munim AI', path: '/ai' },
    { icon: <Settings size={20} />, label: 'Settings', path: '/settings' },
  ];

  return (
    <div className="flex h-screen bg-slate-50 overflow-hidden">
      {/* Sidebar */}
      <aside className="w-64 bg-slate-900 text-white flex flex-col">
        <div className="p-6 flex items-center gap-3 border-b border-slate-800">
          <div className="bg-amber-500 p-2 rounded-lg">
            <Store size={24} className="text-slate-900" />
          </div>
          <div>
            <h1 className="brand-font text-xl font-bold tracking-tight text-white">MunimAI</h1>
            <p className="text-[10px] text-slate-400 uppercase tracking-widest">Jewellers ERP</p>
          </div>
        </div>

        <nav className="flex-1 p-4 space-y-2 mt-4">
          {navItems.map((item) => (
            <NavLink
              key={item.path}
              to={item.path}
              className={({ isActive }) =>
                `flex items-center gap-3 p-3 rounded-xl transition-all ${
                  isActive 
                    ? 'bg-amber-500 text-slate-900 font-semibold shadow-lg shadow-amber-500/20' 
                    : 'text-slate-400 hover:bg-slate-800 hover:text-white'
                }`
              }
            >
              {item.icon}
              <span>{item.label}</span>
            </NavLink>
          ))}
        </nav>

        <div className="p-4 border-t border-slate-800 space-y-3">
          <div className="bg-slate-800/50 p-4 rounded-xl border border-slate-700 relative overflow-hidden group">
            <div className="flex justify-between items-center mb-3">
              <p className="text-[10px] text-slate-500 uppercase font-bold tracking-wider flex items-center gap-1.5">
                <span className="relative flex h-2 w-2">
                  <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-green-400 opacity-75"></span>
                  <span className="relative inline-flex rounded-full h-2 w-2 bg-green-500"></span>
                </span>
                Live Market Rates
              </p>
              <button 
                onClick={fetchRates}
                disabled={loadingRates}
                className="text-slate-500 hover:text-amber-500 transition-colors disabled:opacity-50"
              >
                <RefreshCw size={14} className={loadingRates ? 'animate-spin' : ''} />
              </button>
            </div>
            
            <div className="space-y-2">
              <div className="flex justify-between items-center">
                <p className="text-xs text-slate-400 font-medium">Gold (24K)</p>
                <p className="text-sm font-bold text-amber-500">₹{rates.gold24k}/g</p>
              </div>
              <div className="flex justify-between items-center">
                <p className="text-xs text-slate-400 font-medium">Gold (22K)</p>
                <p className="text-sm font-bold text-amber-400">₹{rates.gold22k}/g</p>
              </div>
              <div className="flex justify-between items-center">
                <p className="text-xs text-slate-400 font-medium">Silver</p>
                <p className="text-sm font-bold text-slate-200">₹{rates.silver.toLocaleString()}/kg</p>
              </div>
            </div>
            
            {rates.timestamp && (
              <div className="mt-3 pt-3 border-t border-slate-700/50 flex items-center gap-1.5 text-[9px] text-slate-500">
                <Clock size={10} />
                Updated: {rates.timestamp}
              </div>
            )}

            {loadingRates && (
              <div className="absolute inset-0 bg-slate-900/40 backdrop-blur-[1px] flex items-center justify-center">
                <RefreshCw size={20} className="text-amber-500 animate-spin" />
              </div>
            )}
          </div>
          
          <div className="text-[10px] text-center text-slate-600 px-2 leading-tight">
            Rates sourced via MunimAI Search (Moneycontrol)
          </div>
        </div>
      </aside>

      {/* Main Content */}
      <main className="flex-1 overflow-y-auto">
        <header className="h-16 bg-white border-b flex items-center justify-between px-8 sticky top-0 z-10">
          <h2 className="text-lg font-semibold text-slate-800">Jewellery Shop Management</h2>
          <div className="flex items-center gap-4">
            <div className="text-right">
              <p className="text-sm font-medium text-slate-900 uppercase">PHOOL CHANDRA SARAF</p>
              <p className="text-xs text-slate-500">Koraon, Allahabad - 212306</p>
            </div>
            <div className="w-10 h-10 rounded-full bg-slate-200 flex items-center justify-center text-slate-600 font-bold border-2 border-slate-100 uppercase">
              PCS
            </div>
          </div>
        </header>
        <div className="p-8">
          {children}
        </div>
      </main>
    </div>
  );
};
