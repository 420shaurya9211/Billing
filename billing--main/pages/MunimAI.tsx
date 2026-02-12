
import React, { useState } from 'react';
import { Sparkles, Activity, ShieldAlert, FileSearch, CalendarRange, Send, Loader2 } from 'lucide-react';
import { munimAI } from '../geminiService';
import { MOCK_CUSTOMERS, MOCK_INVOICES, MOCK_LOANS } from '../mockData';

export const MunimAI: React.FC = () => {
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState<any>(null);
  const [activeTab, setActiveTab] = useState('summary');
  const [userInput, setUserInput] = useState('');

  const handleAIAction = async (action: string) => {
    setLoading(true);
    try {
      let data: any;
      switch (action) {
        case 'customers':
          data = await munimAI.analyzeCustomers(MOCK_CUSTOMERS);
          break;
        case 'loans':
          data = await munimAI.analyzeLoanRisk(MOCK_LOANS);
          break;
        case 'summary':
          data = await munimAI.getBusinessSummary({
            sales: MOCK_INVOICES,
            loans: MOCK_LOANS,
            month: 'May 2024'
          });
          break;
        case 'festival':
          data = await munimAI.getFestivalAdvice({
            previousSales: MOCK_INVOICES,
            season: 'Wedding'
          });
          break;
        case 'query':
          data = await munimAI.queryToFilter(userInput, "Tables: Customers(name, phone), Loans(amount, status), Invoices(date, total)");
          break;
      }
      setResult(data);
    } catch (error) {
      console.error(error);
      setResult({ error: "Failed to fetch AI insights. Check API Key." });
    } finally {
      setLoading(false);
    }
  };

  const tabs = [
    { id: 'summary', icon: <Activity size={18} />, label: 'Business Audit', action: 'summary' },
    { id: 'customers', icon: <ShieldAlert size={18} />, label: 'Risk Analysis', action: 'customers' },
    { id: 'festival', icon: <CalendarRange size={18} />, label: 'Market Advice', action: 'festival' },
  ];

  return (
    <div className="space-y-6">
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-bold flex items-center gap-2">
            <Sparkles className="text-amber-500" /> Munim AI Hub
          </h1>
          <p className="text-slate-500">Intelligent jewelry business assistant powered by Gemini</p>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-4 gap-6">
        {/* Actions Sidebar */}
        <div className="lg:col-span-1 space-y-4">
          <div className="bg-white p-4 rounded-2xl shadow-sm border border-slate-100 space-y-2">
            <p className="text-xs font-bold text-slate-400 uppercase mb-3">Analysis Tools</p>
            {tabs.map((tab) => (
              <button
                key={tab.id}
                onClick={() => {
                  setActiveTab(tab.id);
                  handleAIAction(tab.action);
                }}
                className={`w-full flex items-center gap-3 p-3 rounded-xl transition-all ${
                  activeTab === tab.id 
                  ? 'bg-slate-900 text-white shadow-lg' 
                  : 'text-slate-600 hover:bg-slate-50'
                }`}
              >
                {tab.icon}
                <span className="text-sm font-medium">{tab.label}</span>
              </button>
            ))}
          </div>

          <div className="bg-amber-50 p-6 rounded-2xl border border-amber-100">
            <h4 className="font-bold text-amber-900 mb-2">Ask Munim Anything</h4>
            <div className="relative">
              <input 
                type="text" 
                value={userInput}
                onChange={(e) => setUserInput(e.target.value)}
                placeholder="e.g. Find overdue loans > 50k"
                className="w-full bg-white border border-amber-200 rounded-xl px-4 py-3 pr-10 text-sm text-slate-900 focus:ring-2 focus:ring-amber-500 outline-none"
              />
              <button 
                onClick={() => handleAIAction('query')}
                className="absolute right-2 top-1/2 -translate-y-1/2 text-amber-500 hover:text-amber-600"
              >
                <Send size={18} />
              </button>
            </div>
          </div>
        </div>

        {/* Results Panel */}
        <div className="lg:col-span-3">
          <div className="bg-white min-h-[500px] rounded-2xl shadow-sm border border-slate-100 p-8">
            {loading ? (
              <div className="flex flex-col items-center justify-center h-full space-y-4">
                <Loader2 className="animate-spin text-amber-500" size={48} />
                <p className="text-slate-500 font-medium">Munim is analyzing your ledger...</p>
              </div>
            ) : result ? (
              <div className="space-y-6">
                <div className="flex items-center justify-between border-b pb-4">
                  <h3 className="text-xl font-bold text-slate-900">AI Report: {activeTab.toUpperCase()}</h3>
                  <button onClick={() => setResult(null)} className="text-xs text-slate-400 hover:text-slate-600 underline">Clear</button>
                </div>
                
                {/* Dynamically render JSON based on structure */}
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  {Object.entries(result).map(([key, value]: [string, any]) => (
                    <div key={key} className="p-4 bg-slate-50 rounded-xl border border-slate-100">
                      <h4 className="text-xs font-bold text-slate-400 uppercase mb-2">{key.replace(/_/g, ' ')}</h4>
                      {Array.isArray(value) ? (
                        <div className="space-y-3">
                          {value.map((item, idx) => (
                            <div key={idx} className="text-sm">
                              {typeof item === 'object' ? (
                                <div className="space-y-1">
                                  {Object.entries(item).map(([ik, iv]: [string, any]) => (
                                    <div key={ik} className="flex gap-2">
                                      <span className="font-semibold text-slate-700">{ik}:</span>
                                      <span className="text-slate-600">{iv}</span>
                                    </div>
                                  ))}
                                </div>
                              ) : (
                                <div className="flex items-start gap-2">
                                  <div className="w-1.5 h-1.5 rounded-full bg-amber-500 mt-1.5" />
                                  <span className="text-slate-600">{item}</span>
                                </div>
                              )}
                            </div>
                          ))}
                        </div>
                      ) : typeof value === 'object' ? (
                        <div className="text-sm space-y-2">
                          {Object.entries(value).map(([sk, sv]: [string, any]) => (
                            <div key={sk} className="flex flex-col">
                              <span className="font-semibold text-slate-700">{sk}</span>
                              <span className="text-slate-600">{sv}</span>
                            </div>
                          ))}
                        </div>
                      ) : (
                        <p className="text-sm text-slate-700">{value}</p>
                      )}
                    </div>
                  ))}
                </div>
              </div>
            ) : (
              <div className="flex flex-col items-center justify-center h-full text-center space-y-4 text-slate-400">
                <Sparkles size={64} className="opacity-20" />
                <div>
                  <h3 className="text-lg font-bold text-slate-600">Munim AI is Ready</h3>
                  <p className="max-w-xs mx-auto">Select an analysis tool from the sidebar to generate intelligent business insights.</p>
                </div>
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};
