
import React, { useState, useMemo } from 'react';
// Added Users to imports
import { Plus, Trash2, Printer, Save, Calculator, X, Gem, RefreshCcw, Info, ReceiptText, Users } from 'lucide-react';
import { InvoiceItem, Customer, MetalReturnItem } from '../types';
import { MOCK_CUSTOMERS } from '../mockData';

const InvoiceModal: React.FC<{ 
  show: boolean; 
  onClose: () => void; 
  data: any; 
  customer: Customer | null;
  billType: 'PAKKA' | 'KACHA';
  items: Partial<InvoiceItem>[];
  returns: MetalReturnItem[];
}> = ({ show, onClose, data, customer, billType, items, returns }) => {
  if (!show) return null;

  return (
    <div className="fixed inset-0 bg-black/80 backdrop-blur-md z-50 flex items-center justify-center p-4">
      <div className="bg-white w-full max-w-4xl rounded-[2.5rem] overflow-hidden shadow-2xl flex flex-col max-h-[95vh] border border-white/20">
        <div className="p-8 border-b flex justify-between items-center bg-slate-900 text-white">
          <div className="flex items-center gap-4">
            <div className="bg-amber-500 p-2.5 rounded-2xl">
              <ReceiptText size={24} className="text-slate-900" />
            </div>
            <div>
              <h3 className="font-bold text-xl tracking-tight">Invoice Preview</h3>
              <p className="text-xs text-slate-400 font-medium uppercase tracking-widest">
                {billType === 'PAKKA' ? 'Tax Invoice' : 'Estimate / Kacha Bill'}
              </p>
            </div>
          </div>
          <button onClick={onClose} className="p-2 hover:bg-slate-800 rounded-full transition-colors text-slate-400 hover:text-white">
            <X size={24} />
          </button>
        </div>
        
        <div id="invoice-bill" className="flex-1 overflow-y-auto p-12 bg-white text-slate-800 font-serif">
          {/* Header */}
          <div className="flex justify-between items-start mb-12">
            <div>
              <h1 className="text-5xl font-bold text-amber-600 brand-font mb-1 uppercase tracking-tight">PHOOL CHANDRA SARAF</h1>
              <p className="text-sm font-bold font-sans text-slate-500 tracking-widest uppercase">ASHISH JEWELLERS</p>
              <div className="mt-6 font-sans text-xs space-y-1 text-slate-600">
                <p>Koraon, Allahabad, Uttar Pradesh - 212306</p>
                <p>Tel: +91 9123456789 | Email: contact@ashishjewellers.com</p>
                {billType === 'PAKKA' && <p className="font-bold text-slate-900">GSTIN: 09XXXXXXXXXXXXX</p>}
              </div>
            </div>
            <div className="text-right font-sans">
              <h2 className="text-3xl font-bold uppercase tracking-[0.2em] text-slate-300 mb-6">
                {billType === 'PAKKA' ? 'Invoice' : 'Estimate'}
              </h2>
              <div className="space-y-1.5 text-sm">
                <p><span className="text-slate-400">No:</span> <span className="font-bold text-slate-900">#{Math.floor(Math.random() * 9000) + 1000}</span></p>
                <p><span className="text-slate-400">Date:</span> <span className="font-bold text-slate-900">{new Date().toLocaleDateString('en-IN')}</span></p>
              </div>
            </div>
          </div>

          {/* Billing Info */}
          <div className="grid grid-cols-2 gap-16 mb-12 font-sans border-y py-8 border-slate-100">
            <div>
              <h4 className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-4">Customer Details</h4>
              <p className="text-lg font-bold text-slate-900">{customer?.name || 'Cash Customer'}</p>
              <p className="text-sm text-slate-600 mt-1 leading-relaxed">{customer?.address || 'N/A'}</p>
              <p className="text-sm font-semibold text-slate-800 mt-2">Ph: {customer?.phone || 'N/A'}</p>
              {billType === 'PAKKA' && customer?.gstin && (
                <p className="text-xs text-amber-600 font-bold mt-2">GSTIN: {customer.gstin}</p>
              )}
            </div>
            <div className="text-right">
              <h4 className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-4">Daily Market Benchmark</h4>
              <div className="space-y-2 text-sm">
                <div className="flex justify-between"><span className="text-slate-400">Gold (24K)</span> <span className="font-bold">₹7,250/g</span></div>
                <div className="flex justify-between"><span className="text-slate-400">Gold (22K)</span> <span className="font-bold">₹6,850/g</span></div>
                <div className="flex justify-between"><span className="text-slate-400">Silver</span> <span className="font-bold">₹94.5/g</span></div>
              </div>
            </div>
          </div>

          {/* Items Table */}
          <div className="mb-10 font-sans">
            <h4 className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-4">New Purchases</h4>
            <table className="w-full">
              <thead className="bg-slate-50">
                <tr className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">
                  <th className="py-4 px-4 text-left border-b">Sl.</th>
                  <th className="py-4 px-4 text-left border-b">Item Description</th>
                  <th className="py-4 px-4 text-right border-b">Weight</th>
                  <th className="py-4 px-4 text-right border-b">Purity</th>
                  <th className="py-4 px-4 text-right border-b">Rate</th>
                  <th className="py-4 px-4 text-right border-b">Making</th>
                  <th className="py-4 px-4 text-right border-b">Taxable</th>
                </tr>
              </thead>
              <tbody className="divide-y text-xs">
                {items.map((item, idx) => {
                  const baseValue = (item.weight || 0) * (item.ratePerGram || 0);
                  const making = item.makingChargeType === 'FIXED' 
                    ? (item.makingCharges || 0) 
                    : (baseValue * ((item.makingCharges || 0) / 100));
                  const taxable = baseValue + making - (item.discount || 0);
                  
                  return (
                    <tr key={idx} className="hover:bg-slate-50 transition-colors">
                      <td className="py-4 px-4 text-slate-400">{idx + 1}</td>
                      <td className="py-4 px-4">
                        <p className="font-bold text-slate-900">{item.description || 'Custom Jewellery'}</p>
                        <p className="text-[10px] text-slate-400 uppercase font-medium mt-0.5">{item.metal}</p>
                      </td>
                      <td className="py-4 px-4 text-right font-semibold">{item.weight}g</td>
                      <td className="py-4 px-4 text-right"><span className="bg-amber-50 text-amber-700 px-2 py-0.5 rounded font-bold">{item.purity}</span></td>
                      <td className="py-4 px-4 text-right text-slate-500">₹{item.ratePerGram?.toLocaleString()}</td>
                      <td className="py-4 px-4 text-right text-slate-500">₹{making.toLocaleString()}</td>
                      <td className="py-4 px-4 text-right font-black text-slate-900">₹{taxable.toLocaleString()}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>

          {/* Return Table */}
          {returns.length > 0 && (
            <div className="mb-10 font-sans animate-in fade-in slide-in-from-top-2 duration-500">
              <h4 className="text-[10px] font-bold text-rose-500 uppercase tracking-widest mb-4 flex items-center gap-2">
                <RefreshCcw size={12} /> Old Metal Return (Credit)
              </h4>
              <table className="w-full">
                <thead className="bg-rose-50/50">
                  <tr className="text-[10px] font-bold text-rose-400 uppercase tracking-wider">
                    <th className="py-3 px-4 text-left border-b">Sl.</th>
                    <th className="py-3 px-4 text-left border-b">Description</th>
                    <th className="py-3 px-4 text-right border-b">Weight</th>
                    <th className="py-4 px-4 text-right border-b">Purity</th>
                    <th className="py-3 px-4 text-right border-b">Exch. Rate</th>
                    <th className="py-3 px-4 text-right border-b">Credit Value</th>
                  </tr>
                </thead>
                <tbody className="divide-y text-xs text-rose-900">
                  {returns.map((ret, idx) => (
                    <tr key={ret.id} className="bg-rose-50/20">
                      <td className="py-3 px-4">{idx + 1}</td>
                      <td className="py-3 px-4 font-medium">{ret.description}</td>
                      <td className="py-3 px-4 text-right font-bold">{ret.weight}g</td>
                      <td className="py-3 px-4 text-right">{ret.purity}</td>
                      <td className="py-3 px-4 text-right">₹{ret.ratePerGram.toLocaleString()}</td>
                      <td className="py-3 px-4 text-right font-black text-rose-600">-₹{ret.totalValue.toLocaleString()}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {/* Final Summary */}
          <div className="flex justify-end font-sans">
            <div className="w-72 space-y-3">
              <div className="flex justify-between text-xs text-slate-500">
                <span>Gross Purchase Value:</span>
                <span className="font-bold text-slate-900">₹{data.pureSubTotal.toLocaleString()}</span>
              </div>
              
              {billType === 'PAKKA' && (
                <>
                  <div className="flex justify-between text-xs text-slate-500">
                    <span>CGST (1.5%):</span>
                    <span className="font-bold text-slate-900">₹{data.cgstAmount.toLocaleString()}</span>
                  </div>
                  <div className="flex justify-between text-xs text-slate-500">
                    <span>SGST (1.5%):</span>
                    <span className="font-bold text-slate-900">₹{data.sgstAmount.toLocaleString()}</span>
                  </div>
                </>
              )}

              {returns.length > 0 && (
                <div className="flex justify-between text-xs text-rose-600 font-bold pt-1 border-t border-slate-100">
                  <span>Old Metal Adjustment:</span>
                  <span>-₹{data.totalReturn.toLocaleString()}</span>
                </div>
              )}

              <div className="flex justify-between items-center bg-slate-900 text-white p-5 rounded-2xl shadow-xl shadow-slate-200 mt-4 transition-all hover:scale-105">
                <span className="text-[10px] font-bold uppercase tracking-widest">Grand Payable</span>
                <span className="text-2xl font-black">₹{data.finalPayable.toLocaleString()}</span>
              </div>
            </div>
          </div>

          {/* Footer Signatures */}
          <div className="mt-20 border-t pt-10 text-[9px] text-slate-400 font-sans grid grid-cols-2 gap-16">
            <div className="space-y-4">
              <div>
                <h5 className="font-black uppercase text-slate-600 mb-2 tracking-tighter">Conditions of Sale</h5>
                <ul className="list-disc pl-3 space-y-1 leading-tight">
                  <li>Estimated weight/purity shown is for reference in kacha bills.</li>
                  <li>In Pakka bills, purity is hallmarked and guaranteed.</li>
                  <li>Tax Invoice carries full buy-back guarantee as per standard rates.</li>
                  <li>Subject to Mumbai Jurisdiction.</li>
                </ul>
              </div>
              <div className="italic text-slate-300">Generated via MunimAI Jewellers ERP</div>
            </div>
            <div className="flex flex-col items-end justify-end">
              <div className="text-center w-48">
                <div className="h-12 mb-2 border-b-2 border-slate-100 border-dashed"></div>
                <p className="font-black uppercase text-slate-600 tracking-widest">Authorised Signatory</p>
                <p className="text-[8px] text-slate-300 mt-1 italic uppercase">For PHOOL CHANDRA SARAF</p>
              </div>
            </div>
          </div>
        </div>

        <div className="p-8 bg-slate-900 flex justify-end gap-4">
          <button onClick={onClose} className="px-8 py-3 rounded-2xl text-slate-400 hover:text-white font-bold transition-all hover:bg-slate-800">Discard</button>
          <button className="px-10 py-3 bg-amber-500 text-slate-900 rounded-2xl font-black flex items-center gap-3 hover:bg-amber-400 transition-all shadow-2xl shadow-amber-500/20 active:scale-95">
            <Printer size={20} /> Print {billType === 'PAKKA' ? 'Invoice' : 'Estimate'}
          </button>
        </div>
      </div>
    </div>
  );
};

export const Billing: React.FC = () => {
  const [selectedCustomer, setSelectedCustomer] = useState<Customer | null>(null);
  const [showPreview, setShowPreview] = useState(false);
  const [billType, setBillType] = useState<'PAKKA' | 'KACHA'>('PAKKA');
  
  const [items, setItems] = useState<Partial<InvoiceItem>[]>([
    { id: '1', description: '', metal: 'GOLD', weight: 0, purity: '22K', ratePerGram: 6850, makingCharges: 0, makingChargeType: 'FIXED', discount: 0 }
  ]);

  const [returns, setReturns] = useState<MetalReturnItem[]>([]);

  const calculateTotals = useMemo(() => {
    let subTotal = 0;
    items.forEach(item => {
      const baseValue = (item.weight || 0) * (item.ratePerGram || 0);
      const charges = item.makingChargeType === 'FIXED' 
        ? (item.makingCharges || 0) 
        : (baseValue * ((item.makingCharges || 0) / 100));
      const taxable = baseValue + charges - (item.discount || 0);
      subTotal += taxable;
    });

    const totalReturn = returns.reduce((acc, curr) => acc + curr.totalValue, 0);
    
    // GST is on the new item's total value
    const cgstRate = billType === 'PAKKA' ? 1.5 : 0;
    const sgstRate = billType === 'PAKKA' ? 1.5 : 0;
    const cgstAmount = (subTotal * cgstRate) / 100;
    const sgstAmount = (subTotal * sgstRate) / 100;
    const totalGst = cgstAmount + sgstAmount;
    
    const finalPayable = (subTotal + totalGst) - totalReturn;

    return { 
      pureSubTotal: subTotal, 
      cgstAmount, 
      sgstAmount, 
      totalGst, 
      totalReturn,
      finalPayable 
    };
  }, [items, returns, billType]);

  const addItem = (metal: 'GOLD' | 'SILVER' = 'GOLD') => {
    const rate = metal === 'GOLD' ? 6850 : 94.5;
    setItems([...items, { 
      id: Date.now().toString(), 
      description: '', 
      metal,
      weight: 0, 
      purity: metal === 'GOLD' ? '22K' : '925', 
      ratePerGram: rate, 
      makingCharges: 0, 
      makingChargeType: 'FIXED', 
      discount: 0 
    }]);
  };

  const addReturnItem = () => {
    setReturns([...returns, { 
      id: Date.now().toString(), 
      description: 'Old Gold', 
      weight: 0, 
      ratePerGram: 6650, 
      purity: '22K', 
      totalValue: 0 
    }]);
  };

  // Implemented removeItem to fix "Cannot find name 'removeItem'" error
  const removeItem = (id: string) => {
    setItems(items.filter(item => item.id !== id));
  };

  const updateItem = (id: string, field: string, value: any) => {
    setItems(items.map(i => {
      if (i.id === id) {
        if (field === 'metal') {
          const rate = value === 'GOLD' ? 6850 : 94.5;
          const purity = value === 'GOLD' ? '22K' : '925';
          return { ...i, metal: value, ratePerGram: rate, purity };
        }
        return { ...i, [field]: value };
      }
      return i;
    }));
  };

  const updateReturn = (id: string, field: string, value: any) => {
    setReturns(returns.map(r => {
      if (r.id === id) {
        const updated = { ...r, [field]: value };
        updated.totalValue = updated.weight * updated.ratePerGram;
        return updated;
      }
      return r;
    }));
  };

  return (
    <div className="max-w-7xl mx-auto space-y-8 animate-in fade-in duration-700">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-end gap-6">
        <div>
          <h1 className="text-4xl font-black text-slate-900 tracking-tight flex items-center gap-3">
            Digital Munim Billing
            <div className="bg-amber-100 px-3 py-1 rounded-full text-amber-700 text-xs font-black uppercase tracking-widest">Premium</div>
          </h1>
          <p className="text-slate-500 mt-1 font-medium italic">Streamlining Indian Jewelry Accounting</p>
        </div>

        {/* Bill Type Toggle */}
        <div className="bg-slate-200 p-1.5 rounded-2xl flex items-center gap-1 shadow-inner">
          <button 
            onClick={() => setBillType('PAKKA')}
            className={`px-6 py-2 rounded-xl text-xs font-black uppercase tracking-widest transition-all ${billType === 'PAKKA' ? 'bg-white text-slate-900 shadow-md' : 'text-slate-500 hover:text-slate-700'}`}
          >
            Pakka (GST)
          </button>
          <button 
            onClick={() => setBillType('KACHA')}
            className={`px-6 py-2 rounded-xl text-xs font-black uppercase tracking-widest transition-all ${billType === 'KACHA' ? 'bg-white text-slate-900 shadow-md' : 'text-slate-500 hover:text-slate-700'}`}
          >
            Kacha (Rough)
          </button>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-4 gap-8">
        <div className="lg:col-span-3 space-y-8">
          {/* Customer & Settings Row */}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
            <div className="md:col-span-2 bg-white p-8 rounded-[2rem] shadow-xl shadow-slate-200/50 border border-slate-100 group transition-all hover:shadow-2xl">
              <h3 className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-4 flex items-center gap-2">
                <Users size={14} className="text-amber-500" /> Customer Information
              </h3>
              <div className="flex flex-col md:flex-row gap-4">
                <div className="flex-1 relative">
                  <select 
                    className="w-full pl-5 pr-12 py-4 bg-slate-50 border-2 border-transparent rounded-2xl outline-none focus:border-amber-500 focus:bg-white transition-all appearance-none font-bold text-slate-900"
                    onChange={(e) => setSelectedCustomer(MOCK_CUSTOMERS.find(c => c.id === e.target.value) || null)}
                  >
                    <option value="">Walk-in Cash Customer</option>
                    {MOCK_CUSTOMERS.map(c => (
                      <option key={c.id} value={c.id}>{c.name} ({c.phone})</option>
                    ))}
                  </select>
                  <div className="absolute right-4 top-1/2 -translate-y-1/2 pointer-events-none text-slate-400">
                    <Users size={18} />
                  </div>
                </div>
                <div className="flex-1 p-4 bg-amber-50 rounded-2xl border-2 border-amber-100/50 flex items-center justify-between">
                  <span className="text-[10px] font-bold text-amber-600 uppercase">GST Status</span>
                  {billType === 'PAKKA' ? (
                    <span className="text-xs font-black text-amber-700 bg-amber-200/50 px-3 py-1 rounded-lg">ACTIVE INVOICE</span>
                  ) : (
                    <span className="text-xs font-black text-slate-500 bg-slate-200/50 px-3 py-1 rounded-lg">NON-GST ESTIMATE</span>
                  )}
                </div>
              </div>
            </div>
            
            <div className="bg-slate-900 p-8 rounded-[2rem] text-white flex flex-col justify-center">
              <p className="text-[10px] font-black text-slate-500 uppercase tracking-[0.2em] mb-3">Today's Rates</p>
              <div className="space-y-3">
                <div className="flex justify-between items-center"><span className="text-xs font-bold text-slate-400">Gold (22K)</span> <span className="font-black text-amber-500">₹6,850</span></div>
                <div className="flex justify-between items-center"><span className="text-xs font-bold text-slate-400">Silver (99.9)</span> <span className="font-black text-slate-100">₹94.5</span></div>
              </div>
            </div>
          </div>

          {/* New Purchase Items */}
          <div className="bg-white p-8 rounded-[2rem] shadow-xl shadow-slate-200/50 border border-slate-100">
            <div className="flex justify-between items-center mb-8">
              <h3 className="text-sm font-black text-slate-900 uppercase tracking-widest flex items-center gap-3">
                <div className="w-1.5 h-6 bg-amber-500 rounded-full" /> New Ornaments
              </h3>
              <div className="flex gap-3">
                <button 
                  onClick={() => addItem('GOLD')} 
                  className="bg-amber-50 text-amber-600 px-5 py-2.5 rounded-xl font-black text-xs flex items-center gap-2 hover:bg-amber-100 transition-all active:scale-95 shadow-lg shadow-amber-500/10"
                >
                  <Plus size={16} /> ADD GOLD
                </button>
                <button 
                  onClick={() => addItem('SILVER')} 
                  className="bg-slate-100 text-slate-600 px-5 py-2.5 rounded-xl font-black text-xs flex items-center gap-2 hover:bg-slate-200 transition-all active:scale-95 shadow-lg shadow-slate-500/10"
                >
                  <Plus size={16} /> ADD SILVER
                </button>
              </div>
            </div>
            
            <div className="space-y-6">
              {items.map((item, index) => (
                <div key={item.id} className="group relative p-6 bg-slate-50/50 rounded-3xl space-y-5 border border-slate-100 transition-all hover:bg-white hover:shadow-xl hover:border-amber-200/50">
                  <div className="grid grid-cols-1 md:grid-cols-12 gap-5">
                    <div className="md:col-span-5">
                      <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-2 block">Ornaments Description</label>
                      <input 
                        className="w-full p-3.5 bg-white border-2 border-slate-100 rounded-xl text-sm font-bold text-slate-900 placeholder:font-normal focus:border-amber-400 outline-none transition-all" 
                        placeholder="e.g., Nakshatra Gold Necklace"
                        value={item.description}
                        onChange={(e) => updateItem(item.id!, 'description', e.target.value)}
                      />
                    </div>
                    <div className="md:col-span-2">
                      <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-2 block">Metal</label>
                      <select 
                        className="w-full p-3.5 bg-white border-2 border-slate-100 rounded-xl text-sm font-black text-amber-600 outline-none transition-all focus:border-amber-400"
                        value={item.metal}
                        onChange={(e) => updateItem(item.id!, 'metal', e.target.value)}
                      >
                        <option value="GOLD">GOLD</option>
                        <option value="SILVER">SILVER</option>
                      </select>
                    </div>
                    <div className="md:col-span-3">
                      <label className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-2 block">Purity (Karat/%)</label>
                      <select 
                        className="w-full p-3.5 bg-white border-2 border-slate-100 rounded-xl text-sm font-bold text-slate-900 outline-none transition-all focus:border-amber-400"
                        value={item.purity}
                        onChange={(e) => updateItem(item.id!, 'purity', e.target.value)}
                      >
                        {item.metal === 'GOLD' ? (
                          <>
                            <option value="24K">24K (Pure)</option>
                            <option value="22K">22K (Jewellery)</option>
                            <option value="18K">18K (Diamond)</option>
                          </>
                        ) : (
                          <>
                            <option value="999">99.9 Fine</option>
                            <option value="925">92.5 Sterling</option>
                          </>
                        )}
                      </select>
                    </div>
                    <div className="md:col-span-2 flex items-end justify-center pb-2">
                      <button onClick={() => removeItem(item.id!)} className="p-3 text-slate-300 hover:text-rose-500 transition-colors bg-white rounded-xl shadow-sm hover:shadow-md">
                        <Trash2 size={20} />
                      </button>
                    </div>
                  </div>

                  <div className="grid grid-cols-2 md:grid-cols-5 gap-5">
                    <div className="bg-white p-4 rounded-2xl border-2 border-slate-100">
                      <label className="text-[9px] font-black text-slate-400 uppercase tracking-widest mb-1.5 block">Gross Wt (g)</label>
                      <input 
                        type="number" 
                        className="w-full bg-transparent text-lg font-black text-slate-900 outline-none" 
                        value={item.weight}
                        onChange={(e) => updateItem(item.id!, 'weight', parseFloat(e.target.value))}
                      />
                    </div>
                    <div className="bg-white p-4 rounded-2xl border-2 border-slate-100">
                      <label className="text-[9px] font-black text-slate-400 uppercase tracking-widest mb-1.5 block">Rate/g</label>
                      <input 
                        type="number" 
                        className="w-full bg-transparent text-lg font-bold text-slate-900 outline-none" 
                        value={item.ratePerGram}
                        onChange={(e) => updateItem(item.id!, 'ratePerGram', parseFloat(e.target.value))}
                      />
                    </div>
                    <div className="bg-white p-4 rounded-2xl border-2 border-slate-100">
                      <label className="text-[9px] font-black text-slate-400 uppercase tracking-widest mb-1.5 block">Making Charges</label>
                      <input 
                        type="number" 
                        className="w-full bg-transparent text-lg font-bold text-slate-900 outline-none" 
                        value={item.makingCharges}
                        onChange={(e) => updateItem(item.id!, 'makingCharges', parseFloat(e.target.value))}
                      />
                    </div>
                    <div className="bg-white p-4 rounded-2xl border-2 border-slate-100">
                      <label className="text-[9px] font-black text-slate-400 uppercase tracking-widest mb-1.5 block">Discount</label>
                      <input 
                        type="number" 
                        className="w-full bg-transparent text-lg font-bold text-green-600 outline-none" 
                        value={item.discount}
                        onChange={(e) => updateItem(item.id!, 'discount', parseFloat(e.target.value))}
                      />
                    </div>
                    <div className="bg-slate-900 p-4 rounded-2xl shadow-xl shadow-slate-200">
                      <label className="text-[9px] font-black text-slate-500 uppercase tracking-widest mb-1.5 block">Item Total</label>
                      <div className="text-lg font-black text-amber-500">
                        ₹{(((item.weight || 0) * (item.ratePerGram || 0)) + (item.makingCharges || 0) - (item.discount || 0)).toLocaleString()}
                      </div>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          </div>

          {/* Metal Return Section */}
          <div className="bg-white p-8 rounded-[2rem] shadow-xl shadow-slate-200/50 border border-slate-100 border-dashed border-rose-200">
            <div className="flex justify-between items-center mb-6">
              <h3 className="text-sm font-black text-rose-600 uppercase tracking-widest flex items-center gap-3">
                <RefreshCcw size={18} /> Metal Exchange / Return
              </h3>
              <button 
                onClick={addReturnItem}
                className="bg-rose-50 text-rose-600 px-5 py-2.5 rounded-xl font-black text-xs flex items-center gap-2 hover:bg-rose-100 transition-all active:scale-95 shadow-lg shadow-rose-500/10"
              >
                <Plus size={16} /> ADD OLD METAL
              </button>
            </div>

            <div className="space-y-4">
              {returns.map((ret, idx) => (
                <div key={ret.id} className="flex flex-col md:flex-row gap-4 items-end bg-rose-50/20 p-5 rounded-2xl border border-rose-100 transition-all hover:bg-rose-50/50">
                  <div className="flex-1 space-y-2">
                    <label className="text-[9px] font-black text-rose-400 uppercase">Old Metal Description</label>
                    <input 
                      className="w-full p-2.5 bg-white border border-rose-100 rounded-xl text-sm font-bold text-slate-900"
                      value={ret.description}
                      onChange={(e) => updateReturn(ret.id, 'description', e.target.value)}
                    />
                  </div>
                  <div className="w-24 space-y-2">
                    <label className="text-[9px] font-black text-rose-400 uppercase">Weight (g)</label>
                    <input 
                      type="number"
                      className="w-full p-2.5 bg-white border border-rose-100 rounded-xl text-sm font-black text-slate-900"
                      value={ret.weight}
                      onChange={(e) => updateReturn(ret.id, 'weight', parseFloat(e.target.value))}
                    />
                  </div>
                  <div className="w-32 space-y-2">
                    <label className="text-[9px] font-black text-rose-400 uppercase">Return Rate</label>
                    <input 
                      type="number"
                      className="w-full p-2.5 bg-white border border-rose-100 rounded-xl text-sm font-bold text-slate-900"
                      value={ret.ratePerGram}
                      onChange={(e) => updateReturn(ret.id, 'ratePerGram', parseFloat(e.target.value))}
                    />
                  </div>
                  <div className="w-40 bg-white p-2.5 rounded-xl border border-rose-100 flex flex-col justify-center">
                    <span className="text-[8px] font-black text-rose-300 uppercase">Total Credit</span>
                    <span className="text-sm font-black text-rose-600">₹{ret.totalValue.toLocaleString()}</span>
                  </div>
                  <button 
                    onClick={() => setReturns(returns.filter(r => r.id !== ret.id))}
                    className="p-3 text-rose-300 hover:text-rose-600 transition-colors"
                  >
                    <Trash2 size={18} />
                  </button>
                </div>
              ))}
              {returns.length === 0 && (
                <div className="text-center py-6 text-slate-300 text-xs font-medium border-2 border-dashed border-slate-100 rounded-2xl">
                  No old metal returns added for this transaction.
                </div>
              )}
            </div>
          </div>
        </div>

        {/* Sidebar Summary */}
        <div className="lg:col-span-1 space-y-6">
          <div className="bg-slate-900 p-10 rounded-[2.5rem] shadow-2xl shadow-slate-900/40 sticky top-24 border border-slate-800">
            <h3 className="text-[10px] font-black text-slate-500 uppercase tracking-[0.2em] mb-10 flex items-center gap-3">
              <Calculator size={18} className="text-amber-500" /> Summary Ledger
            </h3>
            
            <div className="space-y-6">
              <div className="flex justify-between items-center text-slate-400">
                <span className="text-xs font-bold">New Sub-total</span>
                <span className="font-black text-slate-100">₹{calculateTotals.pureSubTotal.toLocaleString()}</span>
              </div>
              
              {billType === 'PAKKA' && (
                <div className="space-y-3 pt-3 border-t border-slate-800">
                  <div className="flex justify-between items-center text-slate-500">
                    <span className="text-[10px] font-bold">CGST (1.5%)</span>
                    <span className="font-black text-slate-300">₹{calculateTotals.cgstAmount.toLocaleString()}</span>
                  </div>
                  <div className="flex justify-between items-center text-slate-500">
                    <span className="text-[10px] font-bold">SGST (1.5%)</span>
                    <span className="font-black text-slate-300">₹{calculateTotals.sgstAmount.toLocaleString()}</span>
                  </div>
                </div>
              )}

              {calculateTotals.totalReturn > 0 && (
                <div className="flex justify-between items-center text-rose-400 pt-3 border-t border-slate-800">
                  <span className="text-xs font-black uppercase">Metal Credit</span>
                  <span className="font-black">-₹{calculateTotals.totalReturn.toLocaleString()}</span>
                </div>
              )}

              <div className="pt-10 border-t border-slate-800">
                <p className="text-[10px] text-amber-500 uppercase font-black tracking-widest mb-3">Net Final Payable</p>
                <div className="text-5xl font-black text-white tracking-tighter">
                  ₹{calculateTotals.finalPayable.toLocaleString()}
                </div>
                <p className="text-[9px] text-slate-500 mt-4 italic font-medium">
                  {billType === 'PAKKA' ? '* Includes mandatory 3% GST on jewellery purchases.' : '* Estimates are subject to slight final weight variation.'}
                </p>
              </div>

              <div className="pt-10 space-y-3">
                <button 
                  onClick={() => setShowPreview(true)}
                  className="w-full bg-slate-800 text-slate-200 py-4 rounded-2xl font-black text-xs uppercase tracking-widest hover:bg-slate-700 hover:text-white transition-all active:scale-95 flex items-center justify-center gap-2 border border-slate-700"
                >
                  <Printer size={16} /> Preview Bill
                </button>
                <button className="w-full bg-amber-500 text-slate-900 py-4 rounded-2xl font-black text-sm uppercase tracking-widest hover:bg-amber-400 transition-all active:scale-95 shadow-xl shadow-amber-500/20 flex items-center justify-center gap-2">
                  <Save size={18} /> Finalise & Sync
                </button>
              </div>
            </div>
          </div>

          <div className="bg-white p-6 rounded-[2rem] border border-slate-100 flex items-start gap-4 shadow-sm">
            <div className="bg-blue-50 p-2 rounded-xl text-blue-500">
              <Info size={20} />
            </div>
            <div>
              <p className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-1">Munim Tip</p>
              <p className="text-xs text-slate-600 font-medium leading-relaxed">
                For metal exchange, ensure the old gold rate is adjusted based on melting loss if purity is below standard 22K.
              </p>
            </div>
          </div>
        </div>
      </div>

      <InvoiceModal 
        show={showPreview} 
        onClose={() => setShowPreview(false)} 
        data={calculateTotals}
        customer={selectedCustomer}
        billType={billType}
        items={items}
        returns={returns}
      />
    </div>
  );
};
