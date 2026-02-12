
import React, { useState } from 'react';
import { Plus, Trash2, HandCoins, User, Smartphone, Fingerprint, Gem, Scale, Calendar, Percent, Save, Search, Download, Printer, X, ReceiptText, ShieldCheck } from 'lucide-react';
import { Loan } from '../types';
import { MOCK_LOANS } from '../mockData';

const LoanDocumentModal: React.FC<{
  show: boolean;
  onClose: () => void;
  loan: Loan | null;
}> = ({ show, onClose, loan }) => {
  if (!show || !loan) return null;

  return (
    <div className="fixed inset-0 bg-black/80 backdrop-blur-md z-50 flex items-center justify-center p-4">
      <div className="bg-white w-full max-w-4xl rounded-[2.5rem] overflow-hidden shadow-2xl flex flex-col max-h-[95vh] border border-white/20">
        <div className="p-8 border-b flex justify-between items-center bg-slate-900 text-white">
          <div className="flex items-center gap-4">
            <div className="bg-amber-500 p-2.5 rounded-2xl">
              <HandCoins size={24} className="text-slate-900" />
            </div>
            <div>
              <h3 className="font-bold text-xl tracking-tight">Girvi Agreement Preview</h3>
              <p className="text-xs text-slate-400 font-medium uppercase tracking-widest">
                Legal Loan Document
              </p>
            </div>
          </div>
          <button onClick={onClose} className="p-2 hover:bg-slate-800 rounded-full transition-colors text-slate-400 hover:text-white">
            <X size={24} />
          </button>
        </div>

        <div id="loan-document" className="flex-1 overflow-y-auto p-12 bg-white text-slate-800 font-serif">
          {/* Header */}
          <div className="flex justify-between items-start mb-10 pb-8 border-b-2 border-slate-100">
            <div>
              <h1 className="text-4xl font-bold text-amber-600 brand-font mb-1 uppercase tracking-tight">PHOOL CHANDRA SARAF</h1>
              <p className="text-sm font-bold font-sans text-slate-500 tracking-widest uppercase">ASHISH JEWELLERS</p>
              <div className="mt-4 font-sans text-xs space-y-1 text-slate-600">
                <p>Koraon, Allahabad, Uttar Pradesh - 212306</p>
                <p>Tel: +91 9123456789 | Email: contact@ashishjewellers.com</p>
              </div>
            </div>
            <div className="text-right font-sans">
              <div className="inline-block px-4 py-2 bg-slate-100 rounded-lg mb-4">
                <h2 className="text-xl font-black uppercase tracking-widest text-slate-900">GIRVI RECEIPT</h2>
              </div>
              <div className="space-y-1 text-sm">
                <p><span className="text-slate-400">Loan ID:</span> <span className="font-black text-slate-900">#{loan.id}</span></p>
                <p><span className="text-slate-400">Date:</span> <span className="font-bold text-slate-900">{loan.startDate}</span></p>
              </div>
            </div>
          </div>

          <div className="grid grid-cols-2 gap-12 mb-10 font-sans">
            <div className="space-y-4">
              <h4 className="text-[10px] font-black text-slate-400 uppercase tracking-widest border-b pb-2">Borrower Details</h4>
              <div className="space-y-2">
                <p className="text-lg font-bold text-slate-900">{loan.customerName}</p>
                <p className="text-sm flex items-center gap-2 text-slate-600"><Smartphone size={14} className="text-slate-400" /> {loan.customerPhone}</p>
                <p className="text-sm flex items-center gap-2 text-slate-600"><Fingerprint size={14} className="text-slate-400" /> ID: {loan.govId}</p>
              </div>
            </div>
            <div className="space-y-4">
              <h4 className="text-[10px] font-black text-slate-400 uppercase tracking-widest border-b pb-2">Collateral Summary</h4>
              <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100">
                <p className="text-sm font-bold text-slate-900 mb-1">{loan.productDescription}</p>
                <div className="flex justify-between items-center text-xs">
                  <span className="text-slate-500">{loan.metalType} Purity: {loan.purity}</span>
                  <span className="font-black text-slate-900">Weight: {loan.weight}g</span>
                </div>
              </div>
            </div>
          </div>

          <div className="bg-slate-900 rounded-[2rem] p-8 text-white mb-10 font-sans shadow-xl">
            <div className="grid grid-cols-3 gap-8">
              <div className="text-center border-r border-slate-800">
                <p className="text-[10px] text-slate-500 uppercase font-black mb-2">Principal Amount</p>
                <p className="text-3xl font-black">₹{loan.principalAmount.toLocaleString()}</p>
              </div>
              <div className="text-center border-r border-slate-800">
                <p className="text-[10px] text-slate-500 uppercase font-black mb-2">Interest Rate</p>
                <p className="text-3xl font-black">{loan.interestRate}% <span className="text-xs text-slate-500">/mo</span></p>
              </div>
              <div className="text-center">
                <p className="text-[10px] text-slate-500 uppercase font-black mb-2">Final Maturity</p>
                <p className="text-3xl font-black">{loan.dueDate}</p>
              </div>
            </div>
          </div>

          {/* Terms and Conditions */}
          <div className="font-sans text-[10px] text-slate-500 space-y-4 border-t pt-8">
            <div className="flex items-center gap-2 text-slate-900 font-bold mb-2">
              <ShieldCheck size={14} className="text-amber-500" />
              TERMS & CONDITIONS OF LOAN (GIRVI)
            </div>
            <div className="grid grid-cols-2 gap-8 leading-relaxed">
              {/* Changed ul to ol to fix TypeScript error and properly support the 'start' attribute for ordered lists */}
              <ol className="list-decimal pl-4 space-y-1">
                <li>Loan is valid for 6-12 months. Interest is payable monthly.</li>
                <li>In case of non-payment of interest for 3 consecutive months, the firm reserves the right to auction the collateral.</li>
                <li>Borrower must produce this receipt for redemption or interest payment.</li>
              </ol>
              <ol className="list-decimal pl-4 space-y-1" start={4}>
                <li>Redemption takes 24-48 hours after full payment clears.</li>
                <li>Loss of receipt should be reported immediately to avoid misuse.</li>
                <li>All disputes are subject to Allahabad jurisdiction.</li>
              </ol>
            </div>
          </div>

          <div className="mt-16 flex justify-between items-end font-sans">
            <div className="text-center w-48">
              <div className="h-16 mb-2 border-b-2 border-slate-100 border-dashed"></div>
              <p className="font-black uppercase text-slate-600 tracking-widest text-[10px]">Customer's Signature</p>
              <p className="text-[8px] text-slate-400 mt-1 italic">(LTI / Thumb impression)</p>
            </div>
            <div className="text-center w-48">
              <div className="h-16 mb-2 border-b-2 border-slate-100 border-dashed"></div>
              <p className="font-black uppercase text-slate-600 tracking-widest text-[10px]">Authorised Signatory</p>
              <p className="text-[8px] text-slate-400 mt-1 italic">For PHOOL CHANDRA SARAF</p>
            </div>
          </div>
        </div>

        <div className="p-8 bg-slate-900 flex justify-end gap-4">
          <button onClick={onClose} className="px-8 py-3 rounded-2xl text-slate-400 hover:text-white font-bold transition-all hover:bg-slate-800">Discard</button>
          <button className="px-10 py-3 bg-amber-500 text-slate-900 rounded-2xl font-black flex items-center gap-3 hover:bg-amber-400 transition-all shadow-2xl shadow-amber-500/20 active:scale-95">
            <Printer size={20} /> Print Agreement
          </button>
        </div>
      </div>
    </div>
  );
};

export const Loans: React.FC = () => {
  const [loans, setLoans] = useState<Loan[]>(MOCK_LOANS);
  const [showForm, setShowForm] = useState(false);
  const [showPreview, setShowPreview] = useState(false);
  const [currentLoan, setCurrentLoan] = useState<Loan | null>(null);
  
  const [newLoan, setNewLoan] = useState<Partial<Loan>>({
    metalType: 'GOLD',
    purity: '22K',
    interestRate: 1.5,
    status: 'ACTIVE',
    startDate: new Date().toISOString().split('T')[0],
  });

  const handleAddLoan = (e: React.FormEvent) => {
    e.preventDefault();
    const loan: Loan = {
      id: `L-${Math.floor(Math.random() * 9000) + 1000}`,
      customerName: newLoan.customerName || '',
      customerPhone: newLoan.customerPhone || '',
      govId: newLoan.govId || '',
      metalType: newLoan.metalType as 'GOLD' | 'SILVER',
      productDescription: newLoan.productDescription || '',
      weight: newLoan.weight || 0,
      purity: newLoan.purity || '',
      principalAmount: newLoan.principalAmount || 0,
      interestRate: newLoan.interestRate || 0,
      startDate: newLoan.startDate || '',
      dueDate: new Date(new Date(newLoan.startDate!).setMonth(new Date(newLoan.startDate!).getMonth() + 6)).toISOString().split('T')[0],
      repayments: [],
      status: 'ACTIVE',
    };
    setLoans([loan, ...loans]);
    setCurrentLoan(loan);
    setShowForm(false);
    setShowPreview(true);
    setNewLoan({ metalType: 'GOLD', purity: '22K', interestRate: 1.5, startDate: new Date().toISOString().split('T')[0] });
  };

  const handlePrint = (loan: Loan) => {
    setCurrentLoan(loan);
    setShowPreview(true);
  };

  return (
    <div className="space-y-8 animate-in fade-in duration-700">
      <div className="flex justify-between items-center">
        <div>
          <h1 className="text-3xl font-black text-slate-900 flex items-center gap-3">
            <HandCoins className="text-amber-500" /> Girvi Management
          </h1>
          <p className="text-slate-500 italic">Manage gold and silver collateral loans with precision</p>
        </div>
        <button 
          onClick={() => setShowForm(!showForm)}
          className="bg-slate-900 text-white px-6 py-3 rounded-2xl font-black flex items-center gap-2 hover:bg-slate-800 transition-all shadow-xl shadow-slate-200"
        >
          {showForm ? <X size={20} /> : <Plus size={20} />} 
          {showForm ? 'Cancel' : 'New Loan Entry'}
        </button>
      </div>

      {showForm && (
        <div className="bg-white p-8 rounded-[2.5rem] shadow-2xl border border-amber-100 animate-in slide-in-from-top-4 duration-500">
          <h3 className="text-xs font-black text-amber-600 uppercase tracking-widest mb-8 border-b border-amber-50 pb-4">
            New Loan Document Form
          </h3>
          <form onSubmit={handleAddLoan} className="grid grid-cols-1 md:grid-cols-3 gap-6">
            <div className="space-y-2">
              <label className="text-[10px] font-black text-slate-400 uppercase flex items-center gap-1">
                <User size={12} /> Customer Name
              </label>
              <input 
                required
                className="w-full p-3.5 bg-slate-50 border-2 border-transparent rounded-2xl font-bold text-slate-900 focus:border-amber-500 transition-all outline-none"
                placeholder="Full Name"
                value={newLoan.customerName || ''}
                onChange={e => setNewLoan({...newLoan, customerName: e.target.value})}
              />
            </div>
            <div className="space-y-2">
              <label className="text-[10px] font-black text-slate-400 uppercase flex items-center gap-1">
                <Smartphone size={12} /> Phone Number
              </label>
              <input 
                required
                className="w-full p-3.5 bg-slate-50 border-2 border-transparent rounded-2xl font-bold text-slate-900 focus:border-amber-500 transition-all outline-none"
                placeholder="Mobile Number"
                value={newLoan.customerPhone || ''}
                onChange={e => setNewLoan({...newLoan, customerPhone: e.target.value})}
              />
            </div>
            <div className="space-y-2">
              <label className="text-[10px] font-black text-slate-400 uppercase flex items-center gap-1">
                <Fingerprint size={12} /> Gov ID (Aadhar/PAN)
              </label>
              <input 
                required
                className="w-full p-3.5 bg-slate-50 border-2 border-transparent rounded-2xl font-bold text-slate-900 focus:border-amber-500 transition-all outline-none"
                placeholder="ID Number"
                value={newLoan.govId || ''}
                onChange={e => setNewLoan({...newLoan, govId: e.target.value})}
              />
            </div>
            
            <div className="md:col-span-2 space-y-2">
              <label className="text-[10px] font-black text-slate-400 uppercase flex items-center gap-1">
                <Gem size={12} /> Product Description
              </label>
              <input 
                required
                className="w-full p-3.5 bg-slate-50 border-2 border-transparent rounded-2xl font-bold text-slate-900 focus:border-amber-500 transition-all outline-none"
                placeholder="e.g. Heavy Gold Bangles (2 pairs)"
                value={newLoan.productDescription || ''}
                onChange={e => setNewLoan({...newLoan, productDescription: e.target.value})}
              />
            </div>

            <div className="space-y-2">
              <label className="text-[10px] font-black text-slate-400 uppercase flex items-center gap-1">
                Metal Type
              </label>
              <select 
                className="w-full p-3.5 bg-slate-50 border-2 border-transparent rounded-2xl font-bold text-slate-900 focus:border-amber-500 transition-all outline-none"
                value={newLoan.metalType}
                onChange={e => setNewLoan({...newLoan, metalType: e.target.value as any, purity: e.target.value === 'GOLD' ? '22K' : '925'})}
              >
                <option value="GOLD">Gold</option>
                <option value="SILVER">Silver</option>
              </select>
            </div>

            <div className="grid grid-cols-2 gap-4">
              <div className="space-y-2">
                <label className="text-[10px] font-black text-slate-400 uppercase flex items-center gap-1">
                  <Scale size={12} /> Weight (g)
                </label>
                <input 
                  type="number"
                  required
                  className="w-full p-3.5 bg-slate-50 border-2 border-transparent rounded-2xl font-bold text-slate-900 focus:border-amber-500 transition-all outline-none"
                  value={newLoan.weight || ''}
                  onChange={e => setNewLoan({...newLoan, weight: parseFloat(e.target.value)})}
                />
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-black text-slate-400 uppercase">Purity</label>
                <input 
                  className="w-full p-3.5 bg-slate-50 border-2 border-transparent rounded-2xl font-bold text-slate-900 focus:border-amber-500 transition-all outline-none"
                  value={newLoan.purity || ''}
                  onChange={e => setNewLoan({...newLoan, purity: e.target.value})}
                />
              </div>
            </div>

            <div className="space-y-2">
              <label className="text-[10px] font-black text-slate-400 uppercase flex items-center gap-1">
                <HandCoins size={12} /> Principal Amount
              </label>
              <input 
                type="number"
                required
                className="w-full p-3.5 bg-slate-50 border-2 border-transparent rounded-2xl font-black text-amber-600 focus:border-amber-500 transition-all outline-none"
                value={newLoan.principalAmount || ''}
                onChange={e => setNewLoan({...newLoan, principalAmount: parseFloat(e.target.value)})}
              />
            </div>

            <div className="grid grid-cols-2 gap-4">
              <div className="space-y-2">
                <label className="text-[10px] font-black text-slate-400 uppercase flex items-center gap-1">
                  <Percent size={12} /> Int. Rate (%)
                </label>
                <input 
                  type="number"
                  step="0.1"
                  required
                  className="w-full p-3.5 bg-slate-50 border-2 border-transparent rounded-2xl font-bold text-slate-900 focus:border-amber-500 transition-all outline-none"
                  value={newLoan.interestRate || ''}
                  onChange={e => setNewLoan({...newLoan, interestRate: parseFloat(e.target.value)})}
                />
              </div>
              <div className="space-y-2">
                <label className="text-[10px] font-black text-slate-400 uppercase flex items-center gap-1">
                  <Calendar size={12} /> Loan Date
                </label>
                <input 
                  type="date"
                  required
                  className="w-full p-3.5 bg-slate-50 border-2 border-transparent rounded-2xl font-bold text-slate-900 focus:border-amber-500 transition-all outline-none"
                  value={newLoan.startDate || ''}
                  onChange={e => setNewLoan({...newLoan, startDate: e.target.value})}
                />
              </div>
            </div>

            <div className="md:col-span-3 pt-6">
              <button 
                type="submit"
                className="w-full bg-amber-500 text-slate-900 py-4 rounded-2xl font-black uppercase tracking-widest hover:bg-amber-400 transition-all shadow-xl shadow-amber-500/20 flex items-center justify-center gap-3"
              >
                <Save size={20} /> Sync & Generate Document
              </button>
            </div>
          </form>
        </div>
      )}

      {/* Loan List */}
      <div className="bg-white rounded-[2.5rem] shadow-xl border border-slate-100 overflow-hidden">
        <div className="p-8 border-b flex justify-between items-center">
          <h3 className="font-black text-slate-900 text-xl tracking-tight">Active Portfolio</h3>
          <div className="flex gap-4">
            <div className="relative">
              <Search className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400" size={18} />
              <input 
                placeholder="Search by Name/ID..." 
                className="pl-12 pr-6 py-2.5 bg-slate-50 border border-slate-200 rounded-xl text-sm font-bold text-slate-900 outline-none focus:ring-2 focus:ring-amber-500"
              />
            </div>
            <button className="p-2.5 bg-slate-50 text-slate-400 rounded-xl hover:text-slate-900 transition-colors">
              <Download size={20} />
            </button>
          </div>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left">
            <thead className="bg-slate-50">
              <tr className="text-[10px] font-black text-slate-400 uppercase tracking-widest">
                <th className="px-8 py-5">Loan ID</th>
                <th className="px-8 py-5">Customer & ID</th>
                <th className="px-8 py-5">Metal & Weight</th>
                <th className="px-8 py-5">Amount</th>
                <th className="px-8 py-5">Int. %</th>
                <th className="px-8 py-5">Due Date</th>
                <th className="px-8 py-5">Status</th>
                <th className="px-8 py-5">Actions</th>
              </tr>
            </thead>
            <tbody className="divide-y text-sm">
              {loans.map(loan => (
                <tr key={loan.id} className="hover:bg-slate-50 transition-colors">
                  <td className="px-8 py-6">
                    <span className="font-black text-slate-900">#{loan.id}</span>
                  </td>
                  <td className="px-8 py-6">
                    <p className="font-bold text-slate-900">{loan.customerName}</p>
                    <p className="text-xs text-slate-400">{loan.govId}</p>
                  </td>
                  <td className="px-8 py-6">
                    <p className="font-bold flex items-center gap-1">
                      <div className={`w-2 h-2 rounded-full ${loan.metalType === 'GOLD' ? 'bg-amber-500' : 'bg-slate-400'}`} />
                      {loan.weight}g ({loan.purity})
                    </p>
                    <p className="text-[10px] text-slate-500 truncate w-32">{loan.productDescription}</p>
                  </td>
                  <td className="px-8 py-6">
                    <span className="font-black text-slate-900">₹{loan.principalAmount.toLocaleString()}</span>
                  </td>
                  <td className="px-8 py-6 font-bold text-slate-500">{loan.interestRate}%</td>
                  <td className="px-8 py-6">
                    <p className="font-bold text-slate-900">{loan.dueDate}</p>
                    <p className="text-[10px] text-slate-400">Issued: {loan.startDate}</p>
                  </td>
                  <td className="px-8 py-6">
                    <span className={`px-3 py-1 rounded-full text-[10px] font-black uppercase tracking-widest ${
                      loan.status === 'ACTIVE' ? 'bg-green-100 text-green-700' :
                      loan.status === 'OVERDUE' ? 'bg-rose-100 text-rose-700' :
                      'bg-slate-100 text-slate-700'
                    }`}>
                      {loan.status}
                    </span>
                  </td>
                  <td className="px-8 py-6">
                    <div className="flex gap-2">
                      <button 
                        onClick={() => handlePrint(loan)}
                        className="p-2 text-slate-400 hover:text-amber-500 transition-colors"
                        title="Print Agreement"
                      >
                        <Printer size={18} />
                      </button>
                      <button className="p-2 text-slate-400 hover:text-rose-500 transition-colors">
                        <Trash2 size={18} />
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      <LoanDocumentModal 
        show={showPreview} 
        onClose={() => setShowPreview(false)} 
        loan={currentLoan}
      />
    </div>
  );
};
