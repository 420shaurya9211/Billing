
import React from 'react';
import { HashRouter as Router, Routes, Route } from 'react-router-dom';
import { Layout } from './components/Layout';
import { Dashboard } from './pages/Dashboard';
import { MunimAI } from './pages/MunimAI';
import { Billing } from './pages/Billing';
import { GstReports } from './pages/GstReports';
import { Loans } from './pages/Loans';

const Placeholder = ({ title }: { title: string }) => (
  <div className="bg-white p-8 rounded-2xl shadow-sm border border-slate-100">
    <h1 className="text-2xl font-bold mb-4">{title}</h1>
    <p className="text-slate-500">This feature is coming soon in the MunimAI beta.</p>
  </div>
);

const App: React.FC = () => {
  return (
    <Router>
      <Layout>
        <Routes>
          <Route path="/" element={<Dashboard />} />
          <Route path="/customers" element={<Placeholder title="Customer Management" />} />
          <Route path="/billing" element={<Billing />} />
          <Route path="/gst" element={<GstReports />} />
          <Route path="/loans" element={<Loans />} />
          <Route path="/ai" element={<MunimAI />} />
          <Route path="/settings" element={<Placeholder title="Settings" />} />
        </Routes>
      </Layout>
    </Router>
  );
};

export default App;
