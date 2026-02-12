
import { Customer, Invoice, Loan } from './types';

export const MOCK_CUSTOMERS: Customer[] = [
  { id: 'C001', name: 'Rajesh Sharma', phone: '9876543210', address: 'Bandra West, Mumbai', totalPurchases: 450000, activeLoans: 1, loyaltyPoints: 450, joinDate: '2023-01-15' },
  { id: 'C002', name: 'Priya Patel', phone: '9820012345', address: 'Navrangpura, Ahmedabad', totalPurchases: 1200000, activeLoans: 0, loyaltyPoints: 1200, joinDate: '2022-11-20' },
  { id: 'C003', name: 'Amit Verma', phone: '9988776655', address: 'Karol Bagh, Delhi', totalPurchases: 50000, activeLoans: 2, loyaltyPoints: 50, joinDate: '2024-02-10' },
];

export const MOCK_INVOICES: Invoice[] = [
  {
    id: 'INV-101',
    customerId: 'C002',
    date: '2024-05-01',
    items: [{ 
      id: 'I1', 
      description: 'Gold Bangle 22K', 
      hsnCode: '7113',
      weight: 24.5, 
      purity: '22K', 
      metal: 'GOLD',
      ratePerGram: 7200, 
      makingCharges: 200, 
      makingChargeType: 'FIXED', 
      discount: 1500,
      taxableValue: 175100
    }],
    subTotal: 175100,
    gstAmount: 5253,
    totalAmount: 180353,
    status: 'PAID',
    cgstRate: 1.5,
    sgstRate: 1.5,
    igstRate: 0
  }
];

export const MOCK_LOANS: Loan[] = [
  {
    id: 'L-501',
    customerName: 'Amit Verma',
    customerPhone: '9988776655',
    govId: 'ABCDE1234F',
    metalType: 'GOLD',
    productDescription: 'Gold Chain (Man)',
    weight: 15.0,
    purity: '22K',
    principalAmount: 60000,
    interestRate: 1.5,
    startDate: '2024-01-10',
    dueDate: '2024-07-10',
    repayments: [],
    status: 'OVERDUE'
  },
  {
    id: 'L-502',
    customerName: 'Rajesh Sharma',
    customerPhone: '9876543210',
    govId: 'FGHIJ5678K',
    metalType: 'GOLD',
    productDescription: 'Gold Ring',
    weight: 10.0,
    purity: '22K',
    principalAmount: 40000,
    interestRate: 1.2,
    startDate: '2024-04-15',
    dueDate: '2024-10-15',
    repayments: [],
    status: 'ACTIVE'
  }
];
