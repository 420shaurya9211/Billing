
export interface Customer {
  id: string;
  name: string;
  phone: string;
  address: string;
  gstin?: string;
  totalPurchases: number;
  activeLoans: number;
  loyaltyPoints: number;
  joinDate: string;
}

export interface InvoiceItem {
  id: string;
  description: string;
  metal: 'GOLD' | 'SILVER';
  weight: number;
  purity: string;
  ratePerGram: number;
  makingCharges: number;
  makingChargeType: 'FIXED' | 'PERCENTAGE';
  discount: number;
  taxableValue: number;
  hsnCode?: string;
}

export interface MetalReturnItem {
  id: string;
  description: string;
  weight: number;
  ratePerGram: number;
  purity: string;
  totalValue: number;
}

export interface Invoice {
  id: string;
  customerId: string;
  date: string;
  billType?: 'PAKKA' | 'KACHA';
  items: InvoiceItem[];
  returns?: MetalReturnItem[];
  subTotal: number;
  gstAmount: number;
  totalAmount: number;
  status: 'PAID' | 'PENDING';
  cgstRate?: number;
  sgstRate?: number;
  igstRate?: number;
}

export interface Loan {
  id: string;
  customerName: string;
  customerPhone: string;
  govId: string;
  metalType: 'GOLD' | 'SILVER';
  productDescription: string;
  weight: number;
  purity: string;
  principalAmount: number;
  interestRate: number; // monthly %
  startDate: string;
  dueDate: string;
  repayments: any[];
  status: 'ACTIVE' | 'CLOSED' | 'OVERDUE';
}
