
import { GoogleGenAI, Type } from "@google/genai";

const SYSTEM_INSTRUCTION = `
You are an AI assistant for an Indian jewellery shop billing and loan management system.
Role: Experienced jewellery accountant (munim) and business advisor.
Expertise: Gold purity, weight, making charges, GST compliance (3% rule), gold loans, interest, repayment behavior.
Style: Professional, practical, business-focused, Indian retail context.
Rules: Analyze data carefully, never invent numbers, output ONLY valid JSON.
Special Knowledge: In India, jewelry has 3% GST on (Gold Value + Making Charges). HSN code for gold jewelry is 7113.
`;

export class MunimAIService {
  private ai: GoogleGenAI;

  constructor() {
    this.ai = new GoogleGenAI({ apiKey: process.env.API_KEY });
  }

  private async generate(prompt: string, model: string = 'gemini-3-flash-preview', tools?: any[]): Promise<any> {
    try {
      const response = await this.ai.models.generateContent({
        model: model,
        contents: prompt,
        config: {
          systemInstruction: SYSTEM_INSTRUCTION,
          responseMimeType: "application/json",
          tools: tools,
        },
      });

      const text = response.text || '';
      const jsonStr = text.replace(/```json|```/g, '').trim();
      return JSON.parse(jsonStr);
    } catch (error) {
      console.error("Gemini Error:", error);
      throw error;
    }
  }

  async getLiveMetalRates() {
    const prompt = `Search for the latest gold and silver rates in India today from Moneycontrol (https://www.moneycontrol.com/news/gold-rates-today). 
    Provide the current market price for:
    1. Gold 24K (per 1 gram)
    2. Gold 22K (per 1 gram)
    3. Silver (per 1 kg)
    Return the data in the following JSON format:
    {
      "gold24k": number,
      "gold22k": number,
      "silver": number,
      "source": "Moneycontrol",
      "timestamp": "ISO Date String"
    }`;
    
    // Use Pro model for search grounding
    return this.generate(prompt, 'gemini-3-pro-preview', [{ googleSearch: {} }]);
  }

  async auditGstCompliance(invoices: any[]) {
    const prompt = `Perform a GST audit on these jewelry invoices. 
    Check if GST is correctly calculated at 3%. 
    Check for missing HSN codes (7113). 
    Check for B2B invoices without GSTIN.
    Invoices: ${JSON.stringify(invoices)}`;
    return this.generate(prompt);
  }

  async analyzeCustomers(data: any) {
    const prompt = `Analyze the following customer and transaction data.\nData: ${JSON.stringify(data)}\nTasks:\n1. Identify top 10 high-value customers.\n2. Identify customers with delayed payments.\n3. Classify into SAFE, WATCH, RISKY.`;
    return this.generate(prompt);
  }

  async analyzeLoanRisk(loans: any) {
    const prompt = `Analyze jewellery loan data and repayment history. Identify default risks and suggest actions.\nData: ${JSON.stringify(loans)}`;
    return this.generate(prompt);
  }

  async getBusinessSummary(monthlyData: any) {
    const prompt = `Summarize monthly sales, loan status, and provide 3 practical business suggestions.\nData: ${JSON.stringify(monthlyData)}`;
    return this.generate(prompt);
  }

  async queryToFilter(query: string, schema: string) {
    const prompt = `Convert the user query into a structured database filter.\nQuery: ${query}\nSchema: ${schema}`;
    return this.generate(prompt);
  }

  async getFestivalAdvice(historicalData: any) {
    const prompt = `Analyze past sales and loan trends to advise stock planning for upcoming festivals.\nData: ${JSON.stringify(historicalData)}`;
    return this.generate(prompt);
  }
}

export const munimAI = new MunimAIService();
