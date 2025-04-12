import * as React from "react";

// Define the FinancialData interface
export interface FinancialData {
  filingInfo: {
    fiscalYear: number;
  };
  financialStatements: {
    incomeStatement: {
      revenue: number;
      costOfRevenue: number;
      grossProfit: number;
      sgAndA: number;
      rAndD: number;
      operatingExpenses: number;
      operatingIncome: number;
      interestExpense: number;
      incomeTaxExpense: number;
      netIncome: number;
      depreciationAmortization: number;
    };
    balanceSheet: {
      cashAndEquivalents: number;
      shortTermInvestments: number;
      accountsReceivable: number;
      inventory: number;
      totalCurrentAssets: number;
      propertyPlantEquipment: number;
      goodwill: number;
      intangibleAssets: number;
      totalAssets: number;
      accountsPayable: number;
      shortTermDebt: number;
      totalCurrentLiabilities: number;
      longTermDebt: number;
      totalLiabilities: number;
      totalEquity: number;
    };
    cashFlowStatement: {
      netIncome: number;
      depreciationAmortization: number;
      stockBasedCompensation: number;
      operatingCashFlow: number;
      capitalExpenditures: number;
      acquisitions: number;
      investingCashFlow: number;
      debtIssuance: number;
      debtRepayment: number;
      stockRepurchase: number;
      dividendsPaid: number;
      financingCashFlow: number;
      freeCashFlow: number;
    };
  };
  companyProfile: {
    name: string;
    ticker: string;
  };
}

// Define the context interface
interface FinancialDataContextType {
  financialData: FinancialData | null;
  setFinancialData: (data: FinancialData) => void;
  isLoading: boolean;
  setIsLoading: (loading: boolean) => void;
  status: {
    message: string;
    type: "error" | "success" | "info" | null;
  };
  setStatus: (status: { message: string; type: "error" | "success" | "info" | null }) => void;
}

// Create the context with default values
export const FinancialDataContext = React.createContext<FinancialDataContextType>({
  financialData: null,
  setFinancialData: () => {},
  isLoading: false,
  setIsLoading: () => {},
  status: { message: "", type: null },
  setStatus: () => {},
});

// Create a provider component
export const FinancialDataProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [financialData, setFinancialData] = React.useState<FinancialData | null>(null);
  const [isLoading, setIsLoading] = React.useState<boolean>(false);
  const [status, setStatus] = React.useState<{
    message: string;
    type: "error" | "success" | "info" | null;
  }>({ message: "", type: null });

  return (
    <FinancialDataContext.Provider
      value={{
        financialData,
        setFinancialData,
        isLoading,
        setIsLoading,
        status,
        setStatus,
      }}
    >
      {children}
    </FinancialDataContext.Provider>
  );
};

// Custom hook to use the financial data context
export const useFinancialData = () => React.useContext(FinancialDataContext);
