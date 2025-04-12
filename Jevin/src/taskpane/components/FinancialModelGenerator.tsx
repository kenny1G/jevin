import * as React from "react";
import { Button, MessageBar } from "@fluentui/react-components";

/* global Excel, console */

interface FinancialData {
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

export const FinancialModelGenerator: React.FC = () => {
  const [status, setStatus] = React.useState<{
    message: string;
    type: "error" | "success" | "info" | null;
  }>({ message: "", type: null });

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setStatus({ message: "Reading file...", type: "info" });

    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const jsonData: FinancialData = JSON.parse(e.target?.result as string);
        console.log("Parsed JSON data:", jsonData);
        setStatus({ message: "Generating Excel model...", type: "info" });
        await generateFinancialModel(jsonData);
      } catch (error) {
        console.error("Error parsing JSON:", error);
        setStatus({
          message: "Error parsing JSON file. Please ensure it's in the correct format.",
          type: "error"
        });
      }
    };

    reader.onerror = () => {
      setStatus({
        message: "Error reading file. Please try again.",
        type: "error"
      });
    };

    reader.readAsText(file);
  };

  const generateFinancialModel = async (data: FinancialData) => {
    try {
      await Excel.run(async (context) => {
        console.log("Starting Excel model generation...");

        // Try to get existing worksheet
        const sheetName = "3-Statement Model";
        let sheet: Excel.Worksheet;
        
        try {
          sheet = context.workbook.worksheets.getItem(sheetName);
          sheet.delete();
          await context.sync();
        } catch (e) {
          console.log("No existing worksheet found");
        }

        // Create new worksheet
        sheet = context.workbook.worksheets.add(sheetName);
        sheet.activate();

        // Format headers - Only show current year and 2 forecast years
        const headerRange = sheet.getRange("A1:D1");  // Changed to 4 columns: Label, FY24, FY25, FY26
        const headerValues = [["", "FY24", "FY25", "FY26"]];
        headerRange.values = headerValues;
        headerRange.format.font.bold = true;
        headerRange.format.font.color = "#0066CC";  // Blue headers

        // Set column widths
        sheet.getRange("A:A").format.columnWidth = 150;
        sheet.getRange("B:D").format.columnWidth = 85;  // Changed to only 3 data columns

        let currentRow = 2;

        // Calculate growth rates and ratios from FY24 data
        const revenue = data.financialStatements.incomeStatement.revenue;
        const costOfRevenue = data.financialStatements.incomeStatement.costOfRevenue;
        const operatingExpenses = data.financialStatements.incomeStatement.operatingExpenses;
        const interestExpense = data.financialStatements.incomeStatement.interestExpense;
        const incomeTaxExpense = data.financialStatements.incomeStatement.incomeTaxExpense;

        // Calculate key ratios from latest year
        const costRatio = costOfRevenue / revenue;
        const opexRatio = operatingExpenses / revenue;
        const taxRate = incomeTaxExpense / (data.financialStatements.incomeStatement.operatingIncome - interestExpense);
        
        // Assume 5% revenue growth for projections
        const growthRate = 0.05;
        
        // Project FY25 and FY26
        const fy25Revenue = revenue * (1 + growthRate);
        const fy26Revenue = fy25Revenue * (1 + growthRate);
        
        const fy25CostOfRevenue = -(fy25Revenue * costRatio);
        const fy26CostOfRevenue = -(fy26Revenue * costRatio);
        
        const fy25OpEx = -(fy25Revenue * opexRatio);
        const fy26OpEx = -(fy26Revenue * opexRatio);
        
        const fy25OpIncome = fy25Revenue + fy25CostOfRevenue + fy25OpEx;
        const fy26OpIncome = fy26Revenue + fy26CostOfRevenue + fy26OpEx;
        
        // Assume interest expense stays constant
        const fy25PreTaxIncome = fy25OpIncome - interestExpense;
        const fy26PreTaxIncome = fy26OpIncome - interestExpense;
        
        const fy25TaxExpense = -(fy25PreTaxIncome * taxRate);
        const fy26TaxExpense = -(fy26PreTaxIncome * taxRate);
        
        const fy25NetIncome = fy25PreTaxIncome + fy25TaxExpense;
        const fy26NetIncome = fy26PreTaxIncome + fy26TaxExpense;

        // Balance Sheet Projections
        const arRatio = data.financialStatements.balanceSheet.accountsReceivable / revenue;
        const inventoryRatio = data.financialStatements.balanceSheet.inventory / revenue;
        const apRatio = data.financialStatements.balanceSheet.accountsPayable / revenue;
        
        // Project balance sheet items
        const fy25AR = fy25Revenue * arRatio;
        const fy26AR = fy26Revenue * arRatio;
        
        const fy25Inventory = fy25Revenue * inventoryRatio;
        const fy26Inventory = fy26Revenue * inventoryRatio;
        
        const fy25AP = fy25Revenue * apRatio;
        const fy26AP = fy26Revenue * apRatio;

        // Cash Flow Projections
        const capexRatio = Math.abs(data.financialStatements.cashFlowStatement.capitalExpenditures) / revenue;
        
        const fy25Capex = -(fy25Revenue * capexRatio);
        const fy26Capex = -(fy26Revenue * capexRatio);
        
        const fy25DA = data.financialStatements.incomeStatement.depreciationAmortization * 1.05;  // Assume 5% growth in D&A
        const fy26DA = fy25DA * 1.05;

        const fy25OpCF = fy25NetIncome + fy25DA;  // Simplified operating cash flow
        const fy26OpCF = fy26NetIncome + fy26DA;
        
        const fy25FCF = fy25OpCF + fy25Capex;  // Simplified free cash flow
        const fy26FCF = fy26OpCF + fy26Capex;

        // Income Statement Header with separator
        const headerRangeIS = sheet.getRange(`A${currentRow}:D${currentRow}`);  // Changed to 4 columns
        headerRangeIS.values = [["Income Statement:", "", "", ""]];
        headerRangeIS.format.font.bold = true;
        headerRangeIS.format.font.color = "#666666";
        headerRangeIS.format.borders.getItem('EdgeBottom').style = 'Continuous';
        headerRangeIS.format.borders.getItem('EdgeBottom').color = "#666666";
        currentRow++;

        // Update array format to use only 4 columns
        const incomeStatementItems = [
            ["Products:", data.financialStatements.incomeStatement.revenue || "", fy25Revenue, fy26Revenue],
            ["Services:", "", "", ""],
            ["Total Revenue:", revenue, fy25Revenue, fy26Revenue],
            ["Cost of Products:", costOfRevenue ? `(${costOfRevenue})` : "", fy25CostOfRevenue, fy26CostOfRevenue],
            ["Cost of Services:", "", "", ""],
            ["Operating Expenses:", operatingExpenses ? `(${operatingExpenses})` : "", fy25OpEx, fy26OpEx],
            ["Operating Income:", data.financialStatements.incomeStatement.operatingIncome || "", fy25OpIncome, fy26OpIncome],
            ["Other Income / (Expense):", "", "", ""],
            ["Non-Service Pension Expense:", "", "", ""],
            ["Interest Income / (Expense):", interestExpense ? `(${interestExpense})` : "", interestExpense, interestExpense],
            ["Pre-Tax Income:", "", fy25PreTaxIncome, fy26PreTaxIncome],
            ["Income Taxes:", incomeTaxExpense ? `(${incomeTaxExpense})` : "", fy25TaxExpense, fy26TaxExpense],
            ["Net Income:", data.financialStatements.incomeStatement.netIncome || "", fy25NetIncome, fy26NetIncome],
            ["(-) NCI Net Income:", "", "", ""],
            ["Net Income to Parent:", data.financialStatements.incomeStatement.netIncome || "", fy25NetIncome, fy26NetIncome]
        ];

        const incomeStatementRange = sheet.getRange(`A${currentRow}:D${currentRow + incomeStatementItems.length - 1}`);
        incomeStatementRange.values = incomeStatementItems;
        
        // Format numbers and colors
        const dataRange = sheet.getRange(`B${currentRow}:D${currentRow + incomeStatementItems.length - 1}`);
        dataRange.numberFormat = [["#,##0;(#,##0);-"]];
        dataRange.format.horizontalAlignment = "Right";
        
        // Color the FY24 values blue
        const fy24Range = sheet.getRange(`B${currentRow}:B${currentRow + incomeStatementItems.length - 1}`);
        fy24Range.format.font.color = "#0066CC";  // Blue for JSON values
        
        // Set colors for labels
        const labelRange = sheet.getRange(`A${currentRow}:A${currentRow + incomeStatementItems.length - 1}`);
        labelRange.format.font.color = "#000000";  // Black for regular text
        
        // Bold totals
        sheet.getRange(`A${currentRow + 2}`).format.font.bold = true;  // Total Revenue
        sheet.getRange(`A${currentRow + 6}`).format.font.bold = true;  // Operating Income
        sheet.getRange(`A${currentRow + 12}`).format.font.bold = true; // Net Income
        sheet.getRange(`A${currentRow + 14}`).format.font.bold = true; // Net Income to Parent

        currentRow += incomeStatementItems.length + 1;

        // Balance Sheet Header with separator
        const balanceSheetHeaderRange = sheet.getRange(`A${currentRow}:D${currentRow}`);
        balanceSheetHeaderRange.values = [["Balance Sheet:", "", "", ""]];
        balanceSheetHeaderRange.format.font.bold = true;
        balanceSheetHeaderRange.format.font.color = "#666666";
        balanceSheetHeaderRange.format.borders.getItem('EdgeBottom').style = 'Continuous';
        balanceSheetHeaderRange.format.borders.getItem('EdgeBottom').color = "#666666";
        currentRow++;

        // Update balance sheet items to use 4 columns
        const balanceSheetItems = [
            ["ASSETS:", "", "", ""],
            ["    Cash:", data.financialStatements.balanceSheet.cashAndEquivalents || "", "", ""],
            ["    Accounts Receivable:", data.financialStatements.balanceSheet.accountsReceivable || "", fy25AR, fy26AR],
            ["    Inventory & Other:", data.financialStatements.balanceSheet.inventory || "", fy25Inventory, fy26Inventory],
            ["    Net PP&E, Goodwill & Intangibles:", (data.financialStatements.balanceSheet.propertyPlantEquipment || 0) + (data.financialStatements.balanceSheet.goodwill || 0) + (data.financialStatements.balanceSheet.intangibleAssets || 0), "", ""],
            ["    Op. Lease Assets:", "", "", ""],
            ["    Other Assets:", "", "", ""],
            ["Total Assets:", data.financialStatements.balanceSheet.totalAssets || "", "", ""],
            ["LIABILITIES & EQUITY:", "", "", ""],
            ["    Accounts Payable:", data.financialStatements.balanceSheet.accountsPayable || "", fy25AP, fy26AP],
            ["    Accrued Liabilities:", "", "", ""],
            ["    Contract Liabilities:", "", "", ""],
            ["    Total Debt:", "", "", ""],
            ["    Op. Lease Liabilities:", "", "", ""],
            ["    Other Liabilities:", "", "", ""],
            ["Total Liabilities:", "", "", ""],
            ["Common Shareholders' Equity:", "", "", ""],
            ["Noncontrolling Interests:", "", "", ""],
            ["Total Equity:", "", "", ""],
            ["TOTAL LIABILITIES + EQUITY:", "", "", ""],
            ["Balance Check:", "", "", ""]
        ];

        const balanceSheetRange = sheet.getRange(`A${currentRow}:D${currentRow + balanceSheetItems.length - 1}`);
        balanceSheetRange.values = balanceSheetItems;

        // Format numbers and colors
        const balanceSheetDataRange = sheet.getRange(`B${currentRow}:D${currentRow + balanceSheetItems.length - 1}`);
        balanceSheetDataRange.numberFormat = [["#,##0;(#,##0);-"]];
        balanceSheetDataRange.format.horizontalAlignment = "Right";

        // Set colors for labels and headers
        const balanceSheetLabelRange = sheet.getRange(`A${currentRow}:A${currentRow + balanceSheetItems.length - 1}`);
        balanceSheetLabelRange.format.font.color = "#000000";  // Black for regular text
        
        // Format balance sheet headers and totals
        sheet.getRange(`A${currentRow}`).format.font.color = "#666666";  // ASSETS
        sheet.getRange(`A${currentRow}`).format.font.bold = true;
        sheet.getRange(`A${currentRow + 3}`).format.font.bold = true;  // Total Assets
        sheet.getRange(`A${currentRow + 4}`).format.font.color = "#666666";  // LIABILITIES & EQUITY
        sheet.getRange(`A${currentRow + 4}`).format.font.bold = true;
        sheet.getRange(`A${currentRow + 7}`).format.font.bold = true;  // Total Liabilities
        sheet.getRange(`A${currentRow + 8}`).format.font.bold = true;  // Total Equity
        sheet.getRange(`A${currentRow + 9}`).format.font.bold = true;  // TOTAL LIABILITIES + EQUITY
        sheet.getRange(`A${currentRow + 10}`).format.font.bold = true;  // Balance Check

        // Color the FY24 values blue
        const balanceSheetFY24Range = sheet.getRange(`B${currentRow}:B${currentRow + balanceSheetItems.length - 1}`);
        balanceSheetFY24Range.format.font.color = "#0066CC";

        currentRow += balanceSheetItems.length + 1;

        // Cash Flow Statement Header with separator
        const cashFlowHeaderRange = sheet.getRange(`A${currentRow}:D${currentRow}`);
        cashFlowHeaderRange.values = [["Cash Flow Statement:", "", "", ""]];
        cashFlowHeaderRange.format.font.bold = true;
        cashFlowHeaderRange.format.font.color = "#666666";
        cashFlowHeaderRange.format.borders.getItem('EdgeBottom').style = 'Continuous';
        cashFlowHeaderRange.format.borders.getItem('EdgeBottom').color = "#666666";
        currentRow++;

        // Update cash flow items to use 4 columns
        const cashFlowItems = [
            ["Operating Activities:", "", "", ""],
            ["    Net Income:", data.financialStatements.cashFlowStatement.netIncome || "", fy25NetIncome, fy26NetIncome],
            ["    Depreciation & Amortization:", data.financialStatements.cashFlowStatement.depreciationAmortization || "", fy25DA, fy26DA],
            ["    Stock Based Compensation:", "", "", ""],
            ["Operating Cash Flow:", "", "", ""],
            ["Investing Activities:", "", "", ""],
            ["    Capital Expenditures:", "", "", ""],
            ["    Acquisitions:", "", "", ""],
            ["Investing Cash Flow:", "", "", ""],
            ["Financing Activities:", "", "", ""],
            ["    Debt Issuance:", "", "", ""],
            ["    Debt Repayment:", "", "", ""],
            ["    Share Repurchases:", "", "", ""],
            ["    Dividends:", "", "", ""],
            ["Financing Cash Flow:", "", "", ""],
            ["Free Cash Flow:", "", "", ""]
        ];

        const cashFlowRange = sheet.getRange(`A${currentRow}:D${currentRow + cashFlowItems.length - 1}`);
        cashFlowRange.values = cashFlowItems;

        // Format numbers and colors
        const cashFlowDataRange = sheet.getRange(`B${currentRow}:D${currentRow + cashFlowItems.length - 1}`);
        cashFlowDataRange.numberFormat = [["#,##0;(#,##0);-"]];
        cashFlowDataRange.format.horizontalAlignment = "Right";

        // Set colors for labels and headers
        const cashFlowLabelRange = sheet.getRange(`A${currentRow}:A${currentRow + cashFlowItems.length - 1}`);
        cashFlowLabelRange.format.font.color = "#000000";  // Black for regular text
        
        // Format cash flow headers and totals
        sheet.getRange(`A${currentRow}`).format.font.color = "#666666";  // Operating Activities
        sheet.getRange(`A${currentRow}`).format.font.bold = true;
        sheet.getRange(`A${currentRow + 4}`).format.font.bold = true;  // Operating Cash Flow
        sheet.getRange(`A${currentRow + 5}`).format.font.color = "#666666";  // Investing Activities
        sheet.getRange(`A${currentRow + 5}`).format.font.bold = true;
        sheet.getRange(`A${currentRow + 8}`).format.font.bold = true;  // Investing Cash Flow
        sheet.getRange(`A${currentRow + 9}`).format.font.color = "#666666";  // Financing Activities
        sheet.getRange(`A${currentRow + 9}`).format.font.bold = true;
        sheet.getRange(`A${currentRow + 12}`).format.font.bold = true;  // Financing Cash Flow
        sheet.getRange(`A${currentRow + 13}`).format.font.bold = true;  // Free Cash Flow

        // Color the FY24 values blue
        const cashFlowFY24Range = sheet.getRange(`B${currentRow}:B${currentRow + cashFlowItems.length - 1}`);
        cashFlowFY24Range.format.font.color = "#0066CC";

        // Set white background for entire used range and remove vertical borders
        const usedRange = sheet.getUsedRange();
        usedRange.format.fill.color = "white";
        usedRange.format.borders.getItem('EdgeLeft').style = 'None';
        usedRange.format.borders.getItem('EdgeRight').style = 'None';
        usedRange.format.borders.getItem('InsideVertical').style = 'None';

        await context.sync();
        setStatus({ message: "Financial model generated successfully!", type: "success" });
      });
    } catch (error) {
      console.error("Error in Excel model generation:", error);
      setStatus({
        message: `Error generating Excel model: ${error.message}`,
        type: "error"
      });
    }
  };

  return (
    <div style={{ padding: "20px" }}>
      <input
        type="file"
        accept=".json"
        onChange={handleFileUpload}
        style={{ display: "none" }}
        id="financial-json-input"
      />
      <Button
        appearance="primary"
        onClick={() => document.getElementById("financial-json-input")?.click()}
      >
        Upload Financial JSON
      </Button>
      {status.type && (
        <div style={{ marginTop: "20px" }}>
          <MessageBar intent={status.type === "error" ? "error" : status.type === "success" ? "success" : "info"}>
            {status.message}
          </MessageBar>
        </div>
      )}
    </div>
  );
};

export default FinancialModelGenerator; 