import * as React from "react";
import { Button, MessageBar } from "@fluentui/react-components";
import { useFinancialData, FinancialData } from "../context/FinancialDataContext";

/* global Excel, console */

export const FinancialModelGenerator: React.FC = () => {
  // Use the financial data context
  const { financialData, setFinancialData, status, setStatus } = useFinancialData();

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setStatus({ message: "Reading file...", type: "info" });

    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const jsonData: FinancialData = JSON.parse(e.target?.result as string);
        console.log("Parsed JSON data:", jsonData);

        // Store the data in the context so other components can access it
        setFinancialData(jsonData);

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

        // Set column widths to prevent text cutoff
        sheet.getRange("A:A").format.columnWidth = 250;  // Wider first column for labels
        sheet.getRange("B:D").format.columnWidth = 100;  // Wider data columns

        // Add company information
        const companyNameRange = sheet.getRange("A1:D1");
        companyNameRange.values = [[`${data.companyProfile.name} (${data.companyProfile.ticker})`, "", "", ""]];
        companyNameRange.format.font.bold = true;
        companyNameRange.format.font.size = 14;

        // Add model drivers section
        const driversHeaderRange = sheet.getRange("A3:D3");
        driversHeaderRange.values = [["Model Drivers:", "FY24", "FY25", "FY26"]];
        driversHeaderRange.format.font.bold = true;
        driversHeaderRange.format.font.color = "#666666";
        driversHeaderRange.format.borders.getItem('EdgeBottom').style = 'Continuous';
        driversHeaderRange.format.borders.getItem('EdgeBottom').color = "#666666";

        // Format year headers in drivers section
        const driversYearHeaders = sheet.getRange("B3:D3");
        driversYearHeaders.format.font.color = "#0066CC";
        driversYearHeaders.format.horizontalAlignment = "Right";

        const driversItems = [
            ["Revenue Growth:", "", 0.05, 0.05],
            ["Cost of Revenue (% of Revenue):", (data.financialStatements.incomeStatement.costOfRevenue / data.financialStatements.incomeStatement.revenue),
             "=B5",
             "=B5"],
            ["Operating Expenses (% of Revenue):", (data.financialStatements.incomeStatement.operatingExpenses / data.financialStatements.incomeStatement.revenue),
             "=B6",
             "=B6"],
            ["D&A Growth:", "", 0.05, 0.05],
            ["Tax Rate:", (data.financialStatements.incomeStatement.incomeTaxExpense / (data.financialStatements.incomeStatement.operatingIncome - data.financialStatements.incomeStatement.interestExpense)),
             "=B8",
             "=B8"],
            ["Working Capital Assumptions:", "", "", ""],
            ["    AR (% of Revenue):", (data.financialStatements.balanceSheet.accountsReceivable / data.financialStatements.incomeStatement.revenue),
             "=B10",
             "=B10"],
            ["    Inventory (% of Revenue):", (data.financialStatements.balanceSheet.inventory / data.financialStatements.incomeStatement.revenue),
             "=B11",
             "=B11"],
            ["    AP (% of Revenue):", (data.financialStatements.balanceSheet.accountsPayable / data.financialStatements.incomeStatement.revenue),
             "=B12",
             "=B12"],
            ["Capex (% of Revenue):", Math.abs(data.financialStatements.cashFlowStatement.capitalExpenditures) / data.financialStatements.incomeStatement.revenue,
             "=B13",
             "=B13"]
        ];

        const driversRange = sheet.getRange(`A4:D${4 + driversItems.length - 1}`);
        driversRange.values = driversItems;

        // Format drivers section
        const driversDataRange = sheet.getRange(`B4:D${4 + driversItems.length - 1}`);
        driversDataRange.format.horizontalAlignment = "Right";
        driversDataRange.numberFormat = [["0.0%"]];  // Format as percentage with array syntax

        // Color the FY24 values blue
        const driversFY24Range = sheet.getRange(`B4:B${4 + driversItems.length - 1}`);
        driversFY24Range.format.font.color = "#0066CC";

        // Color the assumptions (FY25 and FY26) green
        const assumptionsRange = sheet.getRange(`C4:D${4 + driversItems.length - 1}`);
        assumptionsRange.format.font.color = "#008000";  // Green color for assumptions

        // Bold the section headers in drivers
        sheet.getRange("A9").format.font.bold = true;  // Working Capital Assumptions

        // Add spacing before financial statements
        let currentRow = 4 + driversItems.length + 2;

        // Format headers for financial statements
        const headerRange = sheet.getRange(`A${currentRow}:D${currentRow}`);
        const headerValues = [["", "FY24", "FY25", "FY26"]];
        headerRange.values = headerValues;
        headerRange.format.font.bold = true;
        headerRange.format.font.color = "#0066CC";

        // Right align the statement year headers
        const statementYearHeaders = sheet.getRange(`B${currentRow}:D${currentRow}`);
        statementYearHeaders.format.horizontalAlignment = "Right";

        currentRow += 2;

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
        const headerRangeIS = sheet.getRange(`A${currentRow}:D${currentRow}`);
        headerRangeIS.values = [["Income Statement:", "", "", ""]];
        headerRangeIS.format.font.bold = true;
        headerRangeIS.format.font.color = "#666666";
        headerRangeIS.format.borders.getItem('EdgeBottom').style = 'Continuous';
        headerRangeIS.format.borders.getItem('EdgeBottom').color = "#666666";
        currentRow++;

        // Update array format to use only 4 columns with Excel formulas
        const incomeStatementItems = [
            ["Revenue:", data.financialStatements.incomeStatement.revenue,
             `=B${currentRow}*(1+C4)`,
             `=C${currentRow}*(1+D4)`],
            ["Cost of Revenue:", data.financialStatements.incomeStatement.costOfRevenue,
             `=-C${currentRow}*C5`,
             `=-D${currentRow}*D5`],
            ["Gross Profit:", data.financialStatements.incomeStatement.grossProfit,
             `=C${currentRow}+C${currentRow + 1}`,
             `=D${currentRow}+D${currentRow + 1}`],
            ["Operating Expenses:", data.financialStatements.incomeStatement.operatingExpenses,
             `=-C${currentRow}*C6`,
             `=-D${currentRow}*D6`],
            ["    R&D:", data.financialStatements.incomeStatement.rAndD, "", ""],
            ["    SG&A:", data.financialStatements.incomeStatement.sgAndA, "", ""],
            ["Operating Income:", data.financialStatements.incomeStatement.operatingIncome,
             `=C${currentRow + 2}+C${currentRow + 3}`,
             `=D${currentRow + 2}+D${currentRow + 3}`],
            ["Interest Expense:", data.financialStatements.incomeStatement.interestExpense,
             `=B${currentRow + 7}`,
             `=B${currentRow + 7}`],
            ["Income Tax Expense:", data.financialStatements.incomeStatement.incomeTaxExpense,
             `=-(C${currentRow + 6}-C${currentRow + 7})*C8`,
             `=-(D${currentRow + 6}-D${currentRow + 7})*D8`],
            ["Net Income:", data.financialStatements.incomeStatement.netIncome,
             `=C${currentRow + 6}+C${currentRow + 7}+C${currentRow + 8}`,
             `=D${currentRow + 6}+D${currentRow + 7}+D${currentRow + 8}`]
        ];

        const incomeStatementRange = sheet.getRange(`A${currentRow}:D${currentRow + incomeStatementItems.length - 1}`);
        incomeStatementRange.values = incomeStatementItems;

        // Format numbers and colors
        const dataRange = sheet.getRange(`B${currentRow}:D${currentRow + incomeStatementItems.length - 1}`);
        dataRange.numberFormat = [["#,##0;(#,##0);-"]];  // Fix array syntax
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

        // Update balance sheet items with Excel formulas
        const balanceSheetItems = [
            ["ASSETS:", "", "", ""],
            ["Current Assets:", "", "", ""],
            ["    Cash and Equivalents:", data.financialStatements.balanceSheet.cashAndEquivalents || "", "", ""],
            ["    Short Term Investments:", data.financialStatements.balanceSheet.shortTermInvestments || "", "", ""],
            ["    Accounts Receivable:", data.financialStatements.balanceSheet.accountsReceivable || "",
             `=C${currentRow - incomeStatementItems.length + 1}*C10`,
             `=D${currentRow - incomeStatementItems.length + 1}*D10`],
            ["    Inventory:", data.financialStatements.balanceSheet.inventory || "",
             `=C${currentRow - incomeStatementItems.length + 1}*C11`,
             `=D${currentRow - incomeStatementItems.length + 1}*D11`],
            ["Total Current Assets:", data.financialStatements.balanceSheet.totalCurrentAssets || "", "", ""],
            ["Property, Plant & Equipment:", data.financialStatements.balanceSheet.propertyPlantEquipment || "", "", ""],
            ["Goodwill:", data.financialStatements.balanceSheet.goodwill || "", "", ""],
            ["Intangible Assets:", data.financialStatements.balanceSheet.intangibleAssets || "", "", ""],
            ["Total Assets:", data.financialStatements.balanceSheet.totalAssets || "", "", ""],
            ["LIABILITIES & EQUITY:", "", "", ""],
            ["Current Liabilities:", "", "", ""],
            ["    Accounts Payable:", data.financialStatements.balanceSheet.accountsPayable || "",
             `=C${currentRow - incomeStatementItems.length + 1}*C12`,
             `=D${currentRow - incomeStatementItems.length + 1}*D12`],
            ["    Short Term Debt:", data.financialStatements.balanceSheet.shortTermDebt || "", "", ""],
            ["Total Current Liabilities:", data.financialStatements.balanceSheet.totalCurrentLiabilities || "", "", ""],
            ["Long Term Debt:", data.financialStatements.balanceSheet.longTermDebt || "", "", ""],
            ["Total Liabilities:", data.financialStatements.balanceSheet.totalLiabilities || "", "", ""],
            ["Total Equity:", data.financialStatements.balanceSheet.totalEquity || "", "", ""],
            ["TOTAL LIABILITIES + EQUITY:", data.financialStatements.balanceSheet.totalAssets || "", "", ""]
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

        // Update cash flow items with Excel formulas
        const cashFlowItems = [
            ["Operating Activities:", "", "", ""],
            ["    Net Income:", data.financialStatements.cashFlowStatement.netIncome || "",
             `=C${currentRow - balanceSheetItems.length - incomeStatementItems.length + 10}`,
             `=D${currentRow - balanceSheetItems.length - incomeStatementItems.length + 10}`],
            ["    Depreciation & Amortization:", data.financialStatements.cashFlowStatement.depreciationAmortization || "",
             `=B${currentRow + 2}*(1+C7)`,
             `=C${currentRow + 2}*(1+D7)`],
            ["Operating Cash Flow:", data.financialStatements.cashFlowStatement.operatingCashFlow || "",
             `=C${currentRow + 2}+C${currentRow + 3}`,
             `=D${currentRow + 2}+D${currentRow + 3}`],
            ["Investing Activities:", "", "", ""],
            ["    Capital Expenditures:", data.financialStatements.cashFlowStatement.capitalExpenditures ? `(${data.financialStatements.cashFlowStatement.capitalExpenditures})` : "",
             `=-C${currentRow - balanceSheetItems.length - incomeStatementItems.length + 1}*C13`,
             `=-D${currentRow - balanceSheetItems.length - incomeStatementItems.length + 1}*D13`],
            ["Free Cash Flow:", data.financialStatements.cashFlowStatement.freeCashFlow || "",
             `=C${currentRow + 4}+C${currentRow + 6}`,
             `=D${currentRow + 4}+D${currentRow + 6}`]
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

  // Use effect to automatically generate model when financial data changes
  React.useEffect(() => {
    const generateModelFromContext = async () => {
      if (financialData) {
        setStatus({ message: "Generating Excel model from context data...", type: "info" });
        try {
          await generateFinancialModel(financialData);
        } catch (error) {
          console.error("Error generating model from context:", error);
          setStatus({
            message: `Error generating model: ${error instanceof Error ? error.message : "Unknown error"}`,
            type: "error"
          });
        }
      }
    };

    generateModelFromContext();
  }, [financialData]); // Run effect when financialData changes

  return (
    <div style={{ padding: "20px" }}>
      {/* <input
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
      </Button> */}
      {status.type && (
        <div style={{ marginTop: "20px" }}>
          <MessageBar intent={status.type === "error" ? "error" : status.type === "success" ? "success" : "info"}>
            {status.message}
          </MessageBar>
        </div>
      )}
      {financialData && (
        <div style={{ marginTop: "10px", fontSize: "14px", color: "#666" }}>
          Using financial data from {financialData.companyProfile?.name || "uploaded file"}
        </div>
      )}
    </div>
  );
};

export default FinancialModelGenerator;