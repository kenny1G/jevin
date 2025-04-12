import * as React from "react";
import { makeStyles, Label, Button, Spinner, MessageBar } from "@fluentui/react-components";
import { useConvex, useQuery } from "convex/react";
import { api } from "../../../convex/_generated/api";

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

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    padding: "16px",
  },
  input: {
    width: "100%",
  },
  button: {
    marginTop: "8px",
  },
  status: {
    marginTop: "8px",
  },
});

const CombinedFilingModelGenerator: React.FC = () => {
  const styles = useStyles();
  const convex = useConvex();
  const [file, setFile] = React.useState<File | null>(null);
  const [uploading, setUploading] = React.useState(false);
  const [status, setStatus] = React.useState<{
    message: string;
    type: "error" | "success" | "info" | null;
  }>({ message: "", type: null });
  const [extractionJobId, setExtractionJobId] = React.useState<string | null>(null);
  const [documentId, setDocumentId] = React.useState<string | null>(null);
  const [existingDocumentId, setExistingDocumentId] = React.useState<string | null>(null);
  const [pollingInterval, setPollingInterval] = React.useState<NodeJS.Timeout | null>(null);

  const convexDocument = useQuery(api.models.document.findDocumentByName, {
    name: file?.name,
  });

  // If we find an existing document, store its ID
  React.useEffect(() => {
    if (convexDocument) {
      setExistingDocumentId(convexDocument._id);
    }
  }, [convexDocument]);

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = event.target.files?.[0];
    if (selectedFile && selectedFile.type === "application/pdf") {
      setFile(selectedFile);
      setStatus({ message: "", type: null });
    }
  };

  const checkExtractionStatus = React.useCallback(
    async (jobId: string) => {
      try {
        // Get the Convex site URL
        const convexSiteUrl = "https://posh-panda-366.convex.site";

        // Fetch extraction result
        const response = await fetch(`${convexSiteUrl}/fetch-extraction-result?jobId=${jobId}`);
        const result = await response.json();

        if (response.ok && result.extractionResult) {
          const runs = result.extractionResult.runs;

          if (runs && runs.length > 0) {
            const latestRun = runs[0];

            if (latestRun.status === "completed") {
              // Clear the polling interval
              if (pollingInterval) {
                clearInterval(pollingInterval);
                setPollingInterval(null);
              }

              setStatus({
                message: "Extraction completed. Generating financial model...",
                type: "info",
              });

              // Process the extraction result
              if (latestRun.output && latestRun.output.data) {
                await generateFinancialModel(latestRun.output.data as FinancialData);
              } else {
                setStatus({
                  message: "Extraction completed but no data was found.",
                  type: "error",
                });
              }
            } else if (latestRun.status === "failed") {
              // Clear the polling interval
              if (pollingInterval) {
                clearInterval(pollingInterval);
                setPollingInterval(null);
              }

              setStatus({
                message: `Extraction failed: ${latestRun.error || "Unknown error"}`,
                type: "error",
              });
            } else {
              // Still in progress
              setStatus({ message: `Extraction in progress: ${latestRun.status}`, type: "info" });
            }
          }
        }
      } catch (error) {
        console.error("Error checking extraction status:", error);
        setStatus({
          message: `Error checking extraction status: ${error instanceof Error ? error.message : "Unknown error"}`,
          type: "error",
        });
      }
    },
    [pollingInterval]
  );

  // Cleanup polling interval on component unmount
  React.useEffect(() => {
    return () => {
      if (pollingInterval) {
        clearInterval(pollingInterval);
      }
    };
  }, [pollingInterval]);

  const fetchExtractionAndGenerateModel = async (docId: any) => {
    try {
      setStatus({ message: "Fetching extraction data...", type: "info" });

      // Get the document extract for this document
      const extractData = await convex.query(
        api.models.documentExtract.getDocumentExtractByDocumentId,
        {
          documentId: docId, // docId is already the correct Id<"documents"> type from convexDocument._id
        }
      );

      if (!extractData) {
        setStatus({ message: "No extraction data found for this document.", type: "error" });
        return;
      }

      // Get the extraction job ID
      const jobId = extractData.extractionJobId;

      // Get the Convex site URL
      const convexSiteUrl = "https://posh-panda-366.convex.site";

      // Fetch extraction result
      const response = await fetch(`${convexSiteUrl}/fetch-extraction-result?jobId=${jobId}`);
      const result = await response.json();
      setStatus({ message: "Got response", type: "info" });

      if (response.ok && result.extractionResult) {
        const latestRun = result.extractionResult.data;

        if (latestRun) {
          setStatus({
            message: "Extraction data found. Generating financial model...",
            type: "info",
          });

          // Process the extraction result
          await generateFinancialModel(latestRun as FinancialData);
        } else {
          setStatus({ message: "Extraction completed but no data was found.", type: "error" });
        }
      } else {
        setStatus({ message: "Failed to fetch extraction data.", type: "error" });
      }
    } catch (error) {
      console.error("Error fetching extraction data:", error);
      setStatus({
        message: `Error fetching extraction data: ${error instanceof Error ? error.message : "Unknown error"}`,
        type: "error",
      });
    }
  };

  const handleUploadAndGenerate = async () => {
    if (!file) return;

    // If we already have this document, use the existing document ID
    if (convexDocument) {
      fetchExtractionAndGenerateModel(convexDocument._id);
      return;
    }

    try {
      setUploading(true);
      setStatus({ message: "Uploading file...", type: "info" });

      // Create form data
      const formData = new FormData();
      formData.append("file", file);

      // Get the Convex site URL
      const convexSiteUrl = "https://posh-panda-366.convex.site";

      // Upload to our Convex proxy
      const response = await fetch(`${convexSiteUrl}/upload-to-llamaindex`, {
        method: "POST",
        body: formData,
      });

      const result = await response.json();

      if (response.ok) {
        setStatus({
          message: "File uploaded successfully! Waiting for extraction...",
          type: "info",
        });

        // Store the extraction job ID for polling
        if (result.extractionResult && result.extractionResult.id) {
          setExtractionJobId(result.extractionResult.id);

          // Store document ID for future reference
          if (result.documentId) {
            setDocumentId(result.documentId);
          }

          // Start polling for extraction status
          const interval = setInterval(() => {
            checkExtractionStatus(result.extractionResult.id);
          }, 5000); // Check every 5 seconds

          setPollingInterval(interval);
        } else {
          setStatus({
            message: "File uploaded but extraction job was not created properly.",
            type: "error",
          });
        }

        // Reset the file input
        const fileInput = document.querySelector('input[type="file"]') as HTMLInputElement;
        if (fileInput) fileInput.value = "";
      } else {
        setStatus({ message: `Upload failed: ${result.error || "Unknown error"}`, type: "error" });
      }
    } catch (error) {
      console.error("Upload error:", error);
      setStatus({
        message: `Upload failed: ${error instanceof Error ? error.message : "Unknown error"}`,
        type: "error",
      });
    } finally {
      setUploading(false);
    }
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
        const headerRange = sheet.getRange("A1:D1"); // Changed to 4 columns: Label, FY24, FY25, FY26
        const headerValues = [["", "FY24", "FY25", "FY26"]];
        headerRange.values = headerValues;
        headerRange.format.font.bold = true;
        headerRange.format.font.color = "#0066CC"; // Blue headers

        // Set column widths
        sheet.getRange("A:A").format.columnWidth = 150;
        sheet.getRange("B:D").format.columnWidth = 85; // Changed to only 3 data columns

        let currentRow = 2;

        // Calculate growth rates and ratios from FY24 data
        const revenueGrowthRate = 0.1; // 10% growth
        const marginExpansionRate = 0.02; // 2% margin expansion

        // Company Information
        const companyInfoItems = [
          ["Company Information", "", "", ""],
          ["Company Name:", data.companyProfile?.name || "", "", ""],
          ["Ticker:", data.companyProfile?.ticker || "", "", ""],
        ];

        const companyInfoRange = sheet.getRange(
          `A${currentRow}:D${currentRow + companyInfoItems.length - 1}`
        );
        companyInfoRange.values = companyInfoItems;

        // Format company info header
        sheet.getRange(`A${currentRow}`).format.font.bold = true;
        sheet.getRange(`A${currentRow}`).format.font.color = "#0066CC";

        currentRow += companyInfoItems.length + 1; // Add an empty row

        // Income Statement
        const revenue = data.financialStatements?.incomeStatement?.revenue || 0;
        const fy25Revenue = revenue * (1 + revenueGrowthRate);
        const fy26Revenue = fy25Revenue * (1 + revenueGrowthRate);

        const costOfRevenue = data.financialStatements?.incomeStatement?.costOfRevenue || 0;
        const grossMargin = revenue > 0 ? (revenue - costOfRevenue) / revenue : 0;
        const fy25GrossMargin = grossMargin + marginExpansionRate;
        const fy26GrossMargin = fy25GrossMargin + marginExpansionRate;

        const fy25CostOfRevenue = fy25Revenue * (1 - fy25GrossMargin);
        const fy26CostOfRevenue = fy26Revenue * (1 - fy26GrossMargin);

        const grossProfit =
          data.financialStatements?.incomeStatement?.grossProfit || revenue - costOfRevenue;
        const fy25GrossProfit = fy25Revenue - fy25CostOfRevenue;
        const fy26GrossProfit = fy26Revenue - fy26CostOfRevenue;

        const sgAndA = data.financialStatements?.incomeStatement?.sgAndA || 0;
        const sgAndAPercent = revenue > 0 ? sgAndA / revenue : 0;
        const fy25SgAndA = fy25Revenue * sgAndAPercent;
        const fy26SgAndA = fy26Revenue * sgAndAPercent;

        const rAndD = data.financialStatements?.incomeStatement?.rAndD || 0;
        const rAndDPercent = revenue > 0 ? rAndD / revenue : 0;
        const fy25RAndD = fy25Revenue * rAndDPercent;
        const fy26RAndD = fy26Revenue * rAndDPercent;

        const operatingExpenses =
          data.financialStatements?.incomeStatement?.operatingExpenses || sgAndA + rAndD;
        const fy25OpEx = fy25SgAndA + fy25RAndD;
        const fy26OpEx = fy26SgAndA + fy26RAndD;

        const operatingIncome =
          data.financialStatements?.incomeStatement?.operatingIncome ||
          grossProfit - operatingExpenses;
        const fy25OpIncome = fy25GrossProfit - fy25OpEx;
        const fy26OpIncome = fy26GrossProfit - fy26OpEx;

        const interestExpense = data.financialStatements?.incomeStatement?.interestExpense || 0;
        const fy25Interest = interestExpense; // Assume flat interest expense
        const fy26Interest = interestExpense;

        const incomeTaxExpense = data.financialStatements?.incomeStatement?.incomeTaxExpense || 0;
        const effectiveTaxRate =
          operatingIncome > 0 ? incomeTaxExpense / (operatingIncome - interestExpense) : 0.21;
        const fy25Tax = (fy25OpIncome - fy25Interest) * effectiveTaxRate;
        const fy26Tax = (fy26OpIncome - fy26Interest) * effectiveTaxRate;

        const netIncome =
          data.financialStatements?.incomeStatement?.netIncome ||
          operatingIncome - interestExpense - incomeTaxExpense;
        const fy25NetIncome = fy25OpIncome - fy25Interest - fy25Tax;
        const fy26NetIncome = fy26OpIncome - fy26Interest - fy26Tax;

        const depreciationAmortization =
          data.financialStatements?.incomeStatement?.depreciationAmortization || 0;
        const daPercent = revenue > 0 ? depreciationAmortization / revenue : 0;
        const fy25DA = fy25Revenue * daPercent;
        const fy26DA = fy26Revenue * daPercent;

        const incomeStatementItems = [
          ["Income Statement", "", "", ""],
          ["Revenue:", revenue, fy25Revenue, fy26Revenue],
          ["Cost of Revenue:", costOfRevenue, fy25CostOfRevenue, fy26CostOfRevenue],
          ["Gross Profit:", grossProfit, fy25GrossProfit, fy26GrossProfit],
          ["Operating Expenses:", "", "", ""],
          ["    SG&A:", sgAndA, fy25SgAndA, fy26SgAndA],
          ["    R&D:", rAndD, fy25RAndD, fy26RAndD],
          ["Total Operating Expenses:", operatingExpenses, fy25OpEx, fy26OpEx],
          ["Operating Income:", operatingIncome, fy25OpIncome, fy26OpIncome],
          ["Interest Expense:", interestExpense, fy25Interest, fy26Interest],
          ["Income Tax Expense:", incomeTaxExpense, fy25Tax, fy26Tax],
          ["Net Income:", netIncome, fy25NetIncome, fy26NetIncome],
        ];

        const incomeStatementRange = sheet.getRange(
          `A${currentRow}:D${currentRow + incomeStatementItems.length - 1}`
        );
        incomeStatementRange.values = incomeStatementItems;

        // Format numbers and colors
        const incomeStatementDataRange = sheet.getRange(
          `B${currentRow + 1}:D${currentRow + incomeStatementItems.length - 1}`
        );
        incomeStatementDataRange.numberFormat = [["#,##0;(#,##0);-"]];
        incomeStatementDataRange.format.horizontalAlignment = "Right";

        // Set colors for labels and headers
        const incomeStatementLabelRange = sheet.getRange(
          `A${currentRow + 1}:A${currentRow + incomeStatementItems.length - 1}`
        );
        incomeStatementLabelRange.format.font.color = "#000000"; // Black for regular text

        // Format income statement header and key subtotals
        sheet.getRange(`A${currentRow}`).format.font.bold = true;
        sheet.getRange(`A${currentRow}`).format.font.color = "#0066CC";
        sheet.getRange(`A${currentRow + 3}`).format.font.bold = true; // Gross Profit
        sheet.getRange(`A${currentRow + 4}`).format.font.color = "#666666"; // Operating Expenses
        sheet.getRange(`A${currentRow + 4}`).format.font.bold = true;
        sheet.getRange(`A${currentRow + 7}`).format.font.bold = true; // Total Operating Expenses
        sheet.getRange(`A${currentRow + 8}`).format.font.bold = true; // Operating Income
        sheet.getRange(`A${currentRow + 11}`).format.font.bold = true; // Net Income

        // Color the FY24 values blue
        const incomeStatementFY24Range = sheet.getRange(
          `B${currentRow + 1}:B${currentRow + incomeStatementItems.length - 1}`
        );
        incomeStatementFY24Range.format.font.color = "#0066CC";

        currentRow += incomeStatementItems.length + 1; // Add an empty row

        // Balance Sheet
        const cashAndEquivalents = data.financialStatements?.balanceSheet?.cashAndEquivalents || 0;
        const shortTermInvestments =
          data.financialStatements?.balanceSheet?.shortTermInvestments || 0;
        const accountsReceivable = data.financialStatements?.balanceSheet?.accountsReceivable || 0;
        const inventory = data.financialStatements?.balanceSheet?.inventory || 0;
        const totalCurrentAssets =
          data.financialStatements?.balanceSheet?.totalCurrentAssets ||
          cashAndEquivalents + shortTermInvestments + accountsReceivable + inventory;

        const ppe = data.financialStatements?.balanceSheet?.propertyPlantEquipment || 0;
        const goodwill = data.financialStatements?.balanceSheet?.goodwill || 0;
        const intangibleAssets = data.financialStatements?.balanceSheet?.intangibleAssets || 0;
        const totalAssets =
          data.financialStatements?.balanceSheet?.totalAssets ||
          totalCurrentAssets + ppe + goodwill + intangibleAssets;

        const accountsPayable = data.financialStatements?.balanceSheet?.accountsPayable || 0;
        const shortTermDebt = data.financialStatements?.balanceSheet?.shortTermDebt || 0;
        const totalCurrentLiabilities =
          data.financialStatements?.balanceSheet?.totalCurrentLiabilities ||
          accountsPayable + shortTermDebt;

        const longTermDebt = data.financialStatements?.balanceSheet?.longTermDebt || 0;
        const totalLiabilities =
          data.financialStatements?.balanceSheet?.totalLiabilities ||
          totalCurrentLiabilities + longTermDebt;

        const totalEquity =
          data.financialStatements?.balanceSheet?.totalEquity || totalAssets - totalLiabilities;

        const balanceSheetItems = [
          ["Balance Sheet", "", "", ""],
          ["Assets:", "", "", ""],
          ["    Cash & Equivalents:", cashAndEquivalents, "", ""],
          ["    Short-term Investments:", shortTermInvestments, "", ""],
          ["    Accounts Receivable:", accountsReceivable, "", ""],
          ["    Inventory:", inventory, "", ""],
          ["Total Current Assets:", totalCurrentAssets, "", ""],
          ["    Property, Plant & Equipment:", ppe, "", ""],
          ["    Goodwill:", goodwill, "", ""],
          ["    Intangible Assets:", intangibleAssets, "", ""],
          ["Total Assets:", totalAssets, "", ""],
          ["Liabilities:", "", "", ""],
          ["    Accounts Payable:", accountsPayable, "", ""],
          ["    Short-term Debt:", shortTermDebt, "", ""],
          ["Total Current Liabilities:", totalCurrentLiabilities, "", ""],
          ["    Long-term Debt:", longTermDebt, "", ""],
          ["Total Liabilities:", totalLiabilities, "", ""],
          ["Total Equity:", totalEquity, "", ""],
          ["Total Liabilities & Equity:", totalLiabilities + totalEquity, "", ""],
        ];

        const balanceSheetRange = sheet.getRange(
          `A${currentRow}:D${currentRow + balanceSheetItems.length - 1}`
        );
        balanceSheetRange.values = balanceSheetItems;

        // Format numbers and colors
        const balanceSheetDataRange = sheet.getRange(
          `B${currentRow + 2}:D${currentRow + balanceSheetItems.length - 1}`
        );
        balanceSheetDataRange.numberFormat = [["#,##0;(#,##0);-"]];
        balanceSheetDataRange.format.horizontalAlignment = "Right";

        // Set colors for labels and headers
        const balanceSheetLabelRange = sheet.getRange(
          `A${currentRow + 2}:A${currentRow + balanceSheetItems.length - 1}`
        );
        balanceSheetLabelRange.format.font.color = "#000000"; // Black for regular text

        // Format balance sheet header and key subtotals
        sheet.getRange(`A${currentRow}`).format.font.bold = true;
        sheet.getRange(`A${currentRow}`).format.font.color = "#0066CC";
        sheet.getRange(`A${currentRow + 1}`).format.font.color = "#666666"; // Assets
        sheet.getRange(`A${currentRow + 1}`).format.font.bold = true;
        sheet.getRange(`A${currentRow + 6}`).format.font.bold = true; // Total Current Assets
        sheet.getRange(`A${currentRow + 10}`).format.font.bold = true; // Total Assets
        sheet.getRange(`A${currentRow + 11}`).format.font.color = "#666666"; // Liabilities
        sheet.getRange(`A${currentRow + 11}`).format.font.bold = true;
        sheet.getRange(`A${currentRow + 14}`).format.font.bold = true; // Total Current Liabilities
        sheet.getRange(`A${currentRow + 16}`).format.font.bold = true; // Total Liabilities
        sheet.getRange(`A${currentRow + 17}`).format.font.bold = true; // Total Equity
        sheet.getRange(`A${currentRow + 18}`).format.font.bold = true; // Total Liabilities & Equity

        // Color the FY24 values blue
        const balanceSheetFY24Range = sheet.getRange(
          `B${currentRow + 2}:B${currentRow + balanceSheetItems.length - 1}`
        );
        balanceSheetFY24Range.format.font.color = "#0066CC";

        currentRow += balanceSheetItems.length + 1; // Add an empty row

        // Cash Flow Statement
        const operatingCashFlow =
          data.financialStatements?.cashFlowStatement?.operatingCashFlow || 0;
        const capex = data.financialStatements?.cashFlowStatement?.capitalExpenditures || 0;
        const acquisitions = data.financialStatements?.cashFlowStatement?.acquisitions || 0;
        const investingCashFlow =
          data.financialStatements?.cashFlowStatement?.investingCashFlow || capex + acquisitions;

        const debtIssuance = data.financialStatements?.cashFlowStatement?.debtIssuance || 0;
        const debtRepayment = data.financialStatements?.cashFlowStatement?.debtRepayment || 0;
        const stockRepurchase = data.financialStatements?.cashFlowStatement?.stockRepurchase || 0;
        const dividendsPaid = data.financialStatements?.cashFlowStatement?.dividendsPaid || 0;
        const financingCashFlow =
          data.financialStatements?.cashFlowStatement?.financingCashFlow ||
          debtIssuance - debtRepayment - stockRepurchase - dividendsPaid;

        const freeCashFlow =
          data.financialStatements?.cashFlowStatement?.freeCashFlow || operatingCashFlow - capex;

        const cashFlowItems = [
          ["Cash Flow Statement", "", "", ""],
          ["Operating Activities:", "", "", ""],
          ["    Net Income:", netIncome, fy25NetIncome, fy26NetIncome],
          ["    Depreciation & Amortization:", depreciationAmortization, fy25DA, fy26DA],
          ["    Stock Based Compensation:", "", "", ""],
          ["Operating Cash Flow:", operatingCashFlow, "", ""],
          ["Investing Activities:", "", "", ""],
          ["    Capital Expenditures:", capex, "", ""],
          ["    Acquisitions:", acquisitions, "", ""],
          ["Investing Cash Flow:", investingCashFlow, "", ""],
          ["Financing Activities:", "", "", ""],
          ["    Debt Issuance:", debtIssuance, "", ""],
          ["    Debt Repayment:", debtRepayment, "", ""],
          ["    Share Repurchases:", stockRepurchase, "", ""],
          ["    Dividends:", dividendsPaid, "", ""],
          ["Financing Cash Flow:", financingCashFlow, "", ""],
          ["Free Cash Flow:", freeCashFlow, "", ""],
        ];

        const cashFlowRange = sheet.getRange(
          `A${currentRow}:D${currentRow + cashFlowItems.length - 1}`
        );
        cashFlowRange.values = cashFlowItems;

        // Format numbers and colors
        const cashFlowDataRange = sheet.getRange(
          `B${currentRow + 2}:D${currentRow + cashFlowItems.length - 1}`
        );
        cashFlowDataRange.numberFormat = [["#,##0;(#,##0);-"]];
        cashFlowDataRange.format.horizontalAlignment = "Right";

        // Set colors for labels and headers
        const cashFlowLabelRange = sheet.getRange(
          `A${currentRow + 2}:A${currentRow + cashFlowItems.length - 1}`
        );
        cashFlowLabelRange.format.font.color = "#000000"; // Black for regular text

        // Format cash flow headers and totals
        sheet.getRange(`A${currentRow}`).format.font.bold = true;
        sheet.getRange(`A${currentRow}`).format.font.color = "#0066CC";
        sheet.getRange(`A${currentRow + 1}`).format.font.color = "#666666"; // Operating Activities
        sheet.getRange(`A${currentRow + 1}`).format.font.bold = true;
        sheet.getRange(`A${currentRow + 5}`).format.font.bold = true; // Operating Cash Flow
        sheet.getRange(`A${currentRow + 6}`).format.font.color = "#666666"; // Investing Activities
        sheet.getRange(`A${currentRow + 6}`).format.font.bold = true;
        sheet.getRange(`A${currentRow + 9}`).format.font.bold = true; // Investing Cash Flow
        sheet.getRange(`A${currentRow + 10}`).format.font.color = "#666666"; // Financing Activities
        sheet.getRange(`A${currentRow + 10}`).format.font.bold = true;
        sheet.getRange(`A${currentRow + 15}`).format.font.bold = true; // Financing Cash Flow
        sheet.getRange(`A${currentRow + 16}`).format.font.bold = true; // Free Cash Flow

        // Color the FY24 values blue
        const cashFlowFY24Range = sheet.getRange(
          `B${currentRow + 2}:B${currentRow + cashFlowItems.length - 1}`
        );
        cashFlowFY24Range.format.font.color = "#0066CC";

        // Set white background for entire used range and remove vertical borders
        const usedRange = sheet.getUsedRange();
        usedRange.format.fill.color = "white";
        usedRange.format.borders.getItem("EdgeLeft").style = "None";
        usedRange.format.borders.getItem("EdgeRight").style = "None";
        usedRange.format.borders.getItem("InsideVertical").style = "None";

        await context.sync();
        setStatus({ message: "Financial model generated successfully!", type: "success" });
      });
    } catch (error) {
      console.error("Error in Excel model generation:", error);
      setStatus({
        message: `Error generating Excel model: ${error instanceof Error ? error.message : "Unknown error"}`,
        type: "error",
      });
    }
  };

  return (
    <div className={styles.root}>
      <Label>Upload 10K Filing PDF</Label>
      <input
        type="file"
        accept=".pdf"
        onChange={handleFileChange}
        className={styles.input}
        disabled={uploading}
      />
      {file && <Label>Selected file: {file.name}</Label>}
      <Button
        className={styles.button}
        onClick={handleUploadAndGenerate}
        disabled={(!file && !existingDocumentId) || uploading}
        appearance="primary"
      >
        {uploading ? (
          <>
            <Spinner size="tiny" />
            <span style={{ marginLeft: "8px" }}>Processing...</span>
          </>
        ) : convexDocument ? (
          "Generate Financial Model from Existing File"
        ) : (
          "Upload and Generate Financial Model"
        )}
      </Button>
      {status.type && (
        <div className={styles.status}>
          <MessageBar
            intent={
              status.type === "error" ? "error" : status.type === "success" ? "success" : "info"
            }
          >
            {status.message}
          </MessageBar>
        </div>
      )}
    </div>
  );
};

export default CombinedFilingModelGenerator;
