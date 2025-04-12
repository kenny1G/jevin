import * as React from "react";
import { makeStyles, Label, Button, Spinner, MessageBar } from "@fluentui/react-components";
import { useConvex, useQuery } from "convex/react";
import { api } from "../../../convex/_generated/api";
import { useFinancialData, FinancialData } from "../context/FinancialDataContext";

/* global Excel, console */

// Using FinancialData interface from context

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
  const [extractionJobId, setExtractionJobId] = React.useState<string | null>(null);
  const [documentId, setDocumentId] = React.useState<string | null>(null);
  const [existingDocumentId, setExistingDocumentId] = React.useState<string | null>(null);
  const [pollingInterval, setPollingInterval] = React.useState<NodeJS.Timeout | null>(null);

  // Use the financial data context
  const { setFinancialData, isLoading, setIsLoading, status, setStatus } = useFinancialData();

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

  // Function to check extraction status
  const checkExtractionStatus = async (jobId: string) => {
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
              // Set the financial data in the context
              setFinancialData(latestRun.output.data as FinancialData);
              // await generateFinancialModel(latestRun.output.data as FinancialData);
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
  };

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

          // Set the financial data in the context
          setFinancialData(latestRun as FinancialData);

          // Process the extraction result
          // await generateFinancialModel(latestRun as FinancialData);
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
