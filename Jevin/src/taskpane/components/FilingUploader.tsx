import * as React from "react";
import { makeStyles, Label, Button, Spinner } from "@fluentui/react-components";
import { useConvex } from "convex/react";

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

const FilingUploader: React.FC = () => {
  const styles = useStyles();
  const convex = useConvex();
  const [file, setFile] = React.useState<File | null>(null);
  const [uploading, setUploading] = React.useState(false);
  const [status, setStatus] = React.useState<string | null>(null);

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFile = event.target.files?.[0];
    if (selectedFile && selectedFile.type === "application/pdf") {
      setFile(selectedFile);
      setStatus(null);
    }
  };

  const handleUpload = async () => {
    if (!file) return;

    try {
      setUploading(true);
      setStatus("Uploading file...");

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
        setStatus("File uploaded successfully!");
        setFile(null);
        // Reset the file input
        const fileInput = document.querySelector('input[type="file"]') as HTMLInputElement;
        if (fileInput) fileInput.value = "";
      } else {
        setStatus(`Upload failed: ${result.error || "Unknown error"}`);
      }
    } catch (error) {
      console.error("Upload error:", error);
      setStatus(`Upload failed in catch: ${error instanceof Error ? error.message : "Unknown error"}`);
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
        onClick={handleUpload}
        disabled={!file || uploading}
      >
        {uploading ? <Spinner size="tiny" /> : "Upload to LlamaIndex"}
      </Button>
      {status && <div className={styles.status}>{status}</div>}
    </div>
  );
};

export default FilingUploader;