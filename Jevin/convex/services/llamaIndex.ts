import { v } from "convex/values";

// Interface for LlamaIndex file response
export interface LlamaIndexFileResponse {
  id: string;
  created_at: string;
  updated_at: string;
  name: string;
  external_file_id: string;
  file_size: number;
  file_type: string;
  project_id: string;
  last_modified_at: string;
  resource_info: {
    file_size: number;
    last_modified_at: string;
  };
  permission_info: any;
  data_source_id: string | null;
}

/**
 * Upload a file to LlamaIndex
 * @param file The file to upload
 * @param apiKey The LlamaIndex API key
 * @returns The LlamaIndex response
 */
export async function uploadFileToLlamaIndex(
  file: File,
  apiKey: string
): Promise<LlamaIndexFileResponse> {
  // Create a new FormData to send to LlamaIndex
  const llamaFormData = new FormData();
  llamaFormData.append("upload_file", file, file.name);

  // Forward the request to LlamaIndex
  const response = await fetch("https://api.cloud.llamaindex.ai/api/v1/files", {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${apiKey}`,
    },
    body: llamaFormData,
  });

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    throw new Error(
      `Failed to upload file to LlamaIndex: ${response.status} ${response.statusText} ${JSON.stringify(errorData)}`
    );
  }

  // Get the response from LlamaIndex
  const responseData = await response.json();
  return responseData as LlamaIndexFileResponse;
}