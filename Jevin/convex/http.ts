import { httpRouter } from "convex/server";
import { httpAction } from "./_generated/server";
import { v } from "convex/values";
import { api } from "./_generated/api";
import { createDocument } from "./models/document";
import { uploadFileToLlamaIndex } from "./services/llamaIndex";
import { createDocumentExtract } from "./models/documentExtract";

const http = httpRouter();

// Helper function to convert headers to a plain object
function headersToObject(headers: Headers): Record<string, string> {
  const result: Record<string, string> = {};
  headers.forEach((value, key) => {
    result[key] = value;
  });
  return result;
}

// Route for uploading files to LlamaIndex
http.route({
  path: "/upload-to-llamaindex",
  method: "POST",
  handler: httpAction(async (ctx, request) => {
    console.log("🚀 HTTP action started: /upload-to-llamaindex");
    console.log("📝 Request headers:", JSON.stringify(headersToObject(request.headers)));

    try {
      // Get the file from the request
      console.log("📦 Parsing form data...");
      const formData = await request.formData();
      const file = formData.get("file") as File;

      if (!file) {
        console.log("❌ No file provided in the request");
        return new Response(JSON.stringify({ error: "No file provided" }), {
          status: 400,
          headers: { "Content-Type": "application/json" },
        });
      }

      console.log("✅ File received:", {
        name: file.name,
        type: file.type,
        size: file.size,
      });

      // Get the API key from environment variables
      console.log("🔑 Checking for API key...");
      const apiKey = process.env.LLAMA_CLOUD_API_KEY;
      if (!apiKey) {
        console.log("❌ API key not configured");
        return new Response(JSON.stringify({ error: "API key not configured" }), {
          status: 500,
          headers: { "Content-Type": "application/json" },
        });
      }
      console.log("✅ API key found (length: " + apiKey.length + ")");

      // Upload the file to LlamaIndex using our service
      console.log("🌐 Uploading file to LlamaIndex...");
      const llamaIndexResponse = await uploadFileToLlamaIndex(file, apiKey);
      console.log("✅ File uploaded to LlamaIndex:", llamaIndexResponse.id);

      // Store the document in our database using our model
      console.log("💾 Storing document in database...");
      const documentId = await ctx.runMutation(api["models/document"].createDocument, {
        name: file.name,
        llamaIndexResponse,
      });
      console.log("✅ Document stored in database with ID:", documentId);

      // Extract JSON from the document
      console.log("🔍 Extracting JSON from document...");
      const extractionAgentId = "b50ed79e-390f-4553-8604-08ba2e5bb22b";
      const fileId = llamaIndexResponse.id;

      // Make the extraction request to LlamaIndex API
      const extractionResponse = await fetch(
        "https://api.cloud.llamaindex.ai/api/v1/extraction/jobs",
        {
          method: "POST",
          headers: {
            "accept": "application/json",
            "Content-Type": "application/json",
            "Authorization": `Bearer ${apiKey}`,
          },
          body: JSON.stringify({
            extraction_agent_id: extractionAgentId,
            file_id: fileId,
          }),
        }
      );

      if (!extractionResponse.ok) {
        const errorText = await extractionResponse.text();
        console.error("❌ Extraction request failed:", errorText);
        throw new Error(`Extraction request failed: ${extractionResponse.status} ${extractionResponse.statusText}`);
      }

      const extractionResult = await extractionResponse.json();
      console.log("✅ Extraction job created:", extractionResult);

      // Store the extraction result in our database
      console.log("💾 Storing extraction result in database...");
      const extractId = await ctx.runMutation(api["models/documentExtract"].createDocumentExtract, {
        documentId,
        extractionJobId: extractionResult.id,
      });
      console.log("✅ Extraction result stored with ID:", extractId);

      // Return the response with CORS headers
      console.log("📤 Sending response back to client...");
      return new Response(JSON.stringify({
        success: true,
        documentId,
        llamaIndexId: llamaIndexResponse.id,
        extractionId: extractionResult.id,
        extractId,
      }), {
        status: 200,
        headers: {
          "Content-Type": "application/json",
          "Access-Control-Allow-Origin": "*", // Adjust this for production
          "Access-Control-Allow-Methods": "POST, OPTIONS",
          "Access-Control-Allow-Headers": "Content-Type, Authorization",
        },
      });
    } catch (error) {
      console.error("❌ Error processing file:", error);
      console.error("❌ Error stack:", error instanceof Error ? error.stack : "No stack trace");
      return new Response(JSON.stringify({
        error: "Failed to process file",
        message: error instanceof Error ? error.message : "Unknown error"
      }), {
        status: 500,
        headers: { "Content-Type": "application/json" },
      });
    } finally {
      console.log("🏁 HTTP action completed: /upload-to-llamaindex");
    }
  }),
});

// Handle preflight requests for CORS
http.route({
  path: "/upload-to-llamaindex",
  method: "OPTIONS",
  handler: httpAction(async () => {
    console.log("🔄 Handling OPTIONS preflight request");
    return new Response(null, {
      status: 204,
      headers: {
        "Access-Control-Allow-Origin": "*", // Adjust this for production
        "Access-Control-Allow-Methods": "POST, OPTIONS",
        "Access-Control-Allow-Headers": "Content-Type, Authorization",
        "Access-Control-Max-Age": "86400", // 24 hours
      },
    });
  }),
});

// Route for fetching extraction results from LlamaIndex
http.route({
  path: "/fetch-extraction-result",
  method: "GET",
  handler: httpAction(async (_, request) => {
    console.log("🚀 HTTP action started: /fetch-extraction-result");

    try {
      // Get the job ID from the URL parameters
      const url = new URL(request.url);
      const jobId = url.searchParams.get("jobId");

      if (!jobId) {
        console.log("❌ No job ID provided in the request");
        return new Response(JSON.stringify({ error: "No job ID provided" }), {
          status: 400,
          headers: { "Content-Type": "application/json" },
        });
      }

      console.log("✅ Job ID received:", jobId);

      // Get the API key from environment variables
      console.log("🔑 Checking for API key...");
      const apiKey = process.env.LLAMA_CLOUD_API_KEY;
      if (!apiKey) {
        console.log("❌ API key not configured");
        return new Response(JSON.stringify({ error: "API key not configured" }), {
          status: 500,
          headers: { "Content-Type": "application/json" },
        });
      }
      console.log("✅ API key found (length: " + apiKey.length + ")");

      // Fetch the extraction result from LlamaIndex API
      console.log("🌐 Fetching extraction result from LlamaIndex...");
      const extractionResponse = await fetch(
        `https://api.cloud.llamaindex.ai/api/v1/extraction/runs/by-job/${jobId}`,
        {
          method: "GET",
          headers: {
            "accept": "application/json",
            "Authorization": `Bearer ${apiKey}`,
          },
        }
      );

      if (!extractionResponse.ok) {
        const errorText = await extractionResponse.text();
        console.error("❌ Extraction result fetch failed:", errorText);
        throw new Error(`Extraction result fetch failed: ${extractionResponse.status} ${extractionResponse.statusText}`);
      }

      const extractionResult = await extractionResponse.json();
      console.log("✅ Extraction result fetched successfully");

      // Return the response with CORS headers
      console.log("📤 Sending response back to client...");
      return new Response(JSON.stringify({
        success: true,
        extractionResult,
      }), {
        status: 200,
        headers: {
          "Content-Type": "application/json",
          "Access-Control-Allow-Origin": "*", // Adjust this for production
          "Access-Control-Allow-Methods": "GET, OPTIONS",
          "Access-Control-Allow-Headers": "Content-Type, Authorization",
        },
      });
    } catch (error) {
      console.error("❌ Error fetching extraction result:", error);
      console.error("❌ Error stack:", error instanceof Error ? error.stack : "No stack trace");
      return new Response(JSON.stringify({
        error: "Failed to fetch extraction result",
        message: error instanceof Error ? error.message : "Unknown error"
      }), {
        status: 500,
        headers: { "Content-Type": "application/json" },
      });
    } finally {
      console.log("🏁 HTTP action completed: /fetch-extraction-result");
    }
  }),
});

// Handle preflight requests for CORS for the new endpoint
http.route({
  path: "/fetch-extraction-result",
  method: "OPTIONS",
  handler: httpAction(async () => {
    console.log("🔄 Handling OPTIONS preflight request for fetch-extraction-result");
    return new Response(null, {
      status: 204,
      headers: {
        "Access-Control-Allow-Origin": "*", // Adjust this for production
        "Access-Control-Allow-Methods": "GET, OPTIONS",
        "Access-Control-Allow-Headers": "Content-Type, Authorization",
        "Access-Control-Max-Age": "86400", // 24 hours
      },
    });
  }),
});

export default http;