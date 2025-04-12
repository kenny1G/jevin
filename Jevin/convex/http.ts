import { httpRouter } from "convex/server";
import { httpAction } from "./_generated/server";
import { v } from "convex/values";
import { api } from "./_generated/api";

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

      // Create a new FormData to send to LlamaIndex
      console.log("🔄 Creating FormData for LlamaIndex...");
      const llamaFormData = new FormData();
      llamaFormData.append("upload_file", file, file.name);
      console.log("✅ FormData created with file:", file.name);

      // Forward the request to LlamaIndex
      console.log("🌐 Sending request to LlamaIndex...");
      const response = await fetch("https://api.cloud.llamaindex.ai/api/v1/files", {
        method: "POST",
        headers: {
          "Authorization": `Bearer ${apiKey}`,
        },
        body: llamaFormData,
      });

      console.log("📥 LlamaIndex response status:", response.status);
      console.log("📥 LlamaIndex response headers:", JSON.stringify(headersToObject(response.headers)));

      // Get the response from LlamaIndex
      console.log("📄 Parsing LlamaIndex response...");
      const responseData = await response.json();
      console.log("✅ LlamaIndex response data:", JSON.stringify(responseData, null, 2));

      // Return the response with CORS headers
      console.log("📤 Sending response back to client...");
      return new Response(JSON.stringify(responseData), {
        status: response.status,
        headers: {
          "Content-Type": "application/json",
          "Access-Control-Allow-Origin": "*", // Adjust this for production
          "Access-Control-Allow-Methods": "POST, OPTIONS",
          "Access-Control-Allow-Headers": "Content-Type, Authorization",
        },
      });
    } catch (error) {
      console.error("❌ Error uploading file:", error);
      console.error("❌ Error stack:", error instanceof Error ? error.stack : "No stack trace");
      return new Response(JSON.stringify({ error: "Failed to upload file" }), {
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


export const helloWorld = httpAction(async () => {
  return new Response("Hello, World!", {
    status: 200,
    headers: {
      "Content-Type": "text/plain",
    },
  });
});



http.route({
  path: "/hello",
  method: "GET",
  handler: helloWorld,
});


export default http;