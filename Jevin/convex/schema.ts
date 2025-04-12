import { defineSchema, defineTable } from "convex/server";
import { v } from "convex/values";

export default defineSchema({
  documents: defineTable({
    // The name of the document as it was uploaded
    name: v.string(),

    // The ID returned by LlamaIndex after successful upload
    llamaIndexId: v.string(),

    // The file type (e.g., "pdf")
    fileType: v.string(),

    // The size of the file in bytes
    fileSize: v.number(),

    // When the document was uploaded to LlamaIndex
    createdAt: v.string(),

    // When the document was last updated in LlamaIndex
    updatedAt: v.string(),

    // External file ID (usually the filename)
    externalFileId: v.string(),

    // LlamaIndex project ID
    projectId: v.string(),

    // Last modified timestamp
    lastModifiedAt: v.string(),

    // Status of the document in our system (e.g., "uploaded", "processing", "ready", "error")
    status: v.string(),

    // Optional error message if status is "error"
    errorMessage: v.optional(v.string()),
  }).index("by_name", ["name"]),

  // Table for storing extracted JSON from documents
  documentExtracts: defineTable({
    // Reference to the document
    documentId: v.id("documents"),

    extractionJobId: v.string(),
  }),
});
