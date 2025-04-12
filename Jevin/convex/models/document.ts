import { v } from "convex/values";
import { mutation } from "../_generated/server";
import { Id } from "../_generated/dataModel";

// Interface for LlamaIndex file response
interface LlamaIndexFileResponse {
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

// Create a new document record in the database
export const createDocument = mutation({
  args: {
    name: v.string(),
    llamaIndexResponse: v.object({
      id: v.string(),
      created_at: v.string(),
      updated_at: v.string(),
      name: v.string(),
      external_file_id: v.string(),
      file_size: v.number(),
      file_type: v.string(),
      project_id: v.string(),
      last_modified_at: v.string(),
      resource_info: v.object({
        file_size: v.number(),
        last_modified_at: v.string(),
      }),
      permission_info: v.any(),
      data_source_id: v.union(v.string(), v.null()),
    }),
  },
  handler: async (ctx, args) => {
    const { name, llamaIndexResponse } = args;

    // Create a new document record
    const documentId = await ctx.db.insert("documents", {
      name,
      llamaIndexId: llamaIndexResponse.id,
      fileType: llamaIndexResponse.file_type,
      fileSize: llamaIndexResponse.file_size,
      createdAt: llamaIndexResponse.created_at,
      updatedAt: llamaIndexResponse.updated_at,
      externalFileId: llamaIndexResponse.external_file_id,
      projectId: llamaIndexResponse.project_id,
      lastModifiedAt: llamaIndexResponse.last_modified_at,
      status: "uploaded", // Initial status
    });

    return documentId;
  },
});

// Update document status
export const updateDocumentStatus = mutation({
  args: {
    documentId: v.id("documents"),
    status: v.string(),
    errorMessage: v.optional(v.string()),
  },
  handler: async (ctx, args) => {
    const { documentId, status, errorMessage } = args;

    await ctx.db.patch(documentId, {
      status,
      ...(errorMessage ? { errorMessage } : {}),
    });

    return documentId;
  },
});

// Get a document by ID
export const getDocumentById = mutation({
  args: {
    documentId: v.id("documents"),
  },
  handler: async (ctx, args) => {
    const { documentId } = args;

    const document = await ctx.db.get(documentId);
    return document;
  },
});