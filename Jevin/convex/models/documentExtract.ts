import { v } from "convex/values";
import { mutation, query } from "../_generated/server";
import { Id } from "../_generated/dataModel";

// Create a new document extract
export const createDocumentExtract = mutation({
  args: {
    documentId: v.id("documents"),
    extractionJobId: v.string(),
    version: v.optional(v.number()),
  },
  handler: async (ctx, args) => {
    const {
      documentId,
      extractionJobId,
    } = args;

    // Create a new document extract record
    const extractId = await ctx.db.insert("documentExtracts", {
      documentId,
      extractionJobId,
    });

    return extractId;
  },
});

// Get a document extract by document ID
export const getDocumentExtractByDocumentId = query({
  args: {
    documentId: v.id("documents"),
  },
  handler: async (ctx, args) => {
    const { documentId } = args;

    // Get the most recent extract for this document
    const extracts = await ctx.db
      .query("documentExtracts")
      .filter((q) => q.eq(q.field("documentId"), documentId))
      .order("desc")
      .take(1);

    return extracts[0] || null;
  },
});

// Get a document extract by ID
export const getDocumentExtractById = query({
  args: {
    extractId: v.id("documentExtracts"),
  },
  handler: async (ctx, args) => {
    const { extractId } = args;

    return await ctx.db.get(extractId);
  },
});

// Update a document extract