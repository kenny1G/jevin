# Jevin - Excel Add-in with LlamaIndex Integration

This Excel add-in allows users to upload PDF files to LlamaIndex for processing.

## Setup

### Prerequisites

- Node.js and npm
- Convex account
- LlamaIndex Cloud API key

### Environment Variables

You need to set up the following environment variable in your Convex project:

1. Go to your Convex dashboard
2. Navigate to Settings > Environment Variables
3. Add a new environment variable:
   - Name: `LLAMA_CLOUD_API_KEY`
   - Value: Your LlamaIndex Cloud API key

Alternatively, you can set the environment variable using the Convex CLI:

```bash
npx convex env set LLAMA_CLOUD_API_KEY "your-api-key"
```

### Development

1. Install dependencies:
   ```bash
   npm install
   ```

2. Start the development server:
   ```bash
   npm run dev-server
   ```

3. Start the Excel add-in:
   ```bash
   npm run start
   ```

## How It Works

1. The user selects a PDF file in the Excel add-in
2. The file is uploaded to a Convex HTTP action
3. The Convex action forwards the file to LlamaIndex with the proper API key
4. The response from LlamaIndex is returned to the user

## Troubleshooting

If you encounter CORS issues, make sure your Convex deployment is properly configured to allow requests from your Excel add-in's origin.