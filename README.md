# Base64 Decode API

A simple Node.js API that receives POST requests with Base64 encoded data, decodes it, and returns status and message.

## Installation

```bash
npm install
```

## Usage

### Start the server

```bash
npm start
```

For development with auto-reload:

```bash
npm run dev
```

The server will run on `http://localhost:3000` by default.

## API Endpoints

### POST /api/decode

Decodes Base64 encoded data with content type information.

**Request Body:**
```json
{
  "$content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  "$content": "UEsDBAoAAAAAAAAAIQD..."
}
```

**Success Response (200):**
```json
{
  "status": "success",
  "message": "[Decoded document text content]"
}
```

**Error Response (400):**
```json
{
  "status": "error",
  "message": "No content provided. Please send Base64 encoded data in the \"$content\" field."
}
```

### GET /health

Health check endpoint to verify the API is running.

**Response (200):**
```json
{
  "status": "success",
  "message": "API is running"
}
```

## Example Usage

### Using curl:

```bash
curl -X POST http://localhost:3000/api/decode \
  -H "Content-Type: application/json" \
  -d '{"$content-type":"text/plain","$content":"SGVsbG8gV29ybGQh"}'
```

### Using Postman:

1. Set method to POST
2. URL: `http://localhost:3000/api/decode`
3. Headers: `Content-Type: application/json`
4. Body (raw JSON):
   ```json
   {
     "$content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
     "$content": "UEsDBAoAAAAAAAAAIQD..."
   }
   ```

### Using JavaScript (fetch):

```javascript
fetch('http://localhost:3000/api/decode', {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json'
  },
  body: JSON.stringify({
    '$content-type': 'text/plain',
    '$content': 'SGVsbG8gV29ybGQh'
  })
})
  .then(response => response.json())
  .then(data => console.log(data));
```
