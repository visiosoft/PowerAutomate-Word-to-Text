## Test the API

### Example 1: Direct Body Format (Using curl)

```bash
curl -X POST http://localhost:3500/api/decode -H "Content-Type: application/json" -d "{\"$content-type\":\"text/plain\",\"$content\":\"SGVsbG8gV29ybGQh\"}"
```

### Example 1b: Nested Body Format (Using curl)

```bash
curl -X POST http://localhost:3500/api/decode -H "Content-Type: application/json" -d '{"uri":"http://localhost:3500/api/decode","method":"POST","headers":{"Content-Type":"application/json"},"body":{"$content-type":"text/plain","$content":"SGVsbG8gV29ybGQh"}}'
```

Expected Response:
```json
{
  "status": "success",
  "message": "Hello World!"
}
```

### Example 2: Direct Format (PowerShell)

```powershell
$body = @{
    '$content-type' = "text/plain"
    '$content' = "SGVsbG8gV29ybGQh"
} | ConvertTo-Json

Invoke-RestMethod -Uri "http://localhost:3500/api/decode" -Method Post -Body $body -ContentType "application/json"
```

### Example 2b: Nested Format (PowerShell)

```powershell
$body = @{
    uri = "http://localhost:3500/api/decode"
    method = "POST"
    headers = @{
        'Content-Type' = "application/json"
    }
    body = @{
        '$content-type' = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        '$content' = "UEsDBAoAAAAAAAAAIQD/////TwEAAE8BAAAQAAAAW3RyYXNoXS8wMDAwLmRhdP..."
    }
} | ConvertTo-Json -Depth 5

Invoke-RestMethod -Uri "http://localhost:3500/api/decode" -Method Post -Body $body -ContentType "application/json"
```

### Example 3: Word Document Direct Format (PowerShell)

```powershell
$body = @{
    '$content-type' = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    '$content' = "UEsDBAoAAAAAAAAAIQD/////TwEAAE8BAAAQAAAAW3RyYXNoXS8wMDAwLmRhdP////8..."
} | ConvertTo-Json

Invoke-RestMethod -Uri "http://localhost:3500/api/decode" -Method Post -Body $body -ContentType "application/json"
```

### Example 4: Test with missing content

```bash
curl -X POST http://localhost:3500/api/decode -H "Content-Type: application/json" -d "{}"
```

Expected Response:
```json
{
  "status": "error",
  "message": "No content provided. Please send Base64 encoded data in the \"$content\" field."
}
```

### Supported Content Types:

- `text/plain` - Plain text
- `application/vnd.openxmlformats-officedocument.wordprocessingml.document` - Word Document (.docx)
- `application/pdf` - PDF Document
- `image/png` - PNG Image
- `image/jpeg` - JPEG Image
