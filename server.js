const express = require('express');
const bodyParser = require('body-parser');
const mammoth = require('mammoth');

const app = express();
const PORT = process.env.PORT || 3500

// Middleware to parse JSON and urlencoded data
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

// POST endpoint to decode Base64 data
app.post('/api/decode', async (req, res) => {
    try {
        // Log incoming request for debugging
        console.log('Received request body:', JSON.stringify(req.body, null, 2));

        // Handle both direct body and nested body formats
        let requestBody = req.body;

        // Check if request has uri, method, headers, body structure (nested format)
        if (req.body.uri && req.body.method && req.body.body) {
            requestBody = req.body.body;
            console.log('Extracted from nested format (with uri):', JSON.stringify(requestBody, null, 2));
        }
        // Check if body is nested inside a wrapper but without uri/method
        else if (req.body.body && req.body.body['$content']) {
            requestBody = req.body.body;
            console.log('Extracted from nested format (body wrapper):', JSON.stringify(requestBody, null, 2));
        }

        const { '$content-type': contentType, '$content': content } = requestBody;
        console.log('Extracted fields - contentType:', contentType, 'content length:', content ? content.length : 'undefined');

        // Validate if content is provided
        if (!content) {
            return res.status(400).json({
                status: 'error',
                message: 'No content provided. Please send Base64 encoded data in the "$content" field.'
            });
        }

        // Validate content-type
        if (!contentType) {
            return res.status(400).json({
                status: 'error',
                message: 'No content-type provided. Please send the content type in the "$content-type" field.'
            });
        }

        // Decode Base64
        const decodedData = Buffer.from(content, 'base64');

        // Check if it's a Word document
        if (contentType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
            try {
                // Extract text from Word document
                const result = await mammoth.extractRawText({ buffer: decodedData });

                return res.status(200).json({
                    status: 'success',
                    message: result.value || 'Word document text extracted successfully'
                });
            } catch (extractError) {
                return res.status(500).json({
                    status: 'error',
                    message: `Failed to extract text from Word document: ${extractError.message}`
                });
            }
        }

        // Determine if content type is text-based or binary
        const textBasedTypes = [
            'text/',
            'application/json',
            'application/xml',
            'application/javascript'
        ];

        const isTextBased = textBasedTypes.some(type => contentType.startsWith(type));

        let responseData;
        if (isTextBased) {
            // For text-based content, decode to UTF-8
            responseData = decodedData.toString('utf-8');
        } else {
            // For binary content (PDFs, images), return base64
            responseData = decodedData.toString('base64');
        }

        // Return success response with decoded data as message
        return res.status(200).json({
            status: 'success',
            message: responseData
        });

    } catch (error) {
        // Handle any errors during decoding
        return res.status(500).json({
            status: 'error',
            message: `Error decoding Base64 data: ${error.message}`
        });
    }
});

// Health check endpoint
app.get('/health', (req, res) => {
    res.status(200).json({
        status: 'success',
        message: 'API is running'
    });
});

// Start the server
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
    console.log(`POST endpoint: http://localhost:${PORT}/api/decode`);
    console.log(`Health check: http://localhost:${PORT}/health`);
});
