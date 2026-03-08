const express = require('express');
const bodyParser = require('body-parser');
const mammoth = require('mammoth');

const app = express();
const PORT = process.env.PORT || 3500

// Parse all requests as raw binary first, then detect format
app.post('/api/decode', bodyParser.raw({ type: '*/*', limit: '50mb' }), async (req, res) => {
    try {
        let isJSON = false;
        let jsonBody = null;

        // Try to detect if it's JSON by checking the first character
        if (req.body.length > 0 && (req.body[0] === 0x7B || req.body[0] === 0x5B)) { // { or [
            try {
                jsonBody = JSON.parse(req.body.toString('utf8'));
                isJSON = true;
                console.log('Detected JSON format');
            } catch (e) {
                // Not valid JSON, treat as binary
                console.log('Not valid JSON, treating as binary');
            }
        }

        // Handle raw binary Word document
        if (!isJSON) {
            console.log('Received raw binary data, length:', req.body.length);

            // Check if it's a Word document (starts with PK signature)
            if (req.body[0] === 0x50 && req.body[1] === 0x4B) {
                try {
                    const result = await mammoth.extractRawText({ buffer: req.body });
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
            } else {
                return res.status(400).json({
                    status: 'error',
                    message: 'Unsupported binary format. Please send a Word document (.docx) or JSON with Base64-encoded content.'
                });
            }
        }

        // Handle JSON format
        console.log('Processing JSON request:', JSON.stringify(jsonBody, null, 2));
        let requestBody = jsonBody;

        // Check if request has uri, method, headers, body structure (nested format)
        if (jsonBody.uri && jsonBody.method && jsonBody.body) {
            requestBody = jsonBody.body;
            console.log('Extracted from nested format (with uri):', JSON.stringify(requestBody, null, 2));
        }
        // Check if body is nested inside a wrapper but without uri/method
        else if (jsonBody.body && jsonBody.body['$content']) {
            requestBody = jsonBody.body;
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
