# Image URL Conversion Test Examples

## Feature Overview
The server now automatically detects any content wrapped in `{{}}` and converts it to an HTML `<img>` tag.

## Supported Formats

### 1. URLs without protocol
```
{{://ibb.co/G4P1rxLC}}
```
Converts to:
```html
<img src="https://i.ibb.co/G4P1rxLC/image.png" alt="Image" style="max-width: 100%; height: auto; display: block; margin: 10px 0;" />
```

### 2. Full URLs with HTTPS
```
{{https://picsum.photos/200/300?grayscale}}
```
Converts to:
```html
<img src="https://picsum.photos/200/300?grayscale" alt="Image" style="max-width: 100%; height: auto; display: block; margin: 10px 0;" />
```

### 3. Full URLs with HTTP
```
{{http://example.com/image.jpg}}
```
Converts to:
```html
<img src="http://example.com/image.jpg" alt="Image" style="max-width: 100%; height: auto; display: block; margin: 10px 0;" />
```

### 4. URLs without any protocol
```
{{example.com/image.png}}
```
Converts to:
```html
<img src="https://example.com/image.png" alt="Image" style="max-width: 100%; height: auto; display: block; margin: 10px 0;" />
```

## Special Handling

### ImgBB Links
The server has special handling for ImgBB share links and converts them to direct image URLs:
- Input: `{{://ibb.co/G4P1rxLC}}` or `{{https://ibb.co/G4P1rxLC}}`
- Output: Uses `https://i.ibb.co/G4P1rxLC/image.png` (direct image URL)

## How to Use in Word Documents

Simply include the image URL placeholder in your Word document text:

```
This is some text with an image: {{https://picsum.photos/200/300?grayscale}}

And here's another image: {{://ibb.co/G4P1rxLC}}

More text continues here...
```

When you upload the Word document to either `/api/decode` or `/api/section`, the server will:
1. Extract the text from the Word document
2. Find all `{{...}}` patterns
3. Convert them to responsive HTML img tags
4. Insert the result into the SharePoint template

## Image Styling

All generated images have responsive styling:
- `max-width: 100%` - Prevents images from overflowing their container
- `height: auto` - Maintains aspect ratio
- `display: block` - Makes image a block element
- `margin: 10px 0` - Adds vertical spacing around the image
