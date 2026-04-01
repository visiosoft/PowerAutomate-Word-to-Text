const express = require('express');
const bodyParser = require('body-parser');
const mammoth = require('mammoth');

const app = express();
const PORT = process.env.PORT || 3500

// Navigation configuration for Company Policy pages
// Maps section titles to their previous/next URLs
// To add a new section, add an entry with the exact H1 title from your Word document
// Example: 'Your Section Title': { previous: 'url-to-previous', next: 'url-to-next' }
const NAVIGATION_CONFIG = {
    'Introduction': {
        sectionNumber: 1,
        previous: '/sites/InnovationArtsEntertainment/SitePages/HANDBOOK.aspx',
        next: '/sites/InnovationArtsEntertainment/SitePages/Employment%20Policies%20and%20Practices.aspx'
    },
    'Employment Policies and Practices': {
        sectionNumber: 2,
        previous: '/sites/InnovationArtsEntertainment/SitePages/Introduction.aspx',
        next: '/sites/InnovationArtsEntertainment/SitePages/Company%20Policies%20and%20Practices.aspx'
    },
    'Company Policies and Practices': {
        sectionNumber: 3,
        previous: '/sites/InnovationArtsEntertainment/SitePages/Employment%20Policies%20and%20Practices.aspx',
        next: '/sites/InnovationArtsEntertainment/SitePages/Compensation%20and%20Benefits.aspx'
    },
    'Compensation and Benefits': {
        sectionNumber: 4,
        previous: '/sites/InnovationArtsEntertainment/SitePages/Company%20Policies%20and%20Practices.aspx',
        next: '/sites/InnovationArtsEntertainment/SitePages/Work%20Performance.aspx'
    },
    'Work Performance': {
        sectionNumber: 5,
        previous: '/sites/InnovationArtsEntertainment/SitePages/Compensation%20and%20Benefits.aspx',
        next: '/sites/InnovationArtsEntertainment/SitePages/Department%20Guides%20and%20Procedures.aspx'
    },
    'Procedures and Guidelines': {
        sectionNumber: 6,
        previous: '/sites/InnovationArtsEntertainment/SitePages/Work%20Performance.aspx',
        next: '/sites/InnovationArtsEntertainment/SitePages/Employee%20Acknowledgment%20Form.aspx'
    },
    'Employee Acknowledgment': {
        sectionNumber: 7,
        previous: '/sites/InnovationArtsEntertainment/SitePages/Department%20Guides%20and%20Procedures.aspx',
        next: '/sites/InnovationArtsEntertainment/SitePages/Executive%20Addendum.aspx'
    },
    'Executive Addendum': {
        sectionNumber: 8,
        previous: '/sites/InnovationArtsEntertainment/SitePages/Employee%20Acknowledgment%20Form.aspx',
        next: '/sites/InnovationArtsEntertainment/SitePages/HANDBOOK.aspx'
    }
};

// Helper function to split content into sections
function splitContentIntoSections(htmlContent) {
    const sections = {
        ceoMessage: '',
        companyInfo: ''
    };

    // Split by heading tags while preserving them in the result
    const headingRegex = /(<h[1-6][^>]*>.*?<\/h[1-6]>)/gi;
    const parts = htmlContent.split(headingRegex);

    // Build an array of sections with headings and their content
    const contentSections = [];
    for (let i = 0; i < parts.length; i++) {
        const part = parts[i].trim();
        if (!part) continue;

        // Check if this part is a heading
        const headingMatch = part.match(/<h[1-6][^>]*>(.*?)<\/h[1-6]>/i);
        if (headingMatch) {
            const headingText = headingMatch[1].replace(/<[^>]*>/g, '').trim().toLowerCase();
            const level = parseInt(part.match(/<h([1-6])/i)[1]);
            const nextContent = i + 1 < parts.length ? parts[i + 1].trim() : '';

            contentSections.push({
                heading: part,
                headingText: headingText,
                level: level,
                content: nextContent
            });
            i++; // Skip the next part as we've already used it as content
        }
    }

    // Find the Introduction section (H1)
    let introSections = contentSections;
    const introIdx = contentSections.findIndex(s => s.level === 1 && s.headingText.includes('introduction'));

    if (introIdx !== -1) {
        // Get all sections from Introduction until the next H1
        const nextH1Idx = contentSections.findIndex((s, idx) => idx > introIdx && s.level === 1);
        const endIdx = nextH1Idx !== -1 ? nextH1Idx : contentSections.length;
        introSections = contentSections.slice(introIdx, endIdx);
    }

    // Within intro sections, find CEO message and Company sections
    let ceoIdx = introSections.findIndex(s =>
        s.headingText.includes('message') && s.headingText.includes('ceo')
    );

    let companyIdx = introSections.findIndex(s =>
        s.headingText.includes('the company') || s.headingText === 'company'
    );

    // Extract CEO message section
    if (ceoIdx !== -1) {
        const ceoEndIdx = companyIdx !== -1 ? companyIdx : introSections.length;
        sections.ceoMessage = introSections
            .slice(ceoIdx, ceoEndIdx)
            .map(s => s.heading + s.content)
            .join('');
    }

    // Extract Company info section
    if (companyIdx !== -1) {
        sections.companyInfo = introSections
            .slice(companyIdx)
            .map(s => s.heading + s.content)
            .join('');
    }

    return sections;
}

// Helper function to extract all H1 sections
function extractAllH1Sections(htmlContent) {
    const h1Sections = [];

    // Split by heading tags while preserving them in the result
    const headingRegex = /(<h[1-6][^>]*>.*?<\/h[1-6]>)/gi;
    const parts = htmlContent.split(headingRegex);

    // Build an array of sections with headings and their content
    const allSections = [];
    for (let i = 0; i < parts.length; i++) {
        const part = parts[i].trim();
        if (!part) continue;

        // Check if this part is a heading
        const headingMatch = part.match(/<h[1-6][^>]*>(.*?)<\/h[1-6]>/i);
        if (headingMatch) {
            const headingText = headingMatch[1].replace(/<[^>]*>/g, '').trim();
            const level = parseInt(part.match(/<h([1-6])/i)[1]);
            const nextContent = i + 1 < parts.length ? parts[i + 1].trim() : '';

            allSections.push({
                heading: part,
                headingText: headingText,
                level: level,
                content: nextContent
            });
            i++; // Skip the next part as we've already used it as content
        }
    }

    // Find all H1 sections
    const h1Indices = [];
    allSections.forEach((section, idx) => {
        if (section.level === 1) {
            h1Indices.push(idx);
        }
    });

    // Extract content for each H1 section (everything until the next H1)
    h1Indices.forEach((h1Idx, i) => {
        const startIdx = h1Idx;
        const endIdx = i + 1 < h1Indices.length ? h1Indices[i + 1] : allSections.length;

        // Combine all content from this H1 to the next H1
        const sectionContent = allSections
            .slice(startIdx, endIdx)
            .map(s => s.heading + s.content)
            .join('');

        h1Sections.push({
            title: allSections[h1Idx].headingText,
            content: sectionContent
        });
    });

    return h1Sections;
}

// Helper function to extract H2 tags from HTML content
function extractH2Sections(htmlContent) {
    const h2Sections = [];
    const h2Regex = /<h2[^>]*>(.*?)<\/h2>/gi;
    let match;

    while ((match = h2Regex.exec(htmlContent)) !== null) {
        const heading = match[1].replace(/<[^>]*>/g, '').trim();
        // Create anchor-friendly ID from heading
        const anchorId = heading.toLowerCase()
            .replace(/[^\w\s-]/g, '')
            .replace(/\s+/g, '-')
            .replace(/-+/g, '-');

        h2Sections.push({
            heading: heading,
            anchorId: anchorId
        });
    }

    return h2Sections;
}

// Helper function to add anchor IDs to H2 tags in HTML content
function addAnchorIdsToH2(htmlContent) {
    return htmlContent.replace(/<h2([^>]*)>(.*?)<\/h2>/gi, (match, attrs, content) => {
        const heading = content.replace(/<[^>]*>/g, '').trim();
        const anchorId = heading.toLowerCase()
            .replace(/[^\w\s-]/g, '')
            .replace(/\s+/g, '-')
            .replace(/-+/g, '-');
        return `<h2${attrs} id="${anchorId}">${content}</h2>`;
    });
}

// Helper function to generate navigation buttons HTML with custom styles
// variant: 'top' for white background, 'bottom' for black background
function generateNavigationButtonsHTML(previousUrl, homeUrl, nextUrl, variant = 'bottom') {
    let buttonStyle, buttonHoverPrefix, textColor;

    if (variant === 'top') {
        // White background, black text for top navigation
        buttonStyle = 'background-color: #ffffff; color: #000000 !important; padding: 12px 24px; text-decoration: none !important; border-radius: 6px; border: 2px solid #000000; box-shadow: 0 2px 6px rgba(0,0,0,0.3); cursor: pointer; font-weight: 600; text-align: center; margin: 5px; transition: all 0.3s ease; flex: 0 1 auto; min-width: 150px; display: inline-block;';
        buttonHoverPrefix = 'onmouseover="this.style.backgroundColor=\'#f0f0f0\'; this.style.borderColor=\'#333333\'; this.style.boxShadow=\'0 4px 10px rgba(0,0,0,0.5)\';" onmouseout="this.style.backgroundColor=\'#ffffff\'; this.style.borderColor=\'#000000\'; this.style.boxShadow=\'0 2px 6px rgba(0,0,0,0.3)\';"';
        textColor = '#000000';
    } else {
        // Black background, white text for bottom navigation
        buttonStyle = 'background-color: #000000; color: #ffffff !important; padding: 12px 24px; text-decoration: none !important; border-radius: 6px; border: 2px solid #000000; box-shadow: 0 2px 6px rgba(0,0,0,0.3); cursor: pointer; font-weight: 600; text-align: center; margin: 5px; transition: all 0.3s ease; flex: 0 1 auto; min-width: 150px; display: inline-block;';
        buttonHoverPrefix = 'onmouseover="this.style.backgroundColor=\'#333333\'; this.style.borderColor=\'#555555\'; this.style.boxShadow=\'0 4px 10px rgba(0,0,0,0.5)\';" onmouseout="this.style.backgroundColor=\'#000000\'; this.style.borderColor=\'#000000\'; this.style.boxShadow=\'0 2px 6px rgba(0,0,0,0.3)\';"';
        textColor = '#ffffff';
    }

    const html = `<style>.nav-button-container{display:flex;flex-wrap:wrap;justify-content:center;align-items:center;padding:20px 10px;margin:20px 0;gap:10px;}@media (max-width:768px){.nav-button-container{flex-direction:column;padding:0px 10px;margin:0px 0 5px 0 !important;}.nav-button-container a{width:100%;max-width:300px;}}</style><div class="nav-button-container"><a href="${previousUrl}" style="${buttonStyle}" ${buttonHoverPrefix}><span style="color: ${textColor} !important; text-decoration: none !important; font-size: 16px;">← Previous Section</span></a><a href="${homeUrl}" style="${buttonStyle}" ${buttonHoverPrefix}><span style="color: ${textColor} !important; text-decoration: none !important; font-size: 16px;">Home</span></a><a href="${nextUrl}" style="${buttonStyle}" ${buttonHoverPrefix}><span style="color: ${textColor} !important; text-decoration: none !important; font-size: 16px;">Next Section →</span></a></div>`;

    return html;
}

// Helper function to generate section contents HTML with numbered links
function generateSectionContentsHTML(h2Sections, sectionNumber = 3) {
    if (!h2Sections || h2Sections.length === 0) {
        return '';
    }

    // Use inline styles for better SharePoint compatibility
    const buttonStyle = 'background-color: #000000; color: #ffffff !important; padding: 14px 20px; margin: 10px 0px; display: block; text-decoration: none !important; border-radius: 6px; border: 2px solid #000000; box-shadow: 0 2px 6px rgba(0,0,0,0.3); cursor: pointer; font-weight: 500; transition: all 0.3s ease;';
    const buttonHoverStyle = 'onmouseover="this.style.backgroundColor=\'#333333\'; this.style.borderColor=\'#555555\'; this.style.transform=\'translateX(5px)\'; this.style.boxShadow=\'0 4px 10px rgba(0,0,0,0.5)\';" onmouseout="this.style.backgroundColor=\'#000000\'; this.style.borderColor=\'#000000\'; this.style.transform=\'translateX(0px)\'; this.style.boxShadow=\'0 2px 6px rgba(0,0,0,0.3)\';"';

    let html = '<style>html { scroll-behavior: smooth; }</style>';

    h2Sections.forEach((section, index) => {
        const number = `${sectionNumber}.${index + 1}`;
        html += `<div style="margin: 10px 0px;"><a href="#${section.anchorId}" style="${buttonStyle}" ${buttonHoverStyle}><span style="color: #ffffff !important; text-decoration: none !important;"><span class="fontSizeMediumPlus"><span lang="EN-US" dir="ltr"><strong>${number}</strong> ${section.heading}</span></span></span></a></div>`;
    });

    return html;
}

// Helper function to add top anchor and back-to-top button to content
function addBackToTopButton(htmlContent) {
    // Add anchor at the beginning
    const contentWithAnchor = '<a id="page-top"></a>' + htmlContent;

    // Add floating back-to-top button at the end with inline styles
    const backToTopStyle = 'position: fixed; bottom: 30px; right: 30px; background-color: #000000; color: #ffffff !important; padding: 15px 20px; border-radius: 50px; text-decoration: none !important; font-weight: bold; box-shadow: 0 4px 12px rgba(0,0,0,0.4); transition: all 0.3s ease; z-index: 1000; border: 2px solid #333333;';
    const backToTopHover = 'onmouseover="this.style.backgroundColor=\'#333333\'; this.style.transform=\'translateY(-5px)\'; this.style.boxShadow=\'0 6px 16px rgba(0,0,0,0.6)\'; this.style.borderColor=\'#555555\';" onmouseout="this.style.backgroundColor=\'#000000\'; this.style.transform=\'translateY(0px)\'; this.style.boxShadow=\'0 4px 12px rgba(0,0,0,0.4)\'; this.style.borderColor=\'#333333\';"';
    const backToTopButton = `<a href="#page-top" style="${backToTopStyle}" ${backToTopHover}>↑ Top</a>`;

    return contentWithAnchor + backToTopButton;
}

// Helper function to add employee paragraph styles to all <p> tags
function applyEmployeeParagraphStyles(htmlContent) {
    // Add inline styles to all paragraph tags
    const styles = 'text-align: justify !important; line-height: 1.8 !important; margin-bottom: 12px !important;';

    // Replace all <p> tags (with or without existing attributes)
    return htmlContent.replace(/<p(\s[^>]*)?>/gi, function (match, attributes) {
        if (attributes && attributes.includes('style=')) {
            // If style attribute exists, append to it
            return match.replace(/style="([^"]*)"/i, `style="$1 ${styles}"`);
        } else {
            // Add new style attribute
            const attrs = attributes || '';
            return `<p${attrs} style="${styles}">`;
        }
    });
}
// Helper function to convert {{image_url}} placeholders to HTML img tags
function convertImagePlaceholdersToImgTags(htmlContent) {
    // Pattern to match {{anything}} - captures content inside double curly braces
    const imagePlaceholderPattern = /\{\{([^}]+)\}\}/g;

    return htmlContent.replace(imagePlaceholderPattern, function (match, urlContent) {
        // Trim whitespace from the URL
        let imageUrl = urlContent.trim();

        // If URL doesn't start with http:// or https://, prepend https://
        if (!imageUrl.startsWith('http://') && !imageUrl.startsWith('https://')) {
            // Handle cases like ://ibb.co/something
            if (imageUrl.startsWith('://')) {
                imageUrl = 'https' + imageUrl;
            } else {
                imageUrl = 'https://' + imageUrl;
            }
        }

        // Return HTML img tag with responsive styling
        return `<img src="${imageUrl}" alt="Image" style="max-width: 100%; height: auto; display: block; margin: 10px 0;" />`;
    });
}
// SharePoint page template with placeholder
const INTRO_TEMPLATE = {
    "__metadata": {
        "type": "SP.Publishing.SitePage"
    },
    "PageRenderingState": 0,
    "CanvasContent1": "[{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"c12d16a8-30de-46cb-8d44-f22d081aace5\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":1},\"zoneHeight\":550,\"id\":\"3dbcc936-0d52-4658-91f8-dc3cf63925c1\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":0,\"y\":0,\"w\":70,\"h\":7,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<h2 class=\\\"headingSpacingAbove headingSpacingBelow lineHeight1_0\\\"><span class=\\\"fontSizeMega rte-fontscale-font-max\\\">1.1 A&nbsp;Message&nbsp;from the CEO</span></h2>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"c12d16a8-30de-46cb-8d44-f22d081aace5\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":2},\"zoneHeight\":550,\"id\":\"1857d513-039d-4197-8e49-4bd8e6f378c4\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":0,\"y\":6,\"w\":70,\"h\":6,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><span class=\\\"fontSizeXLarge\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\">##CEOMESSAGE##</span></span></p><p class=\\\"noSpacingAbove noSpacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" aria-hidden=\\\"true\\\" data-text-type=\\\"noSpacing\\\">&nbsp;</p><p class=\\\"noSpacingAbove noSpacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" aria-hidden=\\\"true\\\" data-text-type=\\\"noSpacing\\\">&nbsp;</p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"c12d16a8-30de-46cb-8d44-f22d081aace5\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":3},\"zoneHeight\":550,\"id\":\"bd7c16e0-b868-43ce-9742-bd272082eb3c\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":25,\"y\":29,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"bd7c16e0-b868-43ce-9742-bd272082eb3c\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Home\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/HANDBOOK.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Center\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"c12d16a8-30de-46cb-8d44-f22d081aace5\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":4},\"zoneHeight\":550,\"id\":\"2a4fe9c1-c46f-4664-ac1b-b0179329f01b\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":38,\"y\":29,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"2a4fe9c1-c46f-4664-ac1b-b0179329f01b\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Next Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Employment%20Policies%20and%20Practices.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Center\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"0ba22869-c07a-4c44-bccc-521aae856fe8\",\"sectionIndex\":1,\"sectionFactor\":8,\"controlIndex\":1},\"id\":\"48047102-2fb5-4efc-ac96-034716e6f250\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"innerHTML\":\"<p><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\">##COMPANYINFO##</span></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"0ba22869-c07a-4c44-bccc-521aae856fe8\",\"sectionIndex\":2,\"sectionFactor\":4,\"controlIndex\":1},\"id\":\"4696d89e-ec85-4c9b-8ebf-e04f59ab5343\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"innerHTML\":\"<p class=\\\"noSpacingAbove noSpacingBelow\\\" data-text-type=\\\"noSpacing\\\"><span class=\\\"fontSizeXLargePlus\\\"><span class=\\\"fontColorThemeDarker\\\"><strong>Section Contents</strong></span></span></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"0ba22869-c07a-4c44-bccc-521aae856fe8\",\"sectionIndex\":2,\"sectionFactor\":4,\"controlIndex\":2},\"id\":\"da32caa7-5244-4649-8293-35d10e746a80\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"innerHTML\":\"<p class=\\\"lineHeight1_2\\\" style=\\\"margin-left:0px;\\\"><a href=\\\"#a-message-from-the-ceo\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>1.1 </strong>CEO</span> Message</span></a></p><p class=\\\"lineHeight1_2\\\" style=\\\"margin-left:0px;\\\"><a href=\\\"#the-company\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>1.2</strong> The Company</span>&nbsp;</span></a></p><p class=\\\"lineHeight1_2\\\" style=\\\"margin-left:0px;\\\"><a href=\\\"#change-in-policy\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>1.3 </strong>Change in Policy</span>&nbsp;</span></a></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":3,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":1},\"zoneHeight\":1600,\"id\":\"9a58ab09-29e4-4d76-b0f5-8b8a114cb40e\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":36,\"y\":87,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"9a58ab09-29e4-4d76-b0f5-8b8a114cb40e\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Next Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Employment%20Policies%20and%20Practices.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Right\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":3,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":2},\"zoneHeight\":1600,\"id\":\"602be227-1a60-4fd8-a4bd-3f69c27a0fa8\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":3,\"y\":87,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"602be227-1a60-4fd8-a4bd-3f69c27a0fa8\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Home\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/HANDBOOK.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Center\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":3,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":3},\"zoneHeight\":1600,\"id\":\"8048e3e4-74fd-4e6b-9d4f-cfb5649079d8\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":5,\"y\":85,\"w\":40,\"h\":0,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"2161a1c6-db61-4731-b97c-3cdb303f7cbb\",\"instanceId\":\"8048e3e4-74fd-4e6b-9d4f-cfb5649079d8\",\"title\":\"Divider\",\"description\":\"Add a line to divide areas on your page.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{},\"links\":{}},\"dataVersion\":\"1.2\",\"properties\":{\"minimumLayoutWidth\":1,\"length\":100,\"weight\":1},\"containsDynamicDataSource\":false},\"webPartId\":\"2161a1c6-db61-4731-b97c-3cdb303f7cbb\",\"reservedWidth\":688,\"reservedHeight\":1},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isAIGeneratedDescription\":false,\"isDefaultThumbnail\":true,\"isSpellCheckEnabled\":true,\"globalRichTextStylingVersion\":1,\"rtePageSettings\":{\"contentVersion\":5,\"indentationVersion\":2},\"isEmailReady\":true,\"webPartsPageSettings\":{\"isTitleHeadingLevelsEnabled\":true,\"isLowQualityImagePlaceholderEnabled\":true}}},{\"controlType\":14,\"webPartData\":{\"properties\":{\"zoneBackground\":{\"c12d16a8-30de-46cb-8d44-f22d081aace5\":{\"type\":\"image\",\"imageData\":{\"fileName\":\"68234-IAE_Display_28x22_FINAL.jpg\",\"width\":2016,\"height\":720,\"source\":0,\"siteId\":\"44583a03-7429-4de6-9641-ae749e56727f\",\"webId\":\"454df1df-1248-4cb7-a41d-a4e482d69fbb\",\"listId\":\"2ab24772-b566-4fe3-abab-a9edde27e5ea\",\"uniqueId\":\"df6e1993-3198-4e40-9f91-db508a198a82\"},\"overlay\":{\"color\":\"#000000\",\"opacity\":0},\"useLightText\":true},\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\":{\"type\":\"image\",\"imageData\":{\"fileName\":\"template11Image0003.png\",\"width\":144,\"height\":144,\"source\":1},\"overlay\":{\"color\":\"#FFFFFF\",\"opacity\":100},\"useLightText\":false}},\"zoneThemeIndex\":{\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\":{\"lightIndex\":0,\"darkIndex\":0}}},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{\"zoneBackground.c12d16a8-30de-46cb-8d44-f22d081aace5.imageData.url\":\"/sites/InnovationArtsEntertainment/SiteAssets/SitePages/VolunteerCenter(1)/68234-IAE_Display_28x22_FINAL.jpg\",\"zoneBackground.54f8a2d2-b737-4a6e-96a2-54ae0a01c7da.imageData.url\":\"https://cdn.hubblecontent.osi.office.net/m365content/publish/4b99b582-026b-4490-ab73-b0091ab721cc/image.png\"},\"links\":{}},\"dataVersion\":\"1.0\"}}]"
}

// SECTION  template with ##LEFT## ETC placeholder
const SECTIONTEMPLATE = {
    "__metadata": {
        "type": "SP.Publishing.SitePage"
    },
    "PageRenderingState": 0,
    "CanvasContent1": "[{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"f374e17d-f830-40c9-859a-b641fed7e09c\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":1},\"zoneHeight\":258,\"id\":\"3dbcc936-0d52-4658-91f8-dc3cf63925c1\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":0,\"y\":3,\"w\":70,\"h\":4,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<style>@media (max-width:768px){#mobile-title-push{margin-bottom:60px !important;}}</style><div id=\\\"mobile-title-push\\\" style=\\\"margin-bottom:0;\\\"><h2 class=\\\"headingSpacingAbove headingSpacingBelow lineHeight2_4\\\" style=\\\"text-align:center;margin-bottom:0;\\\"><span class=\\\"fontSizeMega rte-fontscale-font-max\\\">##TITLE##</span></h2></div>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"f374e17d-f830-40c9-859a-b641fed7e09c\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":2},\"zoneHeight\":258,\"id\":\"nav-buttons-top\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":0,\"y\":7,\"w\":70,\"h\":5,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"##NAVBUTTONS_TOP##\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"2021b9db-8033-4d75-a200-4e720b6b50a1\",\"sectionIndex\":1,\"sectionFactor\":4,\"controlIndex\":1},\"id\":\"2c2b1596-0197-4020-a475-e6fab7eef288\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"innerHTML\":\"##RIGHT##\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"2021b9db-8033-4d75-a200-4e720b6b50a1\",\"sectionIndex\":2,\"sectionFactor\":8,\"controlIndex\":1},\"id\":\"ca89cf71-9ce3-4c01-8f23-ccb08babf9fb\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"innerHTML\":\"##LEFT##\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"2021b9db-8033-4d75-a200-4e720b6b50a1\",\"sectionIndex\":2,\"sectionFactor\":8,\"controlIndex\":2},\"id\":\"e4d6a6d2-b463-48d0-a58a-3c70029ae93a\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"innerHTML\":\"##NAVBUTTONS_BOTTOM##\",\"contentVersion\":5},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isAIGeneratedDescription\":false,\"isDefaultThumbnail\":true,\"isSpellCheckEnabled\":true,\"globalRichTextStylingVersion\":1,\"rtePageSettings\":{\"contentVersion\":5,\"indentationVersion\":2},\"isEmailReady\":true,\"webPartsPageSettings\":{\"isTitleHeadingLevelsEnabled\":true,\"isLowQualityImagePlaceholderEnabled\":true}}},{\"controlType\":14,\"webPartData\":{\"properties\":{\"zoneBackground\":{\"f374e17d-f830-40c9-859a-b641fed7e09c\":{\"type\":\"image\",\"imageData\":{\"fileName\":\"68234-IAE_Display_28x22_FINAL.jpg\",\"width\":2016,\"height\":720,\"source\":0,\"siteId\":\"44583a03-7429-4de6-9641-ae749e56727f\",\"webId\":\"454df1df-1248-4cb7-a41d-a4e482d69fbb\",\"listId\":\"2ab24772-b566-4fe3-abab-a9edde27e5ea\",\"uniqueId\":\"df6e1993-3198-4e40-9f91-db508a198a82\"},\"overlay\":{\"color\":\"#000000\",\"opacity\":0},\"useLightText\":true}}},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{\"zoneBackground.f374e17d-f830-40c9-859a-b641fed7e09c.imageData.url\":\"/sites/InnovationArtsEntertainment/SiteAssets/SitePages/VolunteerCenter(1)/68234-IAE_Display_28x22_FINAL.jpg\"},\"links\":{}},\"dataVersion\":\"1.0\"}}]"

}

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
            } catch (e) {
                // Not valid JSON, treat as binary
            }
        }

        // Handle raw binary Word document
        if (!isJSON) {
            // Check if it's a Word document (starts with PK signature)
            if (req.body[0] === 0x50 && req.body[1] === 0x4B) {
                try {
                    const result = await mammoth.convertToHtml({
                        buffer: req.body
                    }, {
                        styleMap: [
                            "b => strong",
                            "i => em",
                            "u => u",
                            "p[style-name='Heading 3'] => strong",
                            "p[style-name='Heading 4'] => strong",
                            "p[style-name='Heading 5'] => strong",
                            "p[style-name='List Paragraph'] => p:fresh > strong"
                        ].join("\n")
                    });
                    const htmlContent = result.value || '';

                    // Split content into sections
                    const sections = splitContentIntoSections(htmlContent);

                    // Escape the HTML for JSON string (preserve newlines and formatting)
                    const escapedCeoMessage = sections.ceoMessage
                        .replace(/\\/g, '\\\\')
                        .replace(/"/g, '\\"')
                        .replace(/\n/g, '\\n');
                    const escapedCompanyInfo = sections.companyInfo
                        .replace(/\\/g, '\\\\')
                        .replace(/"/g, '\\"')
                        .replace(/\n/g, '\\n');

                    // Replace both placeholders with converted HTML
                    let canvasContent = INTRO_TEMPLATE.CanvasContent1
                        .replace('##CEOMESSAGE##', escapedCeoMessage)
                        .replace('##COMPANYINFO##', escapedCompanyInfo);

                    const response = {
                        ...INTRO_TEMPLATE,
                        CanvasContent1: canvasContent
                    };

                    return res.status(200).json({
                        status: 'success',
                        message: JSON.stringify(response)
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
        let requestBody = jsonBody;

        // Check if request has uri, method, headers, body structure (nested format)
        if (jsonBody.uri && jsonBody.method && jsonBody.body) {
            requestBody = jsonBody.body;
        }
        // Check if body is nested inside a wrapper but without uri/method
        else if (jsonBody.body && jsonBody.body['$content']) {
            requestBody = jsonBody.body;
        }

        const { '$content-type': contentType, '$content': content } = requestBody;

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
                // Extract text with formatting from Word document
                const result = await mammoth.convertToHtml({
                    buffer: decodedData
                }, {
                    styleMap: [
                        "b => strong",
                        "i => em",
                        "u => u",
                        "p[style-name='Heading 3'] => strong",
                        "p[style-name='Heading 4'] => strong",
                        "p[style-name='Heading 5'] => strong",
                        "p[style-name='List Paragraph'] => p:fresh > strong"
                    ].join("\n")
                });
                let htmlContent = result.value || '';

                // Convert {{image_url}} placeholders to img tags
                htmlContent = convertImagePlaceholdersToImgTags(htmlContent);

                // Split content into sections
                const sections = splitContentIntoSections(htmlContent);

                // Escape the HTML for JSON string (preserve newlines and formatting)
                const escapedCeoMessage = sections.ceoMessage
                    .replace(/\\/g, '\\\\')
                    .replace(/"/g, '\\"')
                    .replace(/\n/g, '\\n');
                const escapedCompanyInfo = sections.companyInfo
                    .replace(/\\/g, '\\\\')
                    .replace(/"/g, '\\"')
                    .replace(/\n/g, '\\n');

                // Replace both placeholders with converted HTML
                let canvasContent = INTRO_TEMPLATE.CanvasContent1
                    .replace('##CEOMESSAGE##', escapedCeoMessage)
                    .replace('##COMPANYINFO##', escapedCompanyInfo);

                const response = {
                    ...INTRO_TEMPLATE,
                    CanvasContent1: canvasContent
                };

                return res.status(200).json({
                    status: 'success',
                    message: JSON.stringify(response)
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

// Company Policy endpoint - uses sectionTEMPLATE with ##MESSAGE##
app.post('/api/section', bodyParser.raw({ type: '*/*', limit: '50mb' }), async (req, res) => {
    try {
        let isJSON = false;
        let jsonBody = null;

        // Try to detect if it's JSON by checking the first character
        if (req.body.length > 0 && (req.body[0] === 0x7B || req.body[0] === 0x5B)) {
            try {
                jsonBody = JSON.parse(req.body.toString('utf8'));
                isJSON = true;
            } catch (e) {
                // Not valid JSON, treat as binary
            }
        }
        debugger;
        // Handle raw binary Word document
        if (!isJSON) {
            if (req.body[0] === 0x50 && req.body[1] === 0x4B) {
                try {
                    const result = await mammoth.convertToHtml({
                        buffer: req.body
                    }, {
                        includeDefaultStyleMap: true,
                        styleMap: [
                            "b => strong",
                            "i => em",
                            "u => u",
                            "p[style-name='Heading 1'] => h1:fresh",
                            "p[style-name='Heading 2'] => h2:fresh",
                            "p[style-name='Heading 3'] => h3:fresh",
                            "p[style-name='Heading 4'] => strong",
                            "p[style-name='Heading 5'] => strong",
                            "p[style-name='List Paragraph'] => p:fresh > strong"
                        ].join("\n")
                    });
                    let htmlContent = result.value || '';

                    // Convert {{image_url}} placeholders to img tags
                    htmlContent = convertImagePlaceholdersToImgTags(htmlContent);

                    // Log any warnings about content that couldn't be converted
                    if (result.messages && result.messages.length > 0) {
                        console.log('Mammoth conversion warnings:', result.messages);
                    }

                    // Check for optional title filter
                    const rawTitle = req.query.title || jsonBody?.title || jsonBody?.['$title'];
                    const titleFilter = rawTitle ? rawTitle.trim() : undefined;
                    const h1Sections = extractAllH1Sections(htmlContent);

                    let contentToInsert = htmlContent;

                    if (titleFilter) {
                        const matchedSection = h1Sections.find(section =>
                            section.title.toLowerCase().includes(titleFilter.toLowerCase())
                        );

                        if (matchedSection) {
                            contentToInsert = matchedSection.content;
                        } else {
                            return res.status(404).json({
                                status: 'error',
                                message: `Section with title "${titleFilter}" not found`,
                                availableTitles: h1Sections.map(s => s.title)
                            });
                        }
                    }

                    // Apply employee paragraph styles to content
                    contentToInsert = applyEmployeeParagraphStyles(contentToInsert);

                    // Add anchor IDs to H2 tags so links work
                    contentToInsert = addAnchorIdsToH2(contentToInsert);

                    // Add back-to-top button
                    contentToInsert = addBackToTopButton(contentToInsert);

                    // Extract H1 title for ##Title## from the actual content section
                    const h1Match = contentToInsert.match(/<h1[^>]*>(.*?)<\/h1>/i);
                    const h1Title = h1Match ? h1Match[1].replace(/<[^>]*>/g, '').trim() : 'Company Policy';

                    // Get navigation config based on title
                    const navigationKey = Object.keys(NAVIGATION_CONFIG).find(key =>
                        h1Title.toLowerCase().includes(key.toLowerCase())
                    );
                    const navigation = navigationKey ? NAVIGATION_CONFIG[navigationKey] : {
                        sectionNumber: 3,
                        previous: '/sites/InnovationArtsEntertainment/SitePages/HANDBOOK.aspx',
                        next: '/sites/InnovationArtsEntertainment/SitePages/HANDBOOK.aspx'
                    };

                    // Extract H2 sections for #RightSECTION##
                    const h2Sections = extractH2Sections(contentToInsert);
                    const sectionContentsHTML = generateSectionContentsHTML(h2Sections, navigation.sectionNumber);

                    // Escape the HTML for JSON string (preserve newlines and formatting)
                    const escapedMessage = contentToInsert
                        .replace(/\\/g, '\\\\')
                        .replace(/"/g, '\\"')
                        .replace(/\n/g, '\\n');

                    const escapedTitle = h1Title
                        .replace(/\\/g, '\\\\')
                        .replace(/"/g, '\\"')
                        .replace(/\n/g, '\\n');

                    const escapedSectionContents = sectionContentsHTML
                        .replace(/\\/g, '\\\\')
                        .replace(/"/g, '\\"')
                        .replace(/\n/g, '\\n');

                    // Set navigation URLs from config
                    const previousUrl = navigation.previous;
                    const nextUrl = navigation.next;
                    const homeUrl = '/sites/InnovationArtsEntertainment/SitePages/HANDBOOK.aspx';

                    // Generate custom navigation buttons HTML for top (white bg) and bottom (black bg)
                    const navButtonsTopHTML = generateNavigationButtonsHTML(previousUrl, homeUrl, nextUrl, 'top');
                    const navButtonsBottomHTML = generateNavigationButtonsHTML(previousUrl, homeUrl, nextUrl, 'bottom');

                    const escapedNavButtonsTop = navButtonsTopHTML
                        .replace(/\\/g, '\\\\')
                        .replace(/"/g, '\\"')
                        .replace(/\n/g, '\\n');

                    const escapedNavButtonsBottom = navButtonsBottomHTML
                        .replace(/\\/g, '\\\\')
                        .replace(/"/g, '\\"')
                        .replace(/\n/g, '\\n');

                    // Replace placeholders
                    let canvasContent = SECTIONTEMPLATE.CanvasContent1
                        .replace('##LEFT##', escapedMessage)
                        .replace('##TITLE##', escapedTitle)
                        .replace('##RIGHT##', escapedSectionContents)
                        .replace('##NAVBUTTONS_TOP##', escapedNavButtonsTop)
                        .replace('##NAVBUTTONS_BOTTOM##', escapedNavButtonsBottom)
                        .replaceAll('##PREVIOUS_URL##', previousUrl)
                        .replaceAll('##HOME_URL##', homeUrl)
                        .replaceAll('##NEXT_URL##', nextUrl);

                    const response = {
                        ...SECTIONTEMPLATE,
                        CanvasContent1: canvasContent
                    };

                    return res.status(200).json({
                        status: 'success',
                        message: JSON.stringify(response)
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

        // Handle JSON input with Base64 encoded content
        if (isJSON) {
            const base64Content = jsonBody['$content'];

            if (!base64Content) {
                return res.status(400).json({
                    status: 'error',
                    message: 'Missing $content in request body'
                });
            }

            try {
                const decodedBuffer = Buffer.from(base64Content, 'base64');

                if (decodedBuffer[0] === 0x50 && decodedBuffer[1] === 0x4B) {
                    const result = await mammoth.convertToHtml({
                        buffer: decodedBuffer
                    }, {
                        includeDefaultStyleMap: true,
                        styleMap: [
                            "b => strong",
                            "i => em",
                            "u => u",
                            "p[style-name='Heading 1'] => h1:fresh",
                            "p[style-name='Heading 2'] => h2:fresh",
                            "p[style-name='Heading 3'] => h3:fresh",
                            "p[style-name='Heading 4'] => strong",
                            "p[style-name='Heading 5'] => strong",
                            "p[style-name='List Paragraph'] => p:fresh > strong"
                        ].join("\n")
                    });
                    let htmlContent = result.value || '';

                    // Convert {{image_url}} placeholders to img tags
                    htmlContent = convertImagePlaceholdersToImgTags(htmlContent);

                    // Log any warnings about content that couldn't be converted
                    if (result.messages && result.messages.length > 0) {
                        console.log('Mammoth conversion warnings:', result.messages);
                    }

                    // Check for optional title filter
                    const rawTitle = req.query.title || jsonBody.title || jsonBody['$title'];
                    const titleFilter = rawTitle ? rawTitle.trim() : undefined;
                    const h1Sections = extractAllH1Sections(htmlContent);

                    let contentToInsert = htmlContent;

                    if (titleFilter) {
                        const matchedSection = h1Sections.find(section =>
                            section.title.toLowerCase().includes(titleFilter.toLowerCase())
                        );

                        if (matchedSection) {
                            contentToInsert = matchedSection.content;
                        } else {
                            return res.status(404).json({
                                status: 'error',
                                message: `Section with title "${titleFilter}" not found`,
                                availableTitles: h1Sections.map(s => s.title)
                            });
                        }
                    }

                    // Apply employee paragraph styles to content
                    contentToInsert = applyEmployeeParagraphStyles(contentToInsert);

                    // Add anchor IDs to H2 tags so links work
                    contentToInsert = addAnchorIdsToH2(contentToInsert);

                    // Add back-to-top button
                    contentToInsert = addBackToTopButton(contentToInsert);

                    // Extract H1 title for ##Title## from the actual content section
                    const h1Match = contentToInsert.match(/<h1[^>]*>(.*?)<\/h1>/i);
                    const h1Title = h1Match ? h1Match[1].replace(/<[^>]*>/g, '').trim() : 'Company Policy';

                    // Get navigation config based on title
                    const navigationKey = Object.keys(NAVIGATION_CONFIG).find(key =>
                        h1Title.toLowerCase().includes(key.toLowerCase())
                    );
                    const navigation = navigationKey ? NAVIGATION_CONFIG[navigationKey] : {
                        sectionNumber: 3,
                        previous: '/sites/InnovationArtsEntertainment/SitePages/HANDBOOK.aspx',
                        next: '/sites/InnovationArtsEntertainment/SitePages/HANDBOOK.aspx'
                    };

                    // Extract H2 sections for #RightSECTION##
                    const h2Sections = extractH2Sections(contentToInsert);
                    const sectionContentsHTML = generateSectionContentsHTML(h2Sections, navigation.sectionNumber);

                    // Escape the HTML for JSON string
                    const escapedMessage = contentToInsert
                        .replace(/\\/g, '\\\\')
                        .replace(/"/g, '\\"')
                        .replace(/\n/g, '\\n');

                    const escapedTitle = h1Title
                        .replace(/\\/g, '\\\\')
                        .replace(/"/g, '\\"')
                        .replace(/\n/g, '\\n');

                    const escapedSectionContents = sectionContentsHTML
                        .replace(/\\/g, '\\\\')
                        .replace(/"/g, '\\"')
                        .replace(/\n/g, '\\n');

                    // Set navigation URLs from config
                    const previousUrl = navigation.previous;
                    const nextUrl = navigation.next;
                    const homeUrl = '/sites/InnovationArtsEntertainment/SitePages/HANDBOOK.aspx';

                    // Generate custom navigation buttons HTML for top (white bg) and bottom (black bg)
                    const navButtonsTopHTML = generateNavigationButtonsHTML(previousUrl, homeUrl, nextUrl, 'top');
                    const navButtonsBottomHTML = generateNavigationButtonsHTML(previousUrl, homeUrl, nextUrl, 'bottom');

                    const escapedNavButtonsTop = navButtonsTopHTML
                        .replace(/\\/g, '\\\\')
                        .replace(/"/g, '\\"')
                        .replace(/\n/g, '\\n');

                    const escapedNavButtonsBottom = navButtonsBottomHTML
                        .replace(/\\/g, '\\\\')
                        .replace(/"/g, '\\"')
                        .replace(/\n/g, '\\n');

                    // Replace placeholders
                    let canvasContent = SECTIONTEMPLATE.CanvasContent1
                        .replace('##LEFT##', escapedMessage)
                        .replace('##TITLE##', escapedTitle)
                        .replace('##RIGHT##', escapedSectionContents)
                        .replace('##NAVBUTTONS_TOP##', escapedNavButtonsTop)
                        .replace('##NAVBUTTONS_BOTTOM##', escapedNavButtonsBottom)
                        .replaceAll('##PREVIOUS_URL##', previousUrl)
                        .replaceAll('##HOME_URL##', homeUrl)
                        .replaceAll('##NEXT_URL##', nextUrl);

                    const response = {
                        ...SECTIONTEMPLATE,
                        CanvasContent1: canvasContent
                    };

                    return res.status(200).json({
                        status: 'success',
                        message: JSON.stringify(response)
                    });
                } else {
                    return res.status(400).json({
                        status: 'error',
                        message: 'Content must be a Word document (.docx)'
                    });
                }
            } catch (decodeError) {
                return res.status(500).json({
                    status: 'error',
                    message: `Error processing content: ${decodeError.message}`
                });
            }
        }

        return res.status(400).json({
            status: 'error',
            message: 'Invalid request format'
        });

    } catch (error) {
        return res.status(500).json({
            status: 'error',
            message: `Error processing request: ${error.message}`
        });
    }
});

// Simple HTML conversion endpoint - just decode and convert to HTML
app.post('/api/html', bodyParser.raw({ type: '*/*', limit: '50mb' }), async (req, res) => {
    try {
        let isJSON = false;
        let jsonBody = null;

        // Try to detect if it's JSON
        if (req.body.length > 0 && (req.body[0] === 0x7B || req.body[0] === 0x5B)) {
            try {
                jsonBody = JSON.parse(req.body.toString('utf8'));
                isJSON = true;
            } catch (e) {
                // Not valid JSON, treat as binary
            }
        }

        let wordBuffer = null;

        // Handle raw binary Word document
        if (!isJSON) {
            if (req.body[0] === 0x50 && req.body[1] === 0x4B) {
                wordBuffer = req.body;
            } else {
                return res.status(400).json({
                    status: 'error',
                    message: 'Unsupported format. Send a Word document (.docx) or JSON with Base64-encoded content.'
                });
            }
        } else {
            // Handle JSON with Base64 content
            const base64Content = jsonBody['$content'] || jsonBody.content;
            if (!base64Content) {
                return res.status(400).json({
                    status: 'error',
                    message: 'Missing $content or content field with Base64-encoded data'
                });
            }
            wordBuffer = Buffer.from(base64Content, 'base64');
        }

        // Convert Word document to HTML
        const result = await mammoth.convertToHtml({
            buffer: wordBuffer
        }, {
            includeDefaultStyleMap: true,
            styleMap: [
                "b => strong",
                "i => em",
                "u => u",
                "p[style-name='Heading 1'] => h1:fresh",
                "p[style-name='Heading 2'] => h2:fresh",
                "p[style-name='Heading 3'] => h3:fresh",
                "p[style-name='Heading 4'] => h4:fresh",
                "p[style-name='Heading 5'] => h5:fresh",
                "p[style-name='List Paragraph'] => p:fresh"
            ].join("\n")
        });

        let htmlContent = result.value || '';

        // Convert {{image_url}} placeholders to img tags
        htmlContent = convertImagePlaceholdersToImgTags(htmlContent);

        // Check for optional title filter
        const rawTitle = req.query.title || jsonBody?.title || jsonBody?.['$title'];
        const titleFilter = rawTitle ? rawTitle.trim() : undefined;

        if (titleFilter) {
            const h1Sections = extractAllH1Sections(htmlContent);
            const matchedSection = h1Sections.find(section =>
                section.title.toLowerCase().includes(titleFilter.toLowerCase())
            );

            if (matchedSection) {
                htmlContent = matchedSection.content;
            } else {
                return res.status(404).json({
                    status: 'error',
                    message: `Section with title "${titleFilter}" not found`,
                    availableTitles: h1Sections.map(s => s.title)
                });
            }
        }

        // Return just the HTML
        return res.status(200).json({
            status: 'success',
            html: htmlContent
        });

    } catch (error) {
        return res.status(500).json({
            status: 'error',
            message: `Error converting to HTML: ${error.message}`
        });
    }
});

// Get all H1 section headers from Word document
app.post('/api/headers', bodyParser.raw({ type: '*/*', limit: '50mb' }), async (req, res) => {
    try {
        let isJSON = false;
        let jsonBody = null;

        // Try to detect if it's JSON
        if (req.body.length > 0 && (req.body[0] === 0x7B || req.body[0] === 0x5B)) {
            try {
                jsonBody = JSON.parse(req.body.toString('utf8'));
                isJSON = true;
            } catch (e) {
                // Not valid JSON, treat as binary
            }
        }

        let wordBuffer = null;

        // Handle raw binary Word document
        if (!isJSON) {
            if (req.body[0] === 0x50 && req.body[1] === 0x4B) {
                wordBuffer = req.body;
            } else {
                return res.status(400).json({
                    status: 'error',
                    message: 'Unsupported format. Send a Word document (.docx) or JSON with Base64-encoded content.'
                });
            }
        } else {
            // Handle JSON with Base64 content
            const base64Content = jsonBody['$content'] || jsonBody.content;
            if (!base64Content) {
                return res.status(400).json({
                    status: 'error',
                    message: 'Missing $content or content field with Base64-encoded data'
                });
            }
            wordBuffer = Buffer.from(base64Content, 'base64');
        }

        // Convert Word document to HTML
        const result = await mammoth.convertToHtml({
            buffer: wordBuffer
        }, {
            includeDefaultStyleMap: true,
            styleMap: [
                "b => strong",
                "i => em",
                "u => u",
                "p[style-name='Heading 1'] => h1:fresh",
                "p[style-name='Heading 2'] => h2:fresh",
                "p[style-name='Heading 3'] => h3:fresh",
                "p[style-name='Heading 4'] => h4:fresh",
                "p[style-name='Heading 5'] => h5:fresh",
                "p[style-name='List Paragraph'] => p:fresh"
            ].join("\n")
        });

        let htmlContent = result.value || '';

        // Extract all H1 sections
        const h1Sections = extractAllH1Sections(htmlContent);

        // Return just the H1 titles
        const h1Titles = h1Sections.map(section => section.title);

        return res.status(200).json({
            status: 'success',
            headers: h1Titles,
            count: h1Titles.length
        });

    } catch (error) {
        return res.status(500).json({
            status: 'error',
            message: `Error extracting headers: ${error.message}`
        });
    }
});

// Start the server
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
    console.log(`POST /api/decode - Process Word doc with SharePoint template`);
    console.log(`POST /api/section - Process Word doc with company policy template (supports ?title= query parameter)`);
    console.log(`POST /api/html - Convert Word doc to HTML (no templates)`);
    console.log(`POST /api/headers - Extract all H1 section headers from Word doc`);
});
