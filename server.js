const express = require('express');
const bodyParser = require('body-parser');
const mammoth = require('mammoth');

const app = express();
const PORT = process.env.PORT || 3500

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

// Helper function to generate section contents HTML with numbered links
function generateSectionContentsHTML(h2Sections, sectionNumber = 3) {
    if (!h2Sections || h2Sections.length === 0) {
        return '';
    }

    let html = '';
    h2Sections.forEach((section, index) => {
        const number = `${sectionNumber}.${index + 1}`;
        html += `<p class=\\"noSpacingAbove spacingBelow lineHeight1_2\\" style=\\"color:rgb(50, 49, 48);margin-left:20px;\\" data-text-type=\\"withSpacing\\"><span class=\\"fontSizeMediumPlus\\"><span lang=\\"EN-US\\" dir=\\"ltr\\"><strong>${sectionNumber}</strong></span></span><a href=\\"#${section.anchorId}\\"><span class=\\"fontSizeMediumPlus\\"><span lang=\\"EN-US\\" dir=\\"ltr\\"><strong>.${index + 1} ${section.heading}</strong></span></span></a></p>`;
    });

    return html;
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

// SharePoint page template with placeholder
const SHAREPOINT_TEMPLATE = {
    "__metadata": {
        "type": "SP.Publishing.SitePage"
    },
    "PageRenderingState": 0,
    "CanvasContent1": "[{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"c12d16a8-30de-46cb-8d44-f22d081aace5\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":1},\"zoneHeight\":550,\"id\":\"3dbcc936-0d52-4658-91f8-dc3cf63925c1\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":0,\"y\":0,\"w\":70,\"h\":7,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<h2 class=\\\"headingSpacingAbove headingSpacingBelow lineHeight1_0\\\"><span class=\\\"fontSizeMega rte-fontscale-font-max\\\">1.1 A&nbsp;Message&nbsp;from the CEO</span></h2>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"c12d16a8-30de-46cb-8d44-f22d081aace5\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":2},\"zoneHeight\":550,\"id\":\"1857d513-039d-4197-8e49-4bd8e6f378c4\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":0,\"y\":6,\"w\":70,\"h\":6,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><span class=\\\"fontSizeXLarge\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\">##CEOMESSAGE##</span></span></p><p class=\\\"noSpacingAbove noSpacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" aria-hidden=\\\"true\\\" data-text-type=\\\"noSpacing\\\">&nbsp;</p><p class=\\\"noSpacingAbove noSpacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" aria-hidden=\\\"true\\\" data-text-type=\\\"noSpacing\\\">&nbsp;</p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"c12d16a8-30de-46cb-8d44-f22d081aace5\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":3},\"zoneHeight\":550,\"id\":\"bd7c16e0-b868-43ce-9742-bd272082eb3c\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":25,\"y\":29,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"bd7c16e0-b868-43ce-9742-bd272082eb3c\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Home\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/HANDBOOK.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Center\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"c12d16a8-30de-46cb-8d44-f22d081aace5\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":4},\"zoneHeight\":550,\"id\":\"2a4fe9c1-c46f-4664-ac1b-b0179329f01b\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":38,\"y\":29,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"2a4fe9c1-c46f-4664-ac1b-b0179329f01b\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Next Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Employment%20Policies%20and%20Practices.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Center\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"0ba22869-c07a-4c44-bccc-521aae856fe8\",\"sectionIndex\":1,\"sectionFactor\":8,\"controlIndex\":1},\"id\":\"48047102-2fb5-4efc-ac96-034716e6f250\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"innerHTML\":\"<p><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\">##COMPANYINFO##</span></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"0ba22869-c07a-4c44-bccc-521aae856fe8\",\"sectionIndex\":2,\"sectionFactor\":4,\"controlIndex\":1},\"id\":\"4696d89e-ec85-4c9b-8ebf-e04f59ab5343\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"innerHTML\":\"<p class=\\\"noSpacingAbove noSpacingBelow\\\" data-text-type=\\\"noSpacing\\\"><span class=\\\"fontSizeXLargePlus\\\"><span class=\\\"fontColorThemeDarker\\\"><strong>Section Contents</strong></span></span></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"0ba22869-c07a-4c44-bccc-521aae856fe8\",\"sectionIndex\":2,\"sectionFactor\":4,\"controlIndex\":2},\"id\":\"da32caa7-5244-4649-8293-35d10e746a80\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"innerHTML\":\"<p class=\\\"lineHeight1_2\\\" style=\\\"margin-left:0px;\\\"><a href=\\\"#a-message-from-the-ceo\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>1.1 </strong>CEO</span> Message</span></a></p><p class=\\\"lineHeight1_2\\\" style=\\\"margin-left:0px;\\\"><a href=\\\"#the-company\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>1.2</strong> The Company</span>&nbsp;</span></a></p><p class=\\\"lineHeight1_2\\\" style=\\\"margin-left:0px;\\\"><a href=\\\"#change-in-policy\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>1.3 </strong>Change in Policy</span>&nbsp;</span></a></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":3,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":1},\"zoneHeight\":1600,\"id\":\"9a58ab09-29e4-4d76-b0f5-8b8a114cb40e\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":36,\"y\":87,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"9a58ab09-29e4-4d76-b0f5-8b8a114cb40e\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Next Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Employment%20Policies%20and%20Practices.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Right\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":3,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":2},\"zoneHeight\":1600,\"id\":\"602be227-1a60-4fd8-a4bd-3f69c27a0fa8\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":3,\"y\":87,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"602be227-1a60-4fd8-a4bd-3f69c27a0fa8\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Home\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/HANDBOOK.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Center\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":3,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":3},\"zoneHeight\":1600,\"id\":\"8048e3e4-74fd-4e6b-9d4f-cfb5649079d8\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":5,\"y\":85,\"w\":40,\"h\":0,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"2161a1c6-db61-4731-b97c-3cdb303f7cbb\",\"instanceId\":\"8048e3e4-74fd-4e6b-9d4f-cfb5649079d8\",\"title\":\"Divider\",\"description\":\"Add a line to divide areas on your page.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{},\"links\":{}},\"dataVersion\":\"1.2\",\"properties\":{\"minimumLayoutWidth\":1,\"length\":100,\"weight\":1},\"containsDynamicDataSource\":false},\"webPartId\":\"2161a1c6-db61-4731-b97c-3cdb303f7cbb\",\"reservedWidth\":688,\"reservedHeight\":1},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isAIGeneratedDescription\":false,\"isDefaultThumbnail\":true,\"isSpellCheckEnabled\":true,\"globalRichTextStylingVersion\":1,\"rtePageSettings\":{\"contentVersion\":5,\"indentationVersion\":2},\"isEmailReady\":true,\"webPartsPageSettings\":{\"isTitleHeadingLevelsEnabled\":true,\"isLowQualityImagePlaceholderEnabled\":true}}},{\"controlType\":14,\"webPartData\":{\"properties\":{\"zoneBackground\":{\"c12d16a8-30de-46cb-8d44-f22d081aace5\":{\"type\":\"image\",\"imageData\":{\"fileName\":\"68234-IAE_Display_28x22_FINAL.jpg\",\"width\":2016,\"height\":720,\"source\":0,\"siteId\":\"44583a03-7429-4de6-9641-ae749e56727f\",\"webId\":\"454df1df-1248-4cb7-a41d-a4e482d69fbb\",\"listId\":\"2ab24772-b566-4fe3-abab-a9edde27e5ea\",\"uniqueId\":\"df6e1993-3198-4e40-9f91-db508a198a82\"},\"overlay\":{\"color\":\"#000000\",\"opacity\":0},\"useLightText\":true},\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\":{\"type\":\"image\",\"imageData\":{\"fileName\":\"template11Image0003.png\",\"width\":144,\"height\":144,\"source\":1},\"overlay\":{\"color\":\"#FFFFFF\",\"opacity\":100},\"useLightText\":false}},\"zoneThemeIndex\":{\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\":{\"lightIndex\":0,\"darkIndex\":0}}},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{\"zoneBackground.c12d16a8-30de-46cb-8d44-f22d081aace5.imageData.url\":\"/sites/InnovationArtsEntertainment/SiteAssets/SitePages/VolunteerCenter(1)/68234-IAE_Display_28x22_FINAL.jpg\",\"zoneBackground.54f8a2d2-b737-4a6e-96a2-54ae0a01c7da.imageData.url\":\"https://cdn.hubblecontent.osi.office.net/m365content/publish/4b99b582-026b-4490-ab73-b0091ab721cc/image.png\"},\"links\":{}},\"dataVersion\":\"1.0\"}}]"
}

// SharePoint section template with ##MESSAGE## placeholder
const SHAREPOINT_SECTION_TEMPLATE = {
    "__metadata": {
        "type": "SP.Publishing.SitePage"
    },
    "PageRenderingState": 0,
    "CanvasContent1": "[{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"56eb74b5-aa11-4125-b709-c075b5fb9b35\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":1},\"zoneHeight\":241,\"id\":\"3dbcc936-0d52-4658-91f8-dc3cf63925c1\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":0,\"y\":2,\"w\":70,\"h\":7,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;text-align:center;\\\" data-text-type=\\\"withSpacing\\\"><span class=\\\"fontSizeMega rte-fontscale-font-max\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>Employment Policies and Practices</strong></span><strong>&nbsp;</strong></span></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"56eb74b5-aa11-4125-b709-c075b5fb9b35\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":2},\"zoneHeight\":241,\"id\":\"064d4a8b-0c85-415e-b306-31b3f5d411d4\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":30,\"y\":10,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"064d4a8b-0c85-415e-b306-31b3f5d411d4\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Home\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/HANDBOOK.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Center\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"56eb74b5-aa11-4125-b709-c075b5fb9b35\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":3},\"zoneHeight\":241,\"id\":\"abf4a27b-49df-4fbf-82d9-868ef2a43fc5\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":17,\"y\":10,\"w\":11,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"abf4a27b-49df-4fbf-82d9-868ef2a43fc5\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Previous Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Introduction.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Left\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":189,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"56eb74b5-aa11-4125-b709-c075b5fb9b35\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":4},\"zoneHeight\":241,\"id\":\"048cb962-5cba-402b-892d-d9dc44421c1d\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":42,\"y\":10,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"048cb962-5cba-402b-892d-d9dc44421c1d\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Next Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Company%20Policies%20and%20Practices.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Right\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":1},\"zoneHeight\":3612,\"id\":\"48047102-2fb5-4efc-ac96-034716e6f250\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":4,\"y\":1,\"w\":44,\"h\":3,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<p><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\">##MESSAGE##</span></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":2},\"zoneHeight\":3612,\"id\":\"4696d89e-ec85-4c9b-8ebf-e04f59ab5343\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":52,\"y\":1,\"w\":14,\"h\":3,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<p class=\\\"noSpacingAbove noSpacingBelow\\\" data-text-type=\\\"noSpacing\\\"><span class=\\\"fontSizeXLargePlus\\\"><span class=\\\"fontColorThemeDarker\\\"><strong>Section Contents</strong></span></span></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":3},\"zoneHeight\":3612,\"id\":\"da32caa7-5244-4649-8293-35d10e746a80\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":52,\"y\":3,\"w\":17,\"h\":16,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"#equal-employment-opportunity-%28illinois%29\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>2.1</strong> Equal Employment Opportunity (Illinois employees)</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"#anti-discrimination-and-anti-sexual-harassment-policies\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>2.2</strong> Anti-Discrimination and Anti-Sexual Harassment Policies</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"#at-will-employment\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>2.3</strong> At-Will Employment</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"#employee-classifications\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>2.4</strong> Employee Classifications (exempt, full-time)</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"#confidentiality-and-trade-secrets\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>2.5</strong> Confidentiality and Trade Secrets</span>&nbsp;</span></a></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":4},\"zoneHeight\":3612,\"id\":\"9a58ab09-29e4-4d76-b0f5-8b8a114cb40e\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":38,\"y\":208,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"9a58ab09-29e4-4d76-b0f5-8b8a114cb40e\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Next Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Company%20Policies%20and%20Practices.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Right\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":5},\"zoneHeight\":3612,\"id\":\"602be227-1a60-4fd8-a4bd-3f69c27a0fa8\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":21,\"y\":208,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"602be227-1a60-4fd8-a4bd-3f69c27a0fa8\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Home\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/HANDBOOK.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Center\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":6},\"zoneHeight\":3612,\"id\":\"22a99bb7-d8fb-4e17-96cf-2c0d90d5988f\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":5,\"y\":208,\"w\":11,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"22a99bb7-d8fb-4e17-96cf-2c0d90d5988f\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Previous Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Introduction.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Left\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":189,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":7},\"zoneHeight\":3612,\"id\":\"8048e3e4-74fd-4e6b-9d4f-cfb5649079d8\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":6,\"y\":186,\"w\":40,\"h\":0,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"2161a1c6-db61-4731-b97c-3cdb303f7cbb\",\"instanceId\":\"8048e3e4-74fd-4e6b-9d4f-cfb5649079d8\",\"title\":\"Divider\",\"description\":\"Add a line to divide areas on your page.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{},\"links\":{}},\"dataVersion\":\"1.2\",\"properties\":{\"minimumLayoutWidth\":1,\"length\":100,\"weight\":1},\"containsDynamicDataSource\":false},\"webPartId\":\"2161a1c6-db61-4731-b97c-3cdb303f7cbb\",\"reservedWidth\":688,\"reservedHeight\":1},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isAIGen eratedDescription\":false,\"isDefaultThumbnail\":true,\"isSpellCheckEnabled\":true,\"globalRichTextStylingVersion\":1,\"rtePageSettings\":{\"contentVersion\":5,\"indentationVersion\":2},\"isEmailReady\":true,\"webPartsPageSettings\":{\"isTitleHeadingLevelsEnabled\":true,\"isLowQualityImagePlaceholderEnabled\":true}}},{\"controlType\":14,\"webPartData\":{\"properties\":{\"zoneBackground\":{\"56eb74b5-aa11-4125-b709-c075b5fb9b35\":{\"type\":\"image\",\"imageData\":{\"fileName\":\"68234-IAE_Display_28x22_FINAL.jpg\",\"width\":2016,\"height\":720,\"source\":0,\"siteId\":\"44583a03-7429-4de6-9641-ae749e56727f\",\"webId\":\"454df1df-1248-4cb7-a41d-a4e482d69fbb\",\"listId\":\"2ab24772-b566-4fe3-abab-a9edde27e5ea\",\"uniqueId\":\"df6e1993-3198-4e40-9f91-db508a198a82\"},\"overlay\":{\"color\":\"#000000\",\"opacity\":0},\"useLightText\":true},\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\":{\"type\":\"image\",\"imageData\":{\"fileName\":\"template11Image0003.png\",\"width\":144,\"height\":144,\"source\":1},\"overlay\":{\"color\":\"#FFFFFF\",\"opacity\":100},\"useLightText\":false}},\"zoneThemeIndex\":{\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\":{\"lightIndex\":0,\"darkIndex\":0}}},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{\"zoneBackground.56eb74b5-aa11-4125-b709-c075b5fb9b35.imageData.url\":\"/sites/InnovationArtsEntertainment/SiteAssets/SitePages/VolunteerCenter(1)/68234-IAE_Display_28x22_FINAL.jpg\",\"zoneBackground.54f8a2d2-b737-4a6e-96a2-54ae0a01c7da.imageData.url\":\"https://cdn.hubblecontent.osi.office.net/m365content/publish/4b99b582-026b-4490-ab73-b0091ab721cc/image.png\"},\"links\":{}},\"dataVersion\":\"1.0\"}}]"
}

// Company Policy template with ##MESSAGE## placeholder
const COMPANYPOLICYTEMPLATE = {
    "__metadata": {
        "type": "SP.Publishing.SitePage"
    },
    "PageRenderingState": 0,
    "CanvasContent1": "[{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"f374e17d-f830-40c9-859a-b641fed7e09c\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":1},\"zoneHeight\":344,\"id\":\"3dbcc936-0d52-4658-91f8-dc3cf63925c1\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":0,\"y\":3,\"w\":70,\"h\":13,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<h2 class=\\\"headingSpacingAbove headingSpacingBelow lineHeight2_4\\\" style=\\\"text-align:center;\\\"><span class=\\\"fontSizeMega rte-fontscale-font-max\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\">Company Policies and Practices</span>&nbsp;</span></h2><p class=\\\"noSpacingAbove spacingBelow lineHeight1_3\\\" style=\\\"margin-left:0px;text-align:center;\\\" data-text-type=\\\"withSpacing\\\"><span class=\\\"fontSizeXLarge\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\">The Innovation Arts &amp; Entertainment Employee Handbook is intended to provide&nbsp;employees with a clear understanding of company policies and procedures related to a variety of&nbsp;business areas</span>&nbsp;</span></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;text-align:center;\\\" data-text-type=\\\"withSpacing\\\"><span class=\\\"fontSizeXLarge\\\">&nbsp;</span></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"f374e17d-f830-40c9-859a-b641fed7e09c\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":2},\"zoneHeight\":344,\"id\":\"17459e36-060a-4ca2-99ce-ffbd8d4b59c2\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":19,\"y\":15,\"w\":10,\"h\":3}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"17459e36-060a-4ca2-99ce-ffbd8d4b59c2\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Previous Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Employment%20Policies%20and%20Practices.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Left\",\"minimumLayoutWidth\":9,\"isDynamicWidthEnabled\":false},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":172,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"f374e17d-f830-40c9-859a-b641fed7e09c\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":3},\"zoneHeight\":344,\"id\":\"ff4e4cae-e93d-4692-af6f-eafd6f0a6083\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":31,\"y\":15,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"ff4e4cae-e93d-4692-af6f-eafd6f0a6083\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Home\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/HANDBOOK.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Center\",\"minimumLayoutWidth\":9,\"isDynamicWidthEnabled\":false},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"f374e17d-f830-40c9-859a-b641fed7e09c\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":4},\"zoneHeight\":344,\"id\":\"3d2aa629-fe8c-4209-84bb-a6382cee019c\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":41,\"y\":15,\"w\":9,\"h\":3,\"dataVersion\":\" 1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"3d2aa629-fe8c-4209-84bb-a6382cee019c\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Next Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Compensation%20and%20Benefits.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Right\",\"minimumLayoutWidth\":9,\"isDynamicWidthEnabled\":false},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"2701fa25-eb7f-4ead-ba21-40f9a35a8fb5\",\"sectionIndex\":1,\"sectionFactor\":8,\"controlIndex\":1},\"id\":\"ca89cf71-9ce3-4c01-8f23-ccb08babf9fb\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"innerHTML\":\"<p>&nbsp;</p><h2>Company Policies and Practices</h2><p>The Innovation Arts &amp; Entertainment Employee Handbook is intended to provide&nbsp;employees with a clear understanding of company policies and procedures.</p><h2>Workweek and Work Schedules</h2><p>Given the nature of the live entertainment business,&nbsp;Innovation Arts &amp; Entertainment’s&nbsp;workweek runs from 12:00 AM on Sundays through 11:59 PM on Saturdays.</p><p>Standard office hours for full-time employees working in our Chicago office&nbsp;are as follows:&nbsp;</p><p>Monday 8:30am - 5:30pm&nbsp;Central</p><p>Tuesday 8:30am - 5:30pm&nbsp;Central</p><p>Wednesday 8:30am - 5:30pm&nbsp;Central</p><p>Thursday 8:30am - 5:30pm&nbsp;Central</p><p>Friday 9:00am - 1:00pm&nbsp;Central*</p><p>* <i>Employees are given the option to work from home on Fridays unless, due to business circumstances, in-office attendance is required. Refer to the Company Work From Home Policy in this handbook for further details.</i></p><p>Remote employees outside the&nbsp;Central time zone are encouraged to adapt accordingly.&nbsp;</p><p>IAE recognizes&nbsp;that&nbsp;our business&nbsp;requires&nbsp;that standard office hours be stretched in either direction and/or over weekends. Whether&nbsp;meeting an&nbsp;important deadline, scheduling considerations, travel,&nbsp;or working onsite at an IAE engagement, there&nbsp;will be times&nbsp;when working beyond regularly set hours is&nbsp;required, and advance notice&nbsp;of these occurrences&nbsp;may not be possible.</p><h2>Access Key Card Policy&nbsp;</h2><p>All employees&nbsp;are responsible for maintaining security at Innovation Arts &amp; Entertainment. Employees may receive a key, key fob, or&nbsp;access card to enter&nbsp;the IAE office. Employees must use keys, fobs&nbsp;or access cards&nbsp;adhering to&nbsp;the following requirements:</p><p><strong>Employees should never loan or duplicate a key, fob&nbsp;or access card or unlock&nbsp;the&nbsp;building or&nbsp;office&nbsp;for another person unless the employee knows that individual&nbsp;is authorized to&nbsp;enter.&nbsp;</strong></p><p><strong>Employees should report lost, stolen, or misplaced keys or access cards to employees’ supervisors&nbsp;immediately. Employees&nbsp;are responsible for&nbsp;the cost of replacement. All keys and access cards must be returned to&nbsp;IAE&nbsp;upon separation.&nbsp;</strong></p><h2>Use of Company Property</h2><p>All&nbsp;property owned by&nbsp;Innovation Arts &amp; Entertainment&nbsp;must be&nbsp;maintained&nbsp;in good working order and&nbsp;in accordance with&nbsp;IAE’s&nbsp;rules and regulations.&nbsp;This includes property in IAE’s office and property in the temporary possession of employees, such as laptops and other electronic equipment.</p><p><strong>Employees must maintain proper vigilance and care while in the possession of company property both inside the office, and especially when removing this property from the office for travel purposes. Loss of or damage to company property requires immediate reporting of the matter to the Director of Operations, and the employee's filing of necessary paperwork for police reports and insurance claims in a timely way.</strong></p><p><strong>IAE&nbsp;may inspect all property to ensure compliance with its rules and regulations without giving notice to employees or in employees’ absence.</strong></p><p><strong>Prior authorization must be obtained from the Director of Operations before any&nbsp;IAE&nbsp;property may be removed from the&nbsp;premises.</strong></p><p><strong>Employees&nbsp;in possession&nbsp;of company property acknowledge and accept that said&nbsp;company property&nbsp;will contain&nbsp;tracking devices and will have software that will monitor the property's location and health.</strong></p><p><strong>Upon termination of employment or upon request by&nbsp;IAE, employees must return all&nbsp;IAE&nbsp;property&nbsp;in their possession&nbsp;immediately.&nbsp;</strong></p><h2>Technology Systems and Software&nbsp;</h2><p>Innovation Arts &amp; Entertainment&nbsp;provides employees with access to the&nbsp;company’s&nbsp;computer networks,&nbsp;communication,&nbsp;software,&nbsp;and technology systems to&nbsp;assist&nbsp;employees in conducting&nbsp;IAE&nbsp;business. Everything created, received, sent, or stored in these systems is the&nbsp;sole&nbsp;property of the&nbsp;IAE.</p><p>All&nbsp;IAE&nbsp;policies apply to the conduct of employees on the internet and when using&nbsp;IAE communication,&nbsp;software,&nbsp;and technology systems. The display of any kind of unauthorized sexually explicit material on any&nbsp;IAE&nbsp;system or through any electronic communication&nbsp;method&nbsp;operated&nbsp;by&nbsp;IAE&nbsp;is a violation of the&nbsp;company’s policy against sexual harassment. Employees who are aware of the misuse of technology systems by other employees should report the misuse to a company director&nbsp;immediately.</p><h2>Technical Support</h2><p>IAE provides tech support for all employees, including Remote Employees, via the company’s IT Services provider, Genuity.</p><p>Any employee experiencing technical difficulties should take the following steps:</p><p><strong>Shut all programs down and restart your computer. See if the problem continues.</strong></p><p><strong>If the problem recurs, email the Director of Operations with clear, concise description of the problem.</strong></p><p><strong>Our first effort will attempt to resolve the problem using internal experience and resources. More troublesome issues will be escalated by our Director of Operations and resolved by our Managed IT Services company.</strong></p><h2>Expectation of Privacy&nbsp;</h2><p>Employees should&nbsp;have no expectation of privacy&nbsp;regarding&nbsp;their&nbsp;use of company internet and technology systems&nbsp;and should not use IAE systems for information they wish to keep private.&nbsp;The&nbsp;company reserves the right to&nbsp;monitor&nbsp;and access all&nbsp;company&nbsp;property.&nbsp;</p><h2>Cell Phone Policy</h2><p>Innovation Arts &amp; Entertainment&nbsp;seeks&nbsp;to provide a working environment where employees can focus,&nbsp;with limited distractions. Use of personal cell phones during work hours interferes with employees’ productivity and is discouraged.&nbsp;IAE&nbsp;acknowledges that employees have personal commitments and responsibilities outside of work and expects employees to&nbsp;use&nbsp;professional judgment&nbsp;regarding&nbsp;use of&nbsp;personal cell phones during work hours.&nbsp;</p><p>IAE&nbsp;is not liable for loss or damage to employees’ personal cell phones brought into the workplace.&nbsp;</p><p>Viewing, downloading, or uploading illegal or obscene material while using an IAE internet connection, or in an IAE workplace, is prohibited.&nbsp;</p><p>Employees must&nbsp;comply with&nbsp;all applicable laws&nbsp;regarding&nbsp;the use of cell phones while driving. Employees should never use cell phones to conduct business while driving. Employees must use a hands-free device to make or receive calls while driving, if&nbsp;permitted&nbsp;by law.&nbsp;</p><h2>Social Media Policy&nbsp;</h2><p>While using social media platforms, employees are expected to&nbsp;maintain&nbsp;a professional tone and demeanor.&nbsp;Employees must understand that they are&nbsp;representing&nbsp;Innovation Arts &amp; Entertainment,&nbsp;even in&nbsp;their&nbsp;personal capacity, and follow these guidelines:&nbsp;</p><p><strong>Do not&nbsp;disclose&nbsp;any confidential, proprietary, or sensitive information about&nbsp;IAE, its clients, partners, or&nbsp;associates&nbsp;on social media, including&nbsp;financial information,&nbsp;future plans, and any non-public business matters.&nbsp;</strong></p><p><strong>Always engage in respectful and constructive communication on social media. Avoid using offensive, discriminatory, or defamatory language or content.&nbsp;</strong></p><p><strong>On any personal&nbsp;account, clearly distinguish between&nbsp;personal opinion&nbsp;and those of&nbsp;IAE. If you choose to&nbsp;identify&nbsp;yourself as an employee, include a disclaimer that your views are your own and not reflective of the company's stance.</strong></p><p><strong>Do not register for personal accounts with&nbsp;an @innovationae.com email address.&nbsp;</strong></p><p><strong>Respect the privacy of colleagues, clients, and partners. Do not post photos, videos, or information about them without their explicit consent.&nbsp;</strong></p><p><strong>Do not engage in or tolerate any form of harassment, bullying, or cyberbullying on social media. Report any such behavior to the&nbsp;appropriate internal&nbsp;channels.&nbsp;</strong></p><p><strong>Adhere to all laws and regulations related to copyright, intellectual property, privacy, and defamation when posting content on social media.&nbsp;</strong></p><p><strong>While personal use of social media during breaks is&nbsp;permitted, avoid excessive use that may interfere with your work responsibilities.</strong></p><p>Clear boundaries between personal and professional interactions can enable all employees to feel safe and comfortable in the workplace.&nbsp;</p><p>Employees should treat other employees with respect and not flirt&nbsp;in&nbsp;the workplace. Employees should report unwanted advances or flirting to their supervisor or Human Resources. Sexual harassment is&nbsp;prohibited. For more information on sexual harassment, see the Sexual Harassment policy.&nbsp;</p><p>Dating employees should not allow their personal relationship to disrupt the workplace.&nbsp;During work hours and while at the IAE office, dating employees should behave professionally and not engage in physical contact or personal conversations that would be inappropriate in the workplace.&nbsp;Dating employees should remain productive, focused, and committed to work.&nbsp;</p><p>If employees stop dating, they should&nbsp;maintain&nbsp;a professional relationship. Employees should not disparage each other,&nbsp;disclose&nbsp;details of their relationship, or engage in any other actions that disrupt the workplace.&nbsp;</p><p>Dating employees who work in a direct supervisory relationship or in job positions in which a conflict of interest could arise must&nbsp;immediately&nbsp;disclose&nbsp;their relationship to their supervisor or&nbsp;Director of Operations.</p><p>Additionally, any executive, manager, or other influential company official who is involved in a relationship with a co-worker must&nbsp;disclose&nbsp;their relationship to&nbsp;the Director of Operations.&nbsp;IAE&nbsp;will review the circumstances of the relationship and notify the employees of any necessary actions.</p><h2>Office Voicemail Policy&nbsp;</h2><p>Employees&nbsp;are required to make use of the Company voicemail system to receive incoming voice messages.&nbsp;Employee outgoing messages are pre-programmed with a message using text-to-speech technology,&nbsp;stating:&nbsp;</p><p>“{First Name} is currently unavailable. Please leave him/her/them your name, number, and the reason for your call and he/she/they will respond as soon as possible.”&nbsp;</p><p>Employees may record their own outgoing messages following a similar guideline,&nbsp;following IAE's Voicemail Setup Procedure</p><p>Voicemail Setup Procedure&nbsp;</p><p>Employees must record their own&nbsp;outgoing&nbsp;voicemail message, including&nbsp;out-of-office messages&nbsp;referenced in <i>Out-of-Office Policy</i>,&nbsp;may&nbsp;do so by following these steps:&nbsp;</p><p><strong>Press the “Message” button on&nbsp;desk&nbsp;phone for voicemail.&nbsp;</strong></p><p><strong>Enter&nbsp;personal&nbsp;pin&nbsp;number&nbsp;followed by #&nbsp;</strong></p><p><strong>Press 6 for greetings.&nbsp;</strong></p><p><strong>Press&nbsp;1 to record new greeting.&nbsp;</strong></p><p><strong>Press 2 to review&nbsp;recorded messages.&nbsp;</strong></p><p><strong>Press 3 to select the&nbsp;“situational” greeting for use.&nbsp;</strong></p><p><strong>Employees unaware of their pin number should contact IAE’s Director of Operations to receive a new one.&nbsp;</strong></p><h2>Email Signature&nbsp;Policy</h2><p>Employees&nbsp;are required to&nbsp;employ IAE’s company email signature format on all outgoing&nbsp;emails. IAE&nbsp;signatures are&nbsp;to&nbsp;be&nbsp;displayed&nbsp;on&nbsp;both&nbsp;your&nbsp;initial&nbsp;email and replies – on both desktop and mobile devices – using&nbsp;the&nbsp;IAE’s&nbsp;<a href=\\\"#_Email_Signature_Format\\\">Email Signature Format Procedure.</a></p><h2>Out-of-Office Policy&nbsp;</h2><p>When out of the office for an extended period (e.g.,&nbsp;PTO), employees are required to set up the following notifications to alert people attempting to contact them:</p><p>Email</p><p>Create an out-of-office email auto-response message&nbsp;that alerts the sender that you are out of the office, the date of their return and,&nbsp;for time-sensitive matters, provide an alternative&nbsp;IAE employee with their phone and email information to contact in their absence.</p><p>Voicemail</p><p>Record an outgoing&nbsp;voicemail&nbsp;message that informs&nbsp;callers that you are out of the office, the date of their return and,&nbsp;for time-sensitive matters, provide an alternative&nbsp;IAE employee&nbsp;and extension to contact in their absence.</p><h2>Distribution Lists,&nbsp;Groups, and Teams&nbsp;</h2><p>Employees are not&nbsp;permitted&nbsp;to add new Distribution&nbsp;Lists, Groups&nbsp;or&nbsp;Teams&nbsp;in&nbsp;Microsoft Outlook and Teams. If the need for a new list,&nbsp;group,&nbsp;or team arises,&nbsp;employees should&nbsp;submit&nbsp;an email request to IAE’s Director of Operations detailing the name&nbsp;and&nbsp;members to be associated with the list,&nbsp;group,&nbsp;or team.&nbsp;</p><h2>Labeling and Coding Policy&nbsp;</h2><p>To promote efficiency of interdepartmental workflow and the accuracy of work product, employees&nbsp;are required to&nbsp;strictly adhere to the&nbsp;conventions detailed in IAE’s&nbsp;<a href=\\\"#_Labeling_and_Coding\\\">Labeling and Coding Procedure</a>.</p><h2>Document Management Policy</h2><p>Document management is the process of managing the creation, review, approval, distribution, and revision of documents within the Innovation Arts &amp; Entertainment environment.</p><p>Effective document management is essential for efficient inter-office communication, sharing of information, compliance, and quality. By standardizing our digital systems and processes, IAE ensures that our documents are accurate, consistent, up-to-date, and easily accessible to authorized personnel – and authorized personnel only.</p><p>Employees should save documents to the correct location the first time. This reduces the amount of work involved when identifying records and organizing storage drives later. Each drive has different documents that are appropriate and inappropriate. Employees should not save documents of a personal nature (vacation photos, personal emails, etc.) on any shared storage. These items make it more difficult to find relevant information, create a significant burden for IT systems to back up, and use resources and storage for information that is unrelated to work. Additionally, these items could potentially become subject to a public records request or e‐discovery. In addition to saving documents in the correct location the first time, users should avoid placing duplicate documents on a shared storage. When a document is ready for placement on shared storage, place it in the appropriate existing folder or create the requisite folder on the company’s servers.</p><p>The creation of new documents requires discipline and consideration of your fellow employees who may rely upon the information and documents you create. Each new file needs to be named according to company standards and stored in its right place. Accordingly, IAE’s&nbsp;Labeling and Coding Policy&nbsp;and&nbsp;Labeling and Coding Procedure, including naming conventions and IAE Cloud Server structure and filing system are integral parts of IAE’s document management and internal communication.</p><p>Within their respective departments, Supervisors are authorized and responsible for the review and approval of documents before “publication” and distribution, deciding to whom documents are to be distributed, determining when revisions are necessary, and ensuring that documents are filed and stored in the appropriate folders on IAE’s Cloud Server.</p><p>External File Sharing</p><p>External File Sharing refers to the distribution of digital files to people outside of IAE.</p><p>The distribution of documents requires diligence to ensure that no confidential or sensitive information is shared with people outside the company. Be cautious when sharing and setting permissions for data stored in the cloud. Limit file sharing to only those with a legitimate need to know for purposes of conducting business with IAE.</p><p>For this reason, the company mandates that no attachments may be added to emails. The only acceptable process for distribution of files to people outside of IAE is via the “Share” function built into OneDrive and SharePoint.</p><p>This valuable tool allows access to shared files to be limited to “view only” or “view without being able to download” or “editable.” Most importantly, the share function allows for the access to the file to be limited to a specific period or revoked if you make a mistake.</p><p>When sharing files, make sure to share only with specific people, and avoid sharing files with the default \\\"Anyone with the link can edit\\\" permission, which allows the person you shared that file with to further share the file with anyone for the same level of access. Never use anonymous guest links, especially when sharing confidential or restricted data.</p><p>Employees should discuss document and distribution circumstances with their supervisor to seek guidance and approval.</p><p>Additional Precautions for External Sharing</p><p><strong>Never share access to a folder, as other employees who may put files in that folder may not know that the folder is being shared.</strong></p><p><strong>Never attach a Microsoft Excel file to any email.</strong></p><p><strong>When sharing an Excel file using the Share function be sure to limit access to only the sheet(s) in a workbook that you wish to share. Never share the entire workbook unless you know that you want the recipient to have access to all its contents.</strong></p><h2>Expense Reporting&nbsp;and Reimbursement&nbsp;Policy</h2><p>At&nbsp;Innovation Arts &amp; Entertainment, every company expense must advance our goals and never fund waste. This policy translates that principle into clear guardrails for every employee. If a purchase clearly supports IAE objectives and can be justified by an employee to their supervisor,&nbsp;it is&nbsp;likely&nbsp;in&nbsp;policy.&nbsp;Deeming legitimate expenses for the performance of the employee’s job will be&nbsp;determined&nbsp;by IAE with reasonable, good faith discretion.&nbsp;If uncertain,&nbsp;employees should seek clarity and&nbsp;ask for approval&nbsp;from their supervisor&nbsp;before&nbsp;making expenditures.&nbsp;</p><p>Legally, every expense incurred by all company personnel must qualify as a legitimate business expense. In pursuit of compliance with generally accepted accounting principles (GAAP) and the IRS Code, all expenses must be coded to let our Finance operation account for the purpose of each expenditure. Ramp (app.ramp.com) is IAE’s travel booking and expense reporting tool that automates several tasks, including expense reports, making it easier for employees to manage their own expenses.</p><p>It is not necessary to complete formal monthly expense reports; however, to&nbsp;optimize&nbsp;reporting accuracy and efficiency,&nbsp;employees&nbsp;are required to&nbsp;code their expenses and other financial transactions in Ramp accurately, with photo or PDF receipts attached, within&nbsp;48 hours&nbsp;of the transactions.&nbsp;</p><p>Full step-by-step instructions for entering invoice and expense transactions in Ramp are detailed in IAE’s&nbsp;<a href=\\\"#_Bill_Payment_and\\\">Bill Payment and Expense Reporting</a> Procedure.&nbsp;Additionally, full step-by-step instructions for booking travel are detailed in IAE's&nbsp;Tr<a href=\\\"#_Travel_Booking_and\\\">avel Booking and Expense Reporting</a> Procedure.</p><p>Ramp (app.ramp.com) is IAE’s travel booking and expense reporting tool of choice and an all-in-one financial operation and spend management platform that integrates corporate credit cards, expense management, and accounts payable to help businesses save time and money.&nbsp;</p><p>When&nbsp;paying for&nbsp;IAE business expenses, including travel expenses, employees&nbsp;are required to&nbsp;use their IAE corporate credit card (issued by Ramp).&nbsp;</p><h2>Work From Home&nbsp;Policy</h2><p>To take advantage of work from home (WFH) opportunities, employees must have a fully equipped workspace with a computer.&nbsp;IAE does not provide equipment for home offices.&nbsp;Employees without their own computer systems capable of managing their work from home are required to work from our office. Staff may inquire about IAE surplus computer equipment and associated&nbsp;costs&nbsp;to own.</p><p>When working from home, employees must be&nbsp;available for&nbsp;the entire workday, just like when&nbsp;they are&nbsp;in the IAE office. Employees&nbsp;are responsible for&nbsp;performing all required work to execute IAE business. If, for whatever reason, working from home does not align with appropriately executing IAE business, including attending in-person meetings, employees are expected to come into the IAE office to perform their work.&nbsp;</p><p>Consistent communication between employees and direct supervisors is critical to ensure the best possible result for&nbsp;the execution of&nbsp;IAE business and the establishment of balanced schedules&nbsp;for employees. As such, in-person meetings may be required on days where Work From Home is scheduled. Traveling away from home for personal reasons during WFH periods is not permitted and will result in the reassignment of those hours to Personal Time Off for the employee.</p><h2>Remote Worker Policy</h2><p>Innovation Arts &amp; Entertainment’s remote worker policy establishes guidelines for all employees who work from home on a regular full-time basis, starting with and including the guidelines detailed in the Company <a href=\\\"#_Work_From_Home\\\">Work From Home Policy</a>.</p><p>The key element of ensuring remote employee success is communication with an employee’s supervisor and relevant team members.</p><p>Basic Obligations</p><p><strong>Be fully available via video conferencing and telephone calls during regularly scheduled work hours.</strong></p><p><strong>Respond to critical emails in a timely manner (1-2 hours).</strong></p><p><strong>Check-in regularly throughout a workday with their supervisor via Teams chat.</strong></p><p><strong>Meet weekly with their supervisor 1-on-1 by phone or Teams video conference to discuss ongoing work and progress.</strong></p><p><strong>Provide written summaries of ongoing work, as requested.</strong></p><p><strong>Request PTO when needing to step away from work during the workweek.</strong></p><p><strong>Remote employees are expected to take proper measures to ensure the protection of company data, proprietary information and assets.</strong></p><p>Travel requirements</p><p>Occasionally, remote employees will be required to attend company meetings in person on occasion, as requested. Travel expenses will be reimbursed as outlined in IAE’s&nbsp;Travel Policy. All travel plans must be submitted to your supervisor prior to booking, executed only after receiving written approval.</p><p>Non-adherence</p><p>Failure to fulfill work requirements or adhere to policies and procedures while working remotely may result in disciplinary action, termination of the remote work agreement, or termination of employment.</p><h2>Credit Card Policy&nbsp;</h2><p>The following&nbsp;guidelines&nbsp;are&nbsp;for&nbsp;employees authorized to use IAE corporate&nbsp;credit cards and are provided&nbsp;to ensure&nbsp;their&nbsp;appropriate business&nbsp;use.</p><p>Authorized Use&nbsp;</p><p>Company credit cards are for payment of legitimate business expenses only. This may include travel, lodging, meals, and other approved expenses that relate to conducting IAE business. Personal use of company credit cards is&nbsp;prohibited.</p><p>Cardholder Responsibilities&nbsp;</p><p>Employees&nbsp;are responsible for&nbsp;securing their company credit card to safeguard against unauthorized use.&nbsp;In the event of&nbsp;a missing or stolen credit card, employees should&nbsp;immediately&nbsp;notify IAE’s Director of Operations. Violation of the&nbsp;cardholder's&nbsp;responsibilities may result in disciplinary action, including suspension of credit card usage permission, or termination of employment.&nbsp;</p><p>The use of any company credit card is spending company funds. As such, all employees in possession of a credit card bear the responsibility for disclosing exact details for every purchase by entering the details into the Ramp app according to official coding requirements as described herein.</p><h2>Travel Policy&nbsp;</h2><p>Booking work-related travel takes two forms at Innovation Arts &amp; Entertainment: 1) booking for a group (that may or may not include yourself) and 2) booking for yourself alone.</p><p>Booking Group Travel</p><p>When booking ground transportation or hotels for non-company personnel, employees are required to use our travel agent, Valise Travel, owned and operated by Mitch Levine. Mitch can be reached at <a href=\\\"mailto:mitch@valise.live\\\">mitch@valise.live</a> or (412) 996-9073. Mitch has an IAE Company Visa, so you do not need to make payments.</p><p>To book travel with Valise Travel, employees should be prepared to provide:</p><p>For Hotels</p><p><strong>The number of rooms of each type (King, queens, double beds)</strong></p><p><strong>Check-in and check-out dates</strong></p><p><strong>You should not provide the list of guests until three weeks before the reservation.</strong></p><p>For Ground Transportation</p><p><strong>Provide the date, time, vehicle type, and the location of pickup and drop off.</strong></p><p>Booking Travel for Yourself</p><p>Air Travel</p><p>Employees&nbsp;are required to&nbsp;book air travel via Ramp. Full step-by-step instructions for booking travel are</p><p>detailed in IAE's&nbsp;Travel Booking and Expense Reporting Procedure. IAE Travel Policy booking parameters</p><p>and spend limits&nbsp;established&nbsp;in Ramp settings are as follows:</p><p><strong>Policies &amp; Regulations</strong></p><p>No&nbsp;maximum&nbsp;price per flight leg is set, however adherence to the following booking parameters is&nbsp;required:&nbsp;</p><p><strong>Book non-refundable flights only.&nbsp;</strong></p><p><strong>Book flights ≥&nbsp;14 days&nbsp;in advance. Exceptions require supervisor review.&nbsp;</strong></p><p><strong>Book Economy&nbsp;class&nbsp;for flights&nbsp;less than&nbsp;six&nbsp;hours&nbsp;in duration.&nbsp;</strong></p><p><strong>Premium Economy&nbsp;class&nbsp;is allowed for flights&nbsp;greater than&nbsp;six&nbsp;hours&nbsp;in duration.&nbsp;</strong></p><p><strong>One checked bag and in-flight Wi-Fi are allowed.&nbsp;</strong></p><p><strong>First-class airfare and airport lounge access are prohibited.</strong></p><p><strong>Airline&nbsp;Loyalty Programs&nbsp;</strong></p><p>Airline Loyalty Points earned while traveling on IAE business are kept by employees, and are&nbsp;a fantastic way&nbsp;to receive&nbsp;perks, including seat upgrades, free baggage&nbsp;check,&nbsp;and premium seat selection. Employees are encouraged to enroll in airline loyalty programs, including with those airlines listed below:&nbsp;</p><p><strong>United</strong></p><p><strong>American</strong></p><p><strong>Delta</strong></p><p><strong>Alaska</strong></p><p><strong>Air Canada</strong></p><p><strong>Southwest</strong></p><p>Most airline loyalty programs reward passengers with frequent flyer status complimentary seat selection, bag checks and upgrades to premium seating and first class.</p><p>Hotels</p><p><strong>Policies &amp; Regulations</strong></p><p><strong>Employees&nbsp;are required to&nbsp;book their hotel via Ramp, unless they receive a room rate quote that is a minimum of 10% less expensive than the room rate returned by Ramp.&nbsp; An allowance may be offered to staff presenting proof of pricing to management via email.</strong></p><p><strong>Employees are encouraged to enroll in hotel loyalty&nbsp;programs so that they receive membership benefits including special pricing, free parking, occasional early check-in and late check out privileges, and in-room amenities.&nbsp;Unfortunately, many hotels restrict receipt of “points” and some benefits of loyalty programs when booking through third parties, such as Ramp.</strong></p><p><strong>Employees may book a standard hotel room ≤ $250/night, but should consider the $250 a limit, not an allowance. Management expects our team to book the lowest price, safe, clean, and convenient lodging.</strong></p><p><strong>Employees should book&nbsp;a refundable&nbsp;hotel room only.</strong></p><p><strong>Room, tax, and self-parking are the only allowable expenses on Company credit cards. Employees must put down the Company Visa card at check in to cover these costs.</strong></p><p><strong>Additionally, personal credit cards must be provided to the hotel at check-in for Incidental&nbsp;expenses such as room service, valet parking, in-room entertainment, movies, minibar, and spa purchases are not&nbsp;allowed on company credit cards.</strong></p><p><strong>Paid Valet Parking is not permitted. Self-parking is required.&nbsp;</strong></p><p><strong>Employees should use complimentary internet provided by the hotel only.&nbsp;</strong></p><p><strong>Hotel Trade</strong></p><p>IAE’s Hotel Contact Database,&nbsp;maintained&nbsp;by IAE Marketing, lists IAE hotel partners in the U.S. and Canada. Before booking a hotel, Employees should check with director-level Marketing staff to&nbsp;determine&nbsp;if there are IAE hotel partners&nbsp;located&nbsp;in relevant travel markets.&nbsp;</p><p>Vehicle Rental</p><p><strong>Ramp automates tasks like expense reports, vendor management, and bill payments, and makes it easier for employees to manage their own expenses and&nbsp;maintain&nbsp;clear communication.&nbsp;</strong></p><p>In the interest of&nbsp;maintaining&nbsp;accuracy, which is especially&nbsp;important while traveling,&nbsp;it is mandatory that&nbsp;employees&nbsp;enter&nbsp;their&nbsp;transactions in Ramp accurately - within&nbsp;48 hours&nbsp;- with photo or PDF receipts attached.</p><p>Rules and regulations for travel expenditure are built into Ramp. During the booking process, employees are notified when a flight or hotel&nbsp;selection&nbsp;is out of policy. If there is a need to make travel arrangements that fall outside of company policy, employees should first&nbsp;discuss it&nbsp;with and receive pre-approval from their supervisor. The Ramp system&nbsp;permits&nbsp;employees to book out-of-policy travel, subject to supervisor approval.&nbsp;Accordingly, once pre-approval is received,&nbsp;employees may&nbsp;proceed&nbsp;to make&nbsp;travel arrangements.</p><p><strong>Booking</strong></p><p><strong>Employees must create an Avis.com account and&nbsp;add&nbsp;the IAE&nbsp;Avis AWD Code&nbsp;B903700&nbsp;to&nbsp;their&nbsp;account&nbsp;before&nbsp;soliciting&nbsp;Avis for vehicle rental quotes.&nbsp;Employees are encouraged to enroll in other car rental company rewards programs, but Avis should be among those companies considered.</strong></p><p><strong>Employees are encouraged to shop for the best car rental rates online. Websites like Orbitz, Expedia, or Kayak are useful tools to search for rates, but booking should be done through your selected rental car company’s website.</strong></p><p><strong>Rental car companies often prohibit rentals from anyone under 25 years old. If you are under 25 and need to travel, please discuss with your supervisor.</strong></p><p><strong>Vehicle Type</strong></p><p><strong>Employees are expected to exercise common sense when selecting the approved vehicle type, either standard car, SUV, or van,&nbsp;and base&nbsp;selection on business&nbsp;needs.&nbsp;</strong></p><p><strong>Employees should rent a vehicle that meets your minimal transportation needs. Economy or intermediate class (or below)&nbsp;are permitted.&nbsp;</strong></p><p><strong>SUVs, mini-vans, or 15-passenger vans, may only be rented if transporting equipment, five or more people, or the vehicle is doubling as a&nbsp;Runner vehicle.&nbsp;</strong></p><p><strong>Vehicle Insurance Option</strong></p><p><strong>IAE maintains blanket insurance coverage for rental vehicles. As such, employees should&nbsp;decline all&nbsp;insurance coverage options.&nbsp;</strong></p><p><strong>Prepaid Fuel Option</strong></p><p><strong>Employees should&nbsp;decline&nbsp;the prepaid fuel&nbsp;option and refill the gas tank of rental vehicles before their return.&nbsp;</strong></p><p><strong>Personal Responsibility</strong></p><p><strong>Fines, legal expenses, penalties, moving violations, traffic violations, parking violations, and towing are the responsibility of the employee and may not be paid as a Company expense.</strong></p><p>Ground Transportation&nbsp;</p><p>Rideshare/Taxi</p><p>Rideshare services and taxis may only be used for work-related ground transportation needs when traveling and business-related client events. Employees are encouraged to choose Taxi and Lyft first and discouraged from using Uber. Premium tiers (e.g.&nbsp;Uber Black, Lyft Lux, Lyft Lux Black) are not&nbsp;permitted.</p><p>Rail</p><p><strong>Economy-class/coach-class rail is&nbsp;permitted&nbsp;for business travel.</strong></p><p><strong>First-class rail requires&nbsp;supervisor approval.</strong></p><p>Per Diem</p><p>When traveling for work-related purposes, laundry, meals, groceries, tips, and personal expenses may not be paid using your Company Visa card. When traveling for company business, employees are paid a fixed amount per day to cover these expenditures and do not need to provide copies of receipts or documentation.</p><p>Per diem is an allowance for meals and incidental expenses paid directly to employees who have traveled for work-related purposes. Employees traveling for company business receive a per diem in the amounts as follows:</p><p><strong>Onsite Days: $60/day</strong></p><p><strong>Travel Days: $45/day</strong></p><p>After returning from business travel, employees should&nbsp;submit&nbsp;reimbursement for their per diems via Ramp. Employees must code all travel related Ramp Visa charges spent during the trip before per diems will be distributed.</p><p>Client&nbsp;Meals&nbsp;</p><p><strong>Client meals should be&nbsp;capped&nbsp;at $90 per person&nbsp;and may not include alcohol.</strong></p><p><strong>Gratuities or tips are capped at 20% in the U.S. and 10-15% abroad.&nbsp;</strong></p><p>&nbsp;</p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"2701fa25-eb7f-4ead-ba21-40f9a35a8fb5\",\"sectionIndex\":2,\"sectionFactor\":4,\"controlIndex\":1},\"id\":\"2c2b1596-0197-4020-a475-e6fab7eef288\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"innerHTML\":\"<p class=\\\"noSpacingAbove noSpacingBelow\\\" style=\\\"color:rgb(50, 49, 48);margin-left:0px;\\\" data-text-type=\\\"noSpacing\\\"><span class=\\\"fontSizeXLargePlus\\\"><span class=\\\"fontColorThemeDarker\\\"><strong>Section Contents</strong></span></span></p><p class=\\\"noSpacingAbove noSpacingBelow\\\" style=\\\"color:rgb(50, 49, 48);margin-left:0px;\\\" data-text-type=\\\"noSpacing\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>3</strong></span></span><a href=\\\"#equal-employment-opportunity-%28illinois%29\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>.1 Section 1</strong></span></span></a></p><p class=\\\"noSpacingAbove noSpacingBelow\\\" style=\\\"color:rgb(50, 49, 48);margin-left:0px;\\\" data-text-type=\\\"noSpacing\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>3</strong></span></span><a href=\\\"#equal-employment-opportunity-%28illinois%29\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>.2 Section 2</strong></span></span></a></p><p class=\\\"noSpacingAbove noSpacingBelow\\\" style=\\\"color:rgb(50, 49, 48);margin-left:0px;\\\" aria-hidden=\\\"true\\\" data-text-type=\\\"noSpacing\\\">&nbsp;</p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":3,\"zoneId\":\"f159673e-0fab-48ef-abe0-a6213833cd1b\",\"sectionIndex\":1,\"sectionFactor\":4,\"controlIndex\":1},\"id\":\"db175579-484f-4e33-ba93-e4d5b429f0a2\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"db175579-484f-4e33-ba93-e4d5b429f0a2\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Previous Section\"},\"imageSources\":{},\"links\":{}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Right\",\"minimumLayoutWidth\":1,\"isDynamicWidthEnabled\":false},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":364,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":3,\"zoneId\":\"f159673e-0fab-48ef-abe0-a6213833cd1b\",\"sectionIndex\":2,\"sectionFactor\":4,\"controlIndex\":1},\"id\":\"3eb0e3e0-9ccd-4ca2-88e1-a2ceb2f90dcc\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"3eb0e3e0-9ccd-4ca2-88e1-a2ceb2f90dcc\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Home\"},\"imageSources\":{},\"links\":{}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Center\",\"minimumLayoutWidth\":1,\"isDynamicWidthEnabled\":false},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":364,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":3,\"zoneId\":\"f159673e-0fab-48ef-abe0-a6213833cd1b\",\"sectionIndex\":3,\"sectionFactor\":4,\"controlIndex\":1},\"id\":\"40d0c0b3-1459-4632-a499-28ddb33537ad\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"40d0c0b3-1459-4632-a499-28ddb33537ad\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Next Section\"},\"imageSources\":{},\"links\":{}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Left\",\"minimumLayoutWidth\":1,\"isDynamicWidthEnabled\":false},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":364,\"reservedHeight\":40},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isAIGeneratedDescription\":false,\"isDefaultThumbnail\":true,\"isSpellCheckEnabled\":true,\"globalRichTextStylingVersion\":1,\"rtePageSettings\":{\"contentVersion\":5,\"indentationVersion\":2},\"isEmailReady\":true,\"webPartsPageSettings\":{\"isTitleHeadingLevelsEnabled\":true,\"isLowQualityImagePlaceholderEnabled\":true}}},{\"controlType\":14,\"webPartData\":{\"properties\":{\"zoneBackground\":{\"f374e17d-f830-40c9-859a-b641fed7e09c\":{\"type\":\"image\",\"imageData\":{\"fileName\":\"68234-IAE_Display_28x22_FINAL.jpg\",\"width\":2016,\"height\":720,\"source\":0,\"siteId\":\"44583a03-7429-4de6-9641-ae749e56727f\",\"webId\":\"454df1df-1248-4cb7-a41d-a4e482d69fbb\",\"listId\":\"2ab24772-b566-4fe3-abab-a9edde27e5ea\",\"uniqueId\":\"df6e1993-3198-4e40-9f91-db508a198a82\"},\"overlay\":{\"color\":\"#000000\",\"opacity\":0},\"useLightText\":true}}},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{\"zoneBackground.f374e17d-f830-40c9-859a-b641fed7e09c.imageData.url\":\"/sites/InnovationArtsEntertainment/SiteAssets/SitePages/VolunteerCenter(1)/68234-IAE_Display_28x22_FINAL.jpg\"},\"links\":{}},\"dataVersion\":\"1.0\"}}]"
}

// Message from CEO template with ##MESSAGE## placeholder
const MESSAGEFROMCEO = {
    "__metadata": {
        "type": "SP.Publishing.SitePage"
    },
    "PageRenderingState": 0,
    "CanvasContent1": "[{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"c12d16a8-30de-46cb-8d44-f22d081aace5\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":1},\"zoneHeight\":550,\"id\":\"3dbcc936-0d52-4658-91f8-dc3cf63925c1\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":0,\"y\":0,\"w\":70,\"h\":7,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<h2 class=\\\"headingSpacingAbove headingSpacingBelow lineHeight1_0\\\"><span class=\\\"fontSizeMega rte-fontscale-font-max\\\">1.1 A&nbsp;Message&nbsp;from the CEO</span></h2>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"c12d16a8-30de-46cb-8d44-f22d081aace5\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":2},\"zoneHeight\":550,\"id\":\"1857d513-039d-4197-8e49-4bd8e6f378c4\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":0,\"y\":6,\"w\":70,\"h\":23,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><span class=\\\"fontSizeXLarge\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\">Welcome!&nbsp;</span>&nbsp;</span></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><span class=\\\"fontSizeXLarge\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\">You have just joined a dynamic and rapidly growing company. We hope that your employment with Innovation Arts &amp; Entertainment&nbsp;will be&nbsp;both&nbsp;challenging&nbsp;and rewarding.&nbsp;We are&nbsp;proud of the&nbsp;professional&nbsp;services we provide – and&nbsp;we are&nbsp;equally&nbsp;proud of our employees,&nbsp;who enable and enhance our ability to perform&nbsp;our work at&nbsp;a high level.</span> &nbsp;</span></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><span class=\\\"fontSizeXLarge\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\">Our&nbsp;employee&nbsp;handbook is intended to provide employees with important policy&nbsp;and procedure&nbsp;guidelines to learn and follow, as well as a reference source for many aspects of life&nbsp;at&nbsp;Innovation Arts &amp; Entertainment.</span> &nbsp;</span></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><span class=\\\"fontSizeXLarge\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\">Our&nbsp;people are energetic, and our&nbsp;office&nbsp;life&nbsp;is&nbsp;fast paced.&nbsp;Be prepared to learn, grow, and perform important work,&nbsp;daily – and your days here will fly by.</span>&nbsp;</span></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><span class=\\\"fontSizeXLarge\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\">We&nbsp;wish you&nbsp;remarkable&nbsp;success!</span>&nbsp;</span></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" aria-hidden=\\\"true\\\" data-text-type=\\\"withSpacing\\\">&nbsp;</p><p class=\\\"noSpacingAbove noSpacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"noSpacing\\\"><span class=\\\"fontSizeXLarge\\\"><i><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>Adam Epstein</strong></span></i></span></p><p class=\\\"noSpacingAbove noSpacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"noSpacing\\\"><span class=\\\"fontSizeXLarge\\\"><i><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\">Founder &amp; CEO</span></i></span></p><p class=\\\"noSpacingAbove noSpacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"noSpacing\\\" aria-hidden=\\\"true\\\">&nbsp;</p><p class=\\\"noSpacingAbove noSpacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"noSpacing\\\" aria-hidden=\\\"true\\\">&nbsp;</p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"c12d16a8-30de-46cb-8d44-f22d081aace5\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":3},\"zoneHeight\":550,\"id\":\"bd7c16e0-b868-43ce-9742-bd272082eb3c\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":25,\"y\":29,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"bd7c16e0-b868-43ce-9742-bd272082eb3c\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Home\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/HANDBOOK.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Center\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"c12d16a8-30de-46cb-8d44-f22d081aace5\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":4},\"zoneHeight\":550,\"id\":\"2a4fe9c1-c46f-4664-ac1b-b0179329f01b\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":38,\"y\":29,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"2a4fe9c1-c46f-4664-ac1b-b0179329f01b\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Next Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Employment%20Policies%20and%20Practices.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Center\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"0ba22869-c07a-4c44-bccc-521aae856fe8\",\"sectionIndex\":1,\"sectionFactor\":8,\"controlIndex\":1},\"id\":\"48047102-2fb5-4efc-ac96-034716e6f250\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"innerHTML\":\"<p class=\\\"noSpacingAbove spacingBelow\\\" data-text-type=\\\"withSpacing\\\"><span style=\\\"color:#0451a5;\\\">##MESSAGE##</span></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"0ba22869-c07a-4c44-bccc-521aae856fe8\",\"sectionIndex\":2,\"sectionFactor\":4,\"controlIndex\":1},\"id\":\"4696d89e-ec85-4c9b-8ebf-e04f59ab5343\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"innerHTML\":\"<p class=\\\"noSpacingAbove noSpacingBelow\\\" data-text-type=\\\"noSpacing\\\"><span class=\\\"fontSizeXLargePlus\\\"><span class=\\\"fontColorThemeDarker\\\"><strong>Section Contents</strong></span></span></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"0ba22869-c07a-4c44-bccc-521aae856fe8\",\"sectionIndex\":2,\"sectionFactor\":4,\"controlIndex\":2},\"id\":\"da32caa7-5244-4649-8293-35d10e746a80\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"innerHTML\":\"<p class=\\\"lineHeight1_2\\\" style=\\\"margin-left:0px;\\\"><a href=\\\"#a-message-from-the-ceo\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>1.1 </strong>CEO</span> Message</span></a></p><p class=\\\"lineHeight1_2\\\" style=\\\"margin-left:0px;\\\"><a href=\\\"#the-company\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>1.2</strong> The Company</span>&nbsp;</span></a></p><p class=\\\"lineHeight1_2\\\" style=\\\"margin-left:0px;\\\"><a href=\\\"#change-in-policy\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>1.3 </strong>Change in Policy</span>&nbsp;</span></a></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":3,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":1},\"zoneHeight\":1600,\"id\":\"9a58ab09-29e4-4d76-b0f5-8b8a114cb40e\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":36,\"y\":87,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"9a58ab09-29e4-4d76-b0f5-8b8a114cb40e\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Next Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Employment%20Policies%20and%20Practices.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Right\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":3,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":2},\"zoneHeight\":1600,\"id\":\"602be227-1a60-4fd8-a4bd-3f69c27a0fa8\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":3,\"y\":87,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"602be227-1a60-4fd8-a4bd-3f69c27a0fa8\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Home\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/HANDBOOK.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Center\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":3,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":3},\"zoneHeight\":1600,\"id\":\"8048e3e4-74fd-4e6b-9d4f-cfb5649079d8\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":5,\"y\":85,\"w\":40,\"h\":0,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"2161a1c6-db61-4731-b97c-3cdb303f7cbb\",\"instanceId\":\"8048e3e4-74fd-4e6b-9d4f-cfb5649079d8\",\"title\":\"Divider\",\"description\":\"Add a line to divide areas on your page.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{},\"links\":{}},\"dataVersion\":\"1.2\",\"properties\":{\"minimumLayoutWidth\":1,\"length\":100,\"weight\":1},\"containsDynamicDataSource\":false},\"webPartId\":\"2161a1c6-db61-4731-b97c-3cdb303f7cbb\",\"reservedWidth\":688,\"reservedHeight\":1},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isAIGeneratedDescription\":false,\"isDefaultThumbnail\":true,\"isSpellCheckEnabled\":true,\"globalRichTextStylingVersion\":1,\"rtePageSettings\":{\"contentVersion\":5,\"indentationVersion\":2},\"isEmailReady\":true,\"webPartsPageSettings\":{\"isTitleHeadingLevelsEnabled\":true,\"isLowQualityImagePlaceholderEnabled\":true}}},{\"controlType\":14,\"webPartData\":{\"properties\":{\"zoneBackground\":{\"c12d16a8-30de-46cb-8d44-f22d081aace5\":{\"type\":\"image\",\"imageData\":{\"fileName\":\"68234-IAE_Display_28x22_FINAL.jpg\",\"width\":2016,\"height\":720,\"source\":0,\"siteId\":\"44583a03-7429-4de6-9641-ae749e56727f\",\"webId\":\"454df1df-1248-4cb7-a41d-a4e482d69fbb\",\"listId\":\"2ab24772-b566-4fe3-abab-a9edde27e5ea\",\"uniqueId\":\"df6e1993-3198-4e40-9f91-db508a198a82\"},\"overlay\":{\"color\":\"#000000\",\"opacity\":0},\"useLightText\":true},\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\":{\"type\":\"image\",\"imageData\":{\"fileName\":\"template11Image0003.png\",\"width\":144,\"height\":144,\"source\":1},\"overlay\":{\"color\":\"#FFFFFF\",\"opacity\":100},\"useLightText\":false}},\"zoneThemeIndex\":{\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\":{\"lightIndex\":0,\"darkIndex\":0}}},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{\"zoneBackground.c12d16a8-30de-46cb-8d44-f22d081aace5.imageData.url\":\"/sites/InnovationArtsEntertainment/SiteAssets/SitePages/VolunteerCenter(1)/68234-IAE_Display_28x22_FINAL.jpg\",\"zoneBackground.54f8a2d2-b737-4a6e-96a2-54ae0a01c7da.imageData.url\":\"https://cdn.hubblecontent.osi.office.net/m365content/publish/4b99b582-026b-4490-ab73-b0091ab721cc/image.png\"},\"links\":{}},\"dataVersion\":\"1.0\"}}]"
}

// Compensation and Benefits template with ##MESSAGE## placeholder
const COMPENSATIONBENEFITSTEMPLATE = {
    "__metadata": {
        "type": "SP.Publishing.SitePage"
    },
    "PageRenderingState": 0,
    "CanvasContent1": "[{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"f7362d98-6761-44d4-bed3-ef0a71ae0136\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":1},\"zoneHeight\":241,\"id\":\"3dbcc936-0d52-4658-91f8-dc3cf63925c1\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":0,\"y\":2,\"w\":70,\"h\":7,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;text-align:center;\\\" data-text-type=\\\"withSpacing\\\"><span class=\\\"fontSizeMega rte-fontscale-font-max\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>Compensation and Benefits</strong></span><strong>&nbsp;</strong></span></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"f7362d98-6761-44d4-bed3-ef0a71ae0136\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":2},\"zoneHeight\":241,\"id\":\"caef005c-415f-4175-a7d3-42e39f461ece\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":18,\"y\":10,\"w\":11,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"caef005c-415f-4175-a7d3-42e39f461ece\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Previous Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Company%20Policies%20and%20Practices.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Left\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":189,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"f7362d98-6761-44d4-bed3-ef0a71ae0136\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":3},\"zoneHeight\":241,\"id\":\"e9771282-8641-48eb-8d2c-ed79803de3cf\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":30,\"y\":10,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"e9771282-8641-48eb-8d2c-ed79803de3cf\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Home\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/HANDBOOK.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Center\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"f7362d98-6761-44d4-bed3-ef0a71ae0136\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":4},\"zoneHeight\":241,\"id\":\"1a60be60-f7f0-4661-a273-585d1abc91c5\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":40,\"y\":10,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"1a60be60-f7f0-4661-a273-585d1abc91c5\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Next Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Work%20Performance.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Right\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":1},\"zoneHeight\":5435,\"id\":\"48047102-2fb5-4efc-ac96-034716e6f250\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":4,\"y\":1,\"w\":45,\"h\":3}},\"innerHTML\":\"<p><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\">##MESSAGE##</span></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":2},\"zoneHeight\":5435,\"id\":\"4696d89e-ec85-4c9b-8ebf-e04f59ab5343\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":52,\"y\":1,\"w\":14,\"h\":3,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<p class=\\\"noSpacingAbove noSpacingBelow\\\" data-text-type=\\\"noSpacing\\\"><span class=\\\"fontSizeXLargePlus\\\"><span class=\\\"fontColorThemeDarker\\\"><strong>Section Contents</strong></span></span></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":3},\"zoneHeight\":5435,\"id\":\"da32caa7-5244-4649-8293-35d10e746a80\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":52,\"y\":3,\"w\":17,\"h\":23,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<p class=\\\"noSpacingAbove spacingBelow lineHeight1_0\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"#payroll-schedule\\\"><span class=\\\"fontSizeXLarge\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>4.1 </strong>Payroll Schedule</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_0\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"#payroll-deductions\\\"><span class=\\\"fontSizeXLarge\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>4.2</strong> Payroll Deductions</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_0\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"/sites/InnovationArtsEntertainment/SitePages/Compensation%20and%20Benefits.aspx#4.3.-direct-deposit\\\"><span class=\\\"fontSizeXLarge\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>4.3 </strong>Direct Deposit</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_0\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"/sites/InnovationArtsEntertainment/SitePages/Compensation%20and%20Benefits.aspx#4.4.-healthcare-plans-%28medical%2c-dental%2c-vision%29\\\"><span class=\\\"fontSizeXLarge\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>4.4 </strong>Healthcare Plans (Medical, Dental, Vision)</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_0\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"/sites/InnovationArtsEntertainment/SitePages/Compensation%20and%20Benefits.aspx#4.5.-ira-retirement-plan\\\"><span class=\\\"fontSizeXLarge\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>4.5</strong> IRA Retirement Plan</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_0\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"/sites/InnovationArtsEntertainment/SitePages/Compensation%20and%20Benefits.aspx#4.6.-paid-time-off-%28vacation%2c-sick%2c-and-bereavement-leave%29\\\"><span class=\\\"fontSizeXLarge\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>4.6 </strong>Paid Time Off (vacation,&nbsp;sick, bereavement)&nbsp;</span></span></a><span class=\\\"fontSizeXLarge\\\">&nbsp;</span></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_0\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"/sites/InnovationArtsEntertainment/SitePages/Compensation%20and%20Benefits.aspx#4.7.-holidays\\\"><span class=\\\"fontSizeXLarge\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>4.7</strong> Holidays</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_0\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"/sites/InnovationArtsEntertainment/SitePages/Compensation%20and%20Benefits.aspx#4.8.-pregnancy-accommodation-policy\\\"><span class=\\\"fontSizeXLarge\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>4.8 </strong>Pregnancy Accommodation Policy</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_0\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"/sites/InnovationArtsEntertainment/SitePages/Compensation%20and%20Benefits.aspx#4.9.-unpaid-leave-of-absence\\\"><span class=\\\"fontSizeXLarge\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>4.9 </strong>Unpaid Leave of Absence</span></span></a><span class=\\\"fontSizeXLarge\\\">&nbsp;</span></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":4},\"zoneHeight\":5435,\"id\":\"9a58ab09-29e4-4d76-b0f5-8b8a114cb40e\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":36,\"y\":312,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"9a58ab09-29e4-4d76-b0f5-8b8a114cb40e\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Next Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Work%20Performance.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Right\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":5},\"zoneHeight\":5435,\"id\":\"602be227-1a60-4fd8-a4bd-3f69c27a0fa8\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":21,\"y\":312,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"602be227-1a60-4fd8-a4bd-3f69c27a0fa8\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Home\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/HANDBOOK.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Center\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":6},\"zoneHeight\":5435,\"id\":\"22a99bb7-d8fb-4e17-96cf-2c0d90d5988f\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":5,\"y\":312,\"w\":12,\"h\":3}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"22a99bb7-d8fb-4e17-96cf-2c0d90d5988f\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Previous Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Company%20Policies%20and%20Practices.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Left\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":206,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":7},\"zoneHeight\":5435,\"id\":\"8048e3e4-74fd-4e6b-9d4f-cfb5649079d8\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":5,\"y\":310,\"w\":40,\"h\":0,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"2161a1c6-db61-4731-b97c-3cdb303f7cbb\",\"instanceId\":\"8048e3e4-74fd-4e6b-9d4f-cfb5649079d8\",\"title\":\"Divider\",\"description\":\"Add a line to divide areas on your page.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{},\"links\":{}},\"dataVersion\":\"1.2\",\"properties\":{\"minimumLayoutWidth\":1,\"length\":100,\"weight\":1},\"containsDynamicDataSource\":false},\"webPartId\":\"2161a1c6-db61-4731-b97c-3cdb303f7cbb\",\"reservedWidth\":688,\"reservedHeight\":1},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isAIGeneratedDescription\":false,\"isDefaultThumbnail\":true,\"isSpellCheckEnabled\":true,\"globalRichTextStylingVersion\":1,\"rtePageSettings\":{\"contentVersion\":5,\"indentationVersion\":2},\"isEmailReady\":true,\"webPartsPageSettings\":{\"isTitleHeadingLevelsEnabled\":true,\"isLowQualityImagePlaceholderEnabled\":true}}},{\"controlType\":14,\"webPartData\":{\"properties\":{\"zoneBackground\":{\"f7362d98-6761-44d4-bed3-ef0a71ae0136\":{\"type\":\"image\",\"imageData\":{\"fileName\":\"68234-IAE_Display_28x22_FINAL.jpg\",\"width\":2016,\"height\":720,\"source\":0,\"siteId\":\"44583a03-7429-4de6-9641-ae749e56727f\",\"webId\":\"454df1df-1248-4cb7-a41d-a4e482d69fbb\",\"listId\":\"2ab24772-b566-4fe3-abab-a9edde27e5ea\",\"uniqueId\":\"df6e1993-3198-4e40-9f91-db508a198a82\"},\"overlay\":{\"color\":\"#000000\",\"opacity\":0},\"useLightText\":true},\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\":{\"type\":\"image\",\"imageData\":{\"fileName\":\"template11Image0003.png\",\"width\":144,\"height\":144,\"source\":1},\"overlay\":{\"color\":\"#FFFFFF\",\"opacity\":100},\"useLightText\":false}},\"zoneThemeIndex\":{\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\":{\"lightIndex\":0,\"darkIndex\":0}}},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{\"zoneBackground.f7362d98-6761-44d4-bed3-ef0a71ae0136.imageData.url\":\"/sites/InnovationArtsEntertainment/SiteAssets/SitePages/VolunteerCenter(1)/68234-IAE_Display_28x22_FINAL.jpg\",\"zoneBackground.54f8a2d2-b737-4a6e-96a2-54ae0a01c7da.imageData.url\":\"https://cdn.hubblecontent.osi.office.net/m365content/publish/4b99b582-026b-4490-ab73-b0091ab721cc/image.png\"},\"links\":{}},\"dataVersion\":\"1.0\"}}]"
}

// Work Performance template with ##MESSAGE## placeholder
const WORKPERFORMANCETEMPLATE = {
    "__metadata": {
        "type": "SP.Publishing.SitePage"
    },
    "PageRenderingState": 0,
    "CanvasContent1": "[{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"2e494aed-0777-42c6-9841-4605fb66bddf\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":1},\"zoneHeight\":310,\"id\":\"3dbcc936-0d52-4658-91f8-dc3cf63925c1\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":0,\"y\":3,\"w\":70,\"h\":7,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;text-align:center;\\\" data-text-type=\\\"withSpacing\\\"><span class=\\\"fontSizeMega rte-fontscale-font-max\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>Work Performance</strong></span></span></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"2e494aed-0777-42c6-9841-4605fb66bddf\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":2},\"zoneHeight\":310,\"id\":\"e6c304d1-4ff4-4562-88ea-4b4134354677\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":19,\"y\":12,\"w\":11,\"h\":3}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"e6c304d1-4ff4-4562-88ea-4b4134354677\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Previous Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Compensation%20and%20Benefits.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Left\",\"minimumLayoutWidth\":9,\"isDynamicWidthEnabled\":false},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":189,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"2e494aed-0777-42c6-9841-4605fb66bddf\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":3},\"zoneHeight\":310,\"id\":\"1cc7f6a0-f271-4e20-bc51-45703c059bbe\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":31,\"y\":12,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"1cc7f6a0-f271-4e20-bc51-45703c059bbe\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Home\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/HANDBOOK.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Center\",\"minimumLayoutWidth\":9,\"isDynamicWidthEnabled\":false},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"2e494aed-0777-42c6-9841-4605fb66bddf\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":4},\"zoneHeight\":310,\"id\":\"7da83a8f-9562-4eb8-99d5-e1055ad267b1\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":42,\"y\":12,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"7da83a8f-9562-4eb8-99d5-e1055ad267b1\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Next Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/:u:/s/InnovationArtsEntertainment/IQC0klnHP1udSp9g0bW61VjhATFoXOMrmHzfRcC-N9y-SDM?e=lszZQc\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Right\",\"minimumLayoutWidth\":9,\"isDynamicWidthEnabled\":false},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":1},\"zoneHeight\":3818,\"id\":\"48047102-2fb5-4efc-ac96-034716e6f250\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":4,\"y\":1,\"w\":44,\"h\":3,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<p>##MESSAGE##</p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":2},\"zoneHeight\":3818,\"id\":\"4696d89e-ec85-4c9b-8ebf-e04f59ab5343\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":52,\"y\":1,\"w\":14,\"h\":3,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<p class=\\\"noSpacingAbove noSpacingBelow\\\" data-text-type=\\\"noSpacing\\\"><span class=\\\"fontSizeXLargePlus\\\"><span class=\\\"fontColorThemeDarker\\\"><strong>Section Contents</strong></span></span></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":3},\"zoneHeight\":3818,\"id\":\"da32caa7-5244-4649-8293-35d10e746a80\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":52,\"y\":3,\"w\":17,\"h\":10,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<p class=\\\"noSpacingAbove spacingBelow lineHeight1_0\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"#performance-reviews-%286-months%2c-then-annual%29\\\"><span class=\\\"fontSizeXLarge\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>5.1</strong> Performance Reviews</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_0\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"#discipline-policy\\\"><span class=\\\"fontSizeXLarge\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>5.2</strong> Discipline Policy</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_0\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"#termination-of-employment\\\"><span class=\\\"fontSizeXLarge\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>5.3 </strong>Termination of Employment</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_0\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"#cobra-health-insurance-continuing-coverage\\\"><span class=\\\"fontSizeXLarge\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>5.4 </strong>COBRA Health Insurance Continuing Coverage</span>&nbsp;</span></a></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":4},\"zoneHeight\":3818,\"id\":\"9a58ab09-29e4-4d76-b0f5-8b8a114cb40e\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":37,\"y\":192,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"9a58ab09-29e4-4d76-b0f5-8b8a114cb40e\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Next Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/:u:/s/InnovationArtsEntertainment/IQC0klnHP1udSp9g0bW61VjhATFoXOMrmHzfRcC-N9y-SDM?e=lszZQc\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Right\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":5},\"zoneHeight\":3818,\"id\":\"602be227-1a60-4fd8-a4bd-3f69c27a0fa8\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":22,\"y\":192,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"602be227-1a60-4fd8-a4bd-3f69c27a0fa8\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Home\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/HANDBOOK.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Center\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":6},\"zoneHeight\":3818,\"id\":\"22a99bb7-d8fb-4e17-96cf-2c0d90d5988f\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":6,\"y\":192,\"w\":12,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"22a99bb7-d8fb-4e17-96cf-2c0d90d5988f\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Previous Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Compensation%20and%20Benefits.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Left\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":206,\"reservedHeight\":40},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isAIGeneratedDescription\":false,\"isDefaultThumbnail\":true,\"isSpellCheckEnabled\":true,\"globalRichTextStylingVersion\":1,\"rtePageSettings\":{\"contentVersion\":5,\"indentationVersion\":2},\"isEmailReady\":true,\"webPartsPageSettings\":{\"isTitleHeadingLevelsEnabled\":true,\"isLowQualityImagePlaceholderEnabled\":true}}},{\"controlType\":14,\"webPartData\":{\"properties\":{\"zoneBackground\":{\"2e494aed-0777-42c6-9841-4605fb66bddf\":{\"type\":\"image\",\"imageData\":{\"fileName\":\"68234-IAE_Display_28x22_FINAL.jpg\",\"width\":2016,\"height\":720,\"source\":0,\"siteId\":\"44583a03-7429-4de6-9641-ae749e56727f\",\"webId\":\"454df1df-1248-4cb7-a41d-a4e482d69fbb\",\"listId\":\"2ab24772-b566-4fe3-abab-a9edde27e5ea\",\"uniqueId\":\"df6e1993-3198-4e40-9f91-db508a198a82\"},\"overlay\":{\"color\":\"#000000\",\"opacity\":0},\"useLightText\":true},\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\":{\"type\":\"image\",\"imageData\":{\"fileName\":\"template11Image0003.png\",\"width\":144,\"height\":144,\"source\":1},\"overlay\":{\"color\":\"#FFFFFF\",\"opacity\":100},\"useLightText\":false}},\"zoneThemeIndex\":{\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\":{\"lightIndex\":0,\"darkIndex\":0}}},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{\"zoneBackground.2e494aed-0777-42c6-9841-4605fb66bddf.imageData.url\":\"/sites/InnovationArtsEntertainment/SiteAssets/SitePages/VolunteerCenter(1)/68234-IAE_Display_28x22_FINAL.jpg\",\"zoneBackground.54f8a2d2-b737-4a6e-96a2-54ae0a01c7da.imageData.url\":\"https://cdn.hubblecontent.osi.office.net/m365content/publish/4b99b582-026b-4490-ab73-b0091ab721cc/image.png\"},\"links\":{}},\"dataVersion\":\"1.0\"}}]"
}

// Procedures and Guidelines template with ##MESSAGE## placeholder
const PROCEDURESGUIDELINESTEMPLATE = {
    "__metadata": {
        "type": "SP.Publishing.SitePage"
    },
    "PageRenderingState": 0,
    "CanvasContent1": "[{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"2e494aed-0777-42c6-9841-4605fb66bddf\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":1},\"zoneHeight\":310,\"id\":\"3dbcc936-0d52-4658-91f8-dc3cf63925c1\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":0,\"y\":3,\"w\":70,\"h\":7,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;text-align:center;\\\" data-text-type=\\\"withSpacing\\\"><span class=\\\"fontSizeMega rte-fontscale-font-max\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>Procedures and Guidelines</strong></span></span></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"2e494aed-0777-42c6-9841-4605fb66bddf\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":2},\"zoneHeight\":310,\"id\":\"e6c304d1-4ff4-4562-88ea-4b4134354677\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":19,\"y\":12,\"w\":11,\"h\":3}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"e6c304d1-4ff4-4562-88ea-4b4134354677\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Previous Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Work%20Performance.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Left\",\"minimumLayoutWidth\":9,\"isDynamicWidthEnabled\":false},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":189,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"2e494aed-0777-42c6-9841-4605fb66bddf\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":3},\"zoneHeight\":310,\"id\":\"1cc7f6a0-f271-4e20-bc51-45703c059bbe\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":31,\"y\":12,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"1cc7f6a0-f271-4e20-bc51-45703c059bbe\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Home\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/HANDBOOK.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Center\",\"minimumLayoutWidth\":9,\"isDynamicWidthEnabled\":false},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"2e494aed-0777-42c6-9841-4605fb66bddf\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":4},\"zoneHeight\":310,\"id\":\"7da83a8f-9562-4eb8-99d5-e1055ad267b1\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":42,\"y\":12,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"7da83a8f-9562-4eb8-99d5-e1055ad267b1\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Next Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/:u:/s/InnovationArtsEntertainment/IQC0klnHP1udSp9g0bW61VjhATFoXOMrmHzfRcC-N9y-SDM?e=lszZQc\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Right\",\"minimumLayoutWidth\":9,\"isDynamicWidthEnabled\":false},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":1},\"zoneHeight\":3818,\"id\":\"48047102-2fb5-4efc-ac96-034716e6f250\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":4,\"y\":1,\"w\":44,\"h\":3,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<p>##MESSAGE##</p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":2},\"zoneHeight\":3818,\"id\":\"4696d89e-ec85-4c9b-8ebf-e04f59ab5343\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":52,\"y\":1,\"w\":14,\"h\":3,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<p class=\\\"noSpacingAbove noSpacingBelow\\\" data-text-type=\\\"noSpacing\\\"><span class=\\\"fontSizeXLargePlus\\\"><span class=\\\"fontColorThemeDarker\\\"><strong>Section Contents</strong></span></span></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":3},\"zoneHeight\":3818,\"id\":\"da32caa7-5244-4649-8293-35d10e746a80\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":52,\"y\":3,\"w\":17,\"h\":10,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<p class=\\\"noSpacingAbove spacingBelow lineHeight1_0\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"#procedures\\\"><span class=\\\"fontSizeXLarge\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>6.1</strong> Procedures</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_0\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"#guidelines\\\"><span class=\\\"fontSizeXLarge\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>6.2</strong> Guidelines</span>&nbsp;</span></a></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":4},\"zoneHeight\":3818,\"id\":\"9a58ab09-29e4-4d76-b0f5-8b8a114cb40e\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":37,\"y\":192,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"9a58ab09-29e4-4d76-b0f5-8b8a114cb40e\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Next Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/:u:/s/InnovationArtsEntertainment/IQC0klnHP1udSp9g0bW61VjhATFoXOMrmHzfRcC-N9y-SDM?e=lszZQc\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Right\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":5},\"zoneHeight\":3818,\"id\":\"602be227-1a60-4fd8-a4bd-3f69c27a0fa8\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":22,\"y\":192,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"602be227-1a60-4fd8-a4bd-3f69c27a0fa8\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Home\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/HANDBOOK.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Center\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":6},\"zoneHeight\":3818,\"id\":\"22a99bb7-d8fb-4e17-96cf-2c0d90d5988f\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":6,\"y\":192,\"w\":12,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"22a99bb7-d8fb-4e17-96cf-2c0d90d5988f\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Previous Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Work%20Performance.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Left\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":206,\"reservedHeight\":40},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isAIGeneratedDescription\":false,\"isDefaultThumbnail\":true,\"isSpellCheckEnabled\":true,\"globalRichTextStylingVersion\":1,\"rtePageSettings\":{\"contentVersion\":5,\"indentationVersion\":2},\"isEmailReady\":true,\"webPartsPageSettings\":{\"isTitleHeadingLevelsEnabled\":true,\"isLowQualityImagePlaceholderEnabled\":true}}},{\"controlType\":14,\"webPartData\":{\"properties\":{\"zoneBackground\":{\"2e494aed-0777-42c6-9841-4605fb66bddf\":{\"type\":\"image\",\"imageData\":{\"fileName\":\"68234-IAE_Display_28x22_FINAL.jpg\",\"width\":2016,\"height\":720,\"source\":0,\"siteId\":\"44583a03-7429-4de6-9641-ae749e56727f\",\"webId\":\"454df1df-1248-4cb7-a41d-a4e482d69fbb\",\"listId\":\"2ab24772-b566-4fe3-abab-a9edde27e5ea\",\"uniqueId\":\"df6e1993-3198-4e40-9f91-db508a198a82\"},\"overlay\":{\"color\":\"#000000\",\"opacity\":0},\"useLightText\":true},\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\":{\"type\":\"image\",\"imageData\":{\"fileName\":\"template11Image0003.png\",\"width\":144,\"height\":144,\"source\":1},\"overlay\":{\"color\":\"#FFFFFF\",\"opacity\":100},\"useLightText\":false}},\"zoneThemeIndex\":{\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\":{\"lightIndex\":0,\"darkIndex\":0}}},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{\"zoneBackground.2e494aed-0777-42c6-9841-4605fb66bddf.imageData.url\":\"/sites/InnovationArtsEntertainment/SiteAssets/SitePages/VolunteerCenter(1)/68234-IAE_Display_28x22_FINAL.jpg\",\"zoneBackground.54f8a2d2-b737-4a6e-96a2-54ae0a01c7da.imageData.url\":\"https://cdn.hubblecontent.osi.office.net/m365content/publish/4b99b582-026b-4490-ab73-b0091ab721cc/image.png\"},\"links\":{}},\"dataVersion\":\"1.0\"}}]"
}

// Employee Acknowledgment template with ##MESSAGE## placeholder
const EMPLOYEEACKNOWLEDGMENTTEMPLATE = {
    "__metadata": {
        "type": "SP.Publishing.SitePage"
    },
    "PageRenderingState": 0,
    "CanvasContent1": "[{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"2e494aed-0777-42c6-9841-4605fb66bddf\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":1},\"zoneHeight\":310,\"id\":\"3dbcc936-0d52-4658-91f8-dc3cf63925c1\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":0,\"y\":3,\"w\":70,\"h\":7,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;text-align:center;\\\" data-text-type=\\\"withSpacing\\\"><span class=\\\"fontSizeMega rte-fontscale-font-max\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>Employee Acknowledgment</strong></span></span></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"2e494aed-0777-42c6-9841-4605fb66bddf\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":2},\"zoneHeight\":310,\"id\":\"e6c304d1-4ff4-4562-88ea-4b4134354677\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":19,\"y\":12,\"w\":11,\"h\":3}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"e6c304d1-4ff4-4562-88ea-4b4134354677\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Previous Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Work%20Performance.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Left\",\"minimumLayoutWidth\":9,\"isDynamicWidthEnabled\":false},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":189,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"2e494aed-0777-42c6-9841-4605fb66bddf\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":3},\"zoneHeight\":310,\"id\":\"1cc7f6a0-f271-4e20-bc51-45703c059bbe\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":31,\"y\":12,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"1cc7f6a0-f271-4e20-bc51-45703c059bbe\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Home\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/HANDBOOK.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Center\",\"minimumLayoutWidth\":9,\"isDynamicWidthEnabled\":false},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":1},\"zoneHeight\":3818,\"id\":\"48047102-2fb5-4efc-ac96-034716e6f250\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":4,\"y\":1,\"w\":44,\"h\":3,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<p>##MESSAGE##</p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":2},\"zoneHeight\":3818,\"id\":\"602be227-1a60-4fd8-a4bd-3f69c27a0fa8\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":22,\"y\":192,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"602be227-1a60-4fd8-a4bd-3f69c27a0fa8\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Home\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/HANDBOOK.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Center\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":3},\"zoneHeight\":3818,\"id\":\"22a99bb7-d8fb-4e17-96cf-2c0d90d5988f\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":6,\"y\":192,\"w\":12,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"22a99bb7-d8fb-4e17-96cf-2c0d90d5988f\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Previous Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Work%20Performance.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Left\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":206,\"reservedHeight\":40},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isAIGeneratedDescription\":false,\"isDefaultThumbnail\":true,\"isSpellCheckEnabled\":true,\"globalRichTextStylingVersion\":1,\"rtePageSettings\":{\"contentVersion\":5,\"indentationVersion\":2},\"isEmailReady\":true,\"webPartsPageSettings\":{\"isTitleHeadingLevelsEnabled\":true,\"isLowQualityImagePlaceholderEnabled\":true}}},{\"controlType\":14,\"webPartData\":{\"properties\":{\"zoneBackground\":{\"2e494aed-0777-42c6-9841-4605fb66bddf\":{\"type\":\"image\",\"imageData\":{\"fileName\":\"68234-IAE_Display_28x22_FINAL.jpg\",\"width\":2016,\"height\":720,\"source\":0,\"siteId\":\"44583a03-7429-4de6-9641-ae749e56727f\",\"webId\":\"454df1df-1248-4cb7-a41d-a4e482d69fbb\",\"listId\":\"2ab24772-b566-4fe3-abab-a9edde27e5ea\",\"uniqueId\":\"df6e1993-3198-4e40-9f91-db508a198a82\"},\"overlay\":{\"color\":\"#000000\",\"opacity\":0},\"useLightText\":true},\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\":{\"type\":\"image\",\"imageData\":{\"fileName\":\"template11Image0003.png\",\"width\":144,\"height\":144,\"source\":1},\"overlay\":{\"color\":\"#FFFFFF\",\"opacity\":100},\"useLightText\":false}},\"zoneThemeIndex\":{\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\":{\"lightIndex\":0,\"darkIndex\":0}}},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{\"zoneBackground.2e494aed-0777-42c6-9841-4605fb66bddf.imageData.url\":\"/sites/InnovationArtsEntertainment/SiteAssets/SitePages/VolunteerCenter(1)/68234-IAE_Display_28x22_FINAL.jpg\",\"zoneBackground.54f8a2d2-b737-4a6e-96a2-54ae0a01c7da.imageData.url\":\"https://cdn.hubblecontent.osi.office.net/m365content/publish/4b99b582-026b-4490-ab73-b0091ab721cc/image.png\"},\"links\":{}},\"dataVersion\":\"1.0\"}}]"
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
                    let canvasContent = SHAREPOINT_TEMPLATE.CanvasContent1
                        .replace('##CEOMESSAGE##', escapedCeoMessage)
                        .replace('##COMPANYINFO##', escapedCompanyInfo);

                    const response = {
                        ...SHAREPOINT_TEMPLATE,
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
                let canvasContent = SHAREPOINT_TEMPLATE.CanvasContent1
                    .replace('##CEOMESSAGE##', escapedCeoMessage)
                    .replace('##COMPANYINFO##', escapedCompanyInfo);

                const response = {
                    ...SHAREPOINT_TEMPLATE,
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

// Introduction endpoint - returns processed content without template
app.post('/api/introduction', bodyParser.raw({ type: '*/*', limit: '50mb' }), async (req, res) => {
    try {
        let isJSON = false;
        let jsonBody = null;
        let titleFilter = req.query.title ? req.query.title.trim() : undefined; // Get title from query string and trim

        // Try to detect if it's JSON by checking the first character
        if (req.body.length > 0 && (req.body[0] === 0x7B || req.body[0] === 0x5B)) { // { or [
            try {
                jsonBody = JSON.parse(req.body.toString('utf8'));
                isJSON = true;
                // Also check for title in JSON body
                if (jsonBody.title || jsonBody['$title']) {
                    titleFilter = (jsonBody.title || jsonBody['$title']).trim();
                }
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

                    // Extract all H1 sections
                    const h1Sections = extractAllH1Sections(htmlContent);

                    // If title filter is provided, return only that section
                    if (titleFilter) {
                        const matchedSection = h1Sections.find(section =>
                            section.title.toLowerCase().includes(titleFilter.toLowerCase())
                        );

                        if (matchedSection) {
                            return res.status(200).json({
                                status: 'success',
                                message: matchedSection.content
                            });
                        } else {
                            return res.status(404).json({
                                status: 'error',
                                message: `Section with title "${titleFilter}" not found`
                            });
                        }
                    }

                    // Return all sections if no title filter
                    return res.status(200).json({
                        status: 'success',
                        data: {
                            sections: h1Sections
                        }
                    });
                } catch (extractError) {
                    return res.status(500).json({
                        status: 'error',
                        message: `Error processing Word document: ${extractError.message}`
                    });
                }
            } else {
                return res.status(400).json({
                    status: 'error',
                    message: 'Invalid file format. Please upload a Word document (.docx)'
                });
            }
        }

        // Handle JSON input with Base64 encoded content
        if (isJSON) {
            const contentType = jsonBody['$content-type'];
            const base64Content = jsonBody['$content'];

            if (!base64Content) {
                return res.status(400).json({
                    status: 'error',
                    message: 'Missing $content in request body'
                });
            }

            try {
                const decodedBuffer = Buffer.from(base64Content, 'base64');

                // Check if decoded content is a Word document
                if (decodedBuffer[0] === 0x50 && decodedBuffer[1] === 0x4B) {
                    const result = await mammoth.convertToHtml({
                        buffer: decodedBuffer
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

                    // Extract all H1 sections
                    const h1Sections = extractAllH1Sections(htmlContent);

                    // If title filter is provided, return only that section with MESSAGEFROMCEO template
                    if (titleFilter) {
                        const matchedSection = h1Sections.find(section =>
                            section.title.toLowerCase().includes(titleFilter.toLowerCase())
                        );

                        if (matchedSection) {
                            // Escape the HTML for JSON string
                            const escapedContent = matchedSection.content
                                .replace(/\\/g, '\\\\')
                                .replace(/"/g, '\\"')
                                .replace(/\n/g, '\\n');

                            // Replace ##MESSAGE## placeholder in MESSAGEFROMCEO template
                            let canvasContent = MESSAGEFROMCEO.CanvasContent1.replace('##MESSAGE##', escapedContent);

                            const response = {
                                ...MESSAGEFROMCEO,
                                CanvasContent1: canvasContent
                            };

                            return res.status(200).json({
                                status: 'success',
                                message: JSON.stringify(response)
                            });
                        } else {
                            return res.status(404).json({
                                status: 'error',
                                message: `Section with title "${titleFilter}" not found`
                            });
                        }
                    }

                    // Return all sections if no title filter
                    return res.status(200).json({
                        status: 'success',
                        data: {
                            sections: h1Sections
                        }
                    });
                } else {
                    // Plain text content
                    const decodedText = decodedBuffer.toString('utf8');
                    return res.status(200).json({
                        status: 'success',
                        data: {
                            content: decodedText
                        }
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

// SharePoint section endpoint - uses SHAREPOINT_SECTION_TEMPLATE with ##MESSAGE##
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

        // Handle raw binary Word document
        if (!isJSON) {
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
                                message: `Section with title "${titleFilter}" not found`
                            });
                        }
                    }

                    // Check if companypolicy template should be used
                    const useCompanyPolicy = req.query.companypolicy !== undefined || jsonBody?.companypolicy !== undefined;
                    let selectedTemplate = useCompanyPolicy ? COMPANYPOLICYTEMPLATE : SHAREPOINT_SECTION_TEMPLATE;

                    // Apply employee paragraph styles to content
                    contentToInsert = applyEmployeeParagraphStyles(contentToInsert);

                    // Extract H2 tags and generate section contents
                    const h2Sections = extractH2Sections(contentToInsert);
                    const sectionContentsHTML = generateSectionContentsHTML(h2Sections, 3);

                    // Escape the HTML for JSON string (preserve newlines and formatting)
                    const escapedMessage = contentToInsert
                        .replace(/\\/g, '\\\\')
                        .replace(/"/g, '\\"')
                        .replace(/\n/g, '\\n');

                    // Replace ##MESSAGE## placeholder with converted HTML
                    let canvasContent = selectedTemplate.CanvasContent1
                        .replace('##MESSAGE##', escapedMessage);

                    // Replace the section contents if H2 sections were found
                    if (sectionContentsHTML && useCompanyPolicy) {
                        // For COMPANYPOLICYTEMPLATE, replace the specific section contents area
                        const sectionContentsPattern = /<p class=\\"noSpacingAbove noSpacingBelow\\" style=\\"color:rgb\(50, 49, 48\);margin-left:0px;\\" data-text-type=\\"noSpacing\\"><span class=\\"fontSizeMediumPlus\\"><span lang=\\"EN-US\\" dir=\\"ltr\\"><strong>3<\/strong><\/span><\/span><a href=\\"#equal-employment-opportunity-%28illinois%29\\"><span class=\\"fontSizeMediumPlus\\"><span lang=\\"EN-US\\" dir=\\"ltr\\"><strong>\.1 Section 1<\/strong><\/span><\/span><\/a><\/p><p class=\\"noSpacingAbove noSpacingBelow\\" style=\\"color:rgb\(50, 49, 48\);margin-left:0px;\\" data-text-type=\\"noSpacing\\"><span class=\\"fontSizeMediumPlus\\"><span lang=\\"EN-US\\" dir=\\"ltr\\"><strong>3<\/strong><\/span><\/span><a href=\\"#equal-employment-opportunity-%28illinois%29\\"><span class=\\"fontSizeMediumPlus\\"><span lang=\\"EN-US\\" dir=\\"ltr\\"><strong>\.2 Section 2<\/strong><\/span><\/span><\/a><\/p>/;
                        canvasContent = canvasContent.replace(sectionContentsPattern, sectionContentsHTML);
                    }

                    const response = {
                        ...selectedTemplate,
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
                                message: `Section with title "${titleFilter}" not found`
                            });
                        }
                    }

                    // Check if companypolicy template should be used
                    const useCompanyPolicy = req.query.companypolicy !== undefined || jsonBody.companypolicy !== undefined;
                    let selectedTemplate = useCompanyPolicy ? COMPANYPOLICYTEMPLATE : SHAREPOINT_SECTION_TEMPLATE;

                    // Apply employee paragraph styles to content
                    contentToInsert = applyEmployeeParagraphStyles(contentToInsert);

                    // Extract H2 tags and generate section contents
                    const h2Sections = extractH2Sections(contentToInsert);
                    const sectionContentsHTML = generateSectionContentsHTML(h2Sections, 3);

                    // Escape the HTML for JSON string
                    const escapedMessage = contentToInsert
                        .replace(/\\/g, '\\\\')
                        .replace(/"/g, '\\"')
                        .replace(/\n/g, '\\n');

                    // Replace ##MESSAGE## placeholder
                    let canvasContent = selectedTemplate.CanvasContent1
                        .replace('##MESSAGE##', escapedMessage);

                    // Replace the section contents if H2 sections were found
                    if (sectionContentsHTML && useCompanyPolicy) {
                        // For COMPANYPOLICYTEMPLATE, replace the specific section contents area
                        const sectionContentsPattern = /<p class=\\"noSpacingAbove noSpacingBelow\\" style=\\"color:rgb\(50, 49, 48\);margin-left:0px;\\" data-text-type=\\"noSpacing\\"><span class=\\"fontSizeMediumPlus\\"><span lang=\\"EN-US\\" dir=\\"ltr\\"><strong>3<\/strong><\/span><\/span><a href=\\"#equal-employment-opportunity-%28illinois%29\\"><span class=\\"fontSizeMediumPlus\\"><span lang=\\"EN-US\\" dir=\\"ltr\\"><strong>\.1 Section 1<\/strong><\/span><\/span><\/a><\/p><p class=\\"noSpacingAbove noSpacingBelow\\" style=\\"color:rgb\(50, 49, 48\);margin-left:0px;\\" data-text-type=\\"noSpacing\\"><span class=\\"fontSizeMediumPlus\\"><span lang=\\"EN-US\\" dir=\\"ltr\\"><strong>3<\/strong><\/span><\/span><a href=\\"#equal-employment-opportunity-%28illinois%29\\"><span class=\\"fontSizeMediumPlus\\"><span lang=\\"EN-US\\" dir=\\"ltr\\"><strong>\.2 Section 2<\/strong><\/span><\/span><\/a><\/p>/;
                        canvasContent = canvasContent.replace(sectionContentsPattern, sectionContentsHTML);
                    }

                    const response = {
                        ...selectedTemplate,
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

// Company Policy endpoint - uses COMPANYPOLICYTEMPLATE with ##MESSAGE##
app.post('/api/companypolicy', bodyParser.raw({ type: '*/*', limit: '50mb' }), async (req, res) => {
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

        // Handle raw binary Word document
        if (!isJSON) {
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
                                message: `Section with title "${titleFilter}" not found`
                            });
                        }
                    }

                    // Apply employee paragraph styles to content
                    contentToInsert = applyEmployeeParagraphStyles(contentToInsert);

                    // Escape the HTML for JSON string (preserve newlines and formatting)
                    const escapedMessage = contentToInsert
                        .replace(/\\/g, '\\\\')
                        .replace(/"/g, '\\"')
                        .replace(/\n/g, '\\n');

                    // Replace ##MESSAGE## placeholder with converted HTML
                    let canvasContent = COMPANYPOLICYTEMPLATE.CanvasContent1
                        .replace('##MESSAGE##', escapedMessage);

                    const response = {
                        ...COMPANYPOLICYTEMPLATE,
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
                                message: `Section with title "${titleFilter}" not found`
                            });
                        }
                    }

                    // Apply employee paragraph styles to content
                    contentToInsert = applyEmployeeParagraphStyles(contentToInsert);

                    // Escape the HTML for JSON string
                    const escapedMessage = contentToInsert
                        .replace(/\\/g, '\\\\')
                        .replace(/"/g, '\\"')
                        .replace(/\n/g, '\\n');

                    // Replace ##MESSAGE## placeholder
                    let canvasContent = COMPANYPOLICYTEMPLATE.CanvasContent1
                        .replace('##MESSAGE##', escapedMessage);

                    const response = {
                        ...COMPANYPOLICYTEMPLATE,
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

// Compensation and Benefits endpoint - uses COMPENSATIONBENEFITSTEMPLATE with ##MESSAGE##
app.post('/api/compensationbenefits', bodyParser.raw({ type: '*/*', limit: '50mb' }), async (req, res) => {
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

        // Handle raw binary Word document
        if (!isJSON) {
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
                                message: `Section with title "${titleFilter}" not found`
                            });
                        }
                    }

                    // Apply employee paragraph styles to content
                    contentToInsert = applyEmployeeParagraphStyles(contentToInsert);

                    // Escape the HTML for JSON string (preserve newlines and formatting)
                    const escapedMessage = contentToInsert
                        .replace(/\\/g, '\\\\')
                        .replace(/"/g, '\\"')
                        .replace(/\n/g, '\\n');

                    // Replace ##MESSAGE## placeholder with converted HTML
                    let canvasContent = COMPENSATIONBENEFITSTEMPLATE.CanvasContent1
                        .replace('##MESSAGE##', escapedMessage);

                    const response = {
                        ...COMPENSATIONBENEFITSTEMPLATE,
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
                                message: `Section with title "${titleFilter}" not found`
                            });
                        }
                    }

                    // Apply employee paragraph styles to content
                    contentToInsert = applyEmployeeParagraphStyles(contentToInsert);

                    // Escape the HTML for JSON string
                    const escapedMessage = contentToInsert
                        .replace(/\\/g, '\\\\')
                        .replace(/"/g, '\\"')
                        .replace(/\n/g, '\\n');

                    // Replace ##MESSAGE## placeholder
                    let canvasContent = COMPENSATIONBENEFITSTEMPLATE.CanvasContent1
                        .replace('##MESSAGE##', escapedMessage);

                    const response = {
                        ...COMPENSATIONBENEFITSTEMPLATE,
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

// Work Performance endpoint - uses WORKPERFORMANCETEMPLATE with ##MESSAGE##
app.post('/api/workperformance', bodyParser.raw({ type: '*/*', limit: '50mb' }), async (req, res) => {
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

        // Handle raw binary Word document
        if (!isJSON) {
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
                                message: `Section with title "${titleFilter}" not found`
                            });
                        }
                    }

                    // Apply employee paragraph styles to content
                    contentToInsert = applyEmployeeParagraphStyles(contentToInsert);

                    // Escape the HTML for JSON string (preserve newlines and formatting)
                    const escapedMessage = contentToInsert
                        .replace(/\\/g, '\\\\')
                        .replace(/"/g, '\\"')
                        .replace(/\n/g, '\\n');

                    // Replace ##MESSAGE## placeholder with converted HTML
                    let canvasContent = WORKPERFORMANCETEMPLATE.CanvasContent1
                        .replace('##MESSAGE##', escapedMessage);

                    const response = {
                        ...WORKPERFORMANCETEMPLATE,
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
                                message: `Section with title "${titleFilter}" not found`
                            });
                        }
                    }

                    // Apply employee paragraph styles to content
                    contentToInsert = applyEmployeeParagraphStyles(contentToInsert);

                    // Escape the HTML for JSON string
                    const escapedMessage = contentToInsert
                        .replace(/\\/g, '\\\\')
                        .replace(/"/g, '\\"')
                        .replace(/\n/g, '\\n');

                    // Replace ##MESSAGE## placeholder
                    let canvasContent = WORKPERFORMANCETEMPLATE.CanvasContent1
                        .replace('##MESSAGE##', escapedMessage);

                    const response = {
                        ...WORKPERFORMANCETEMPLATE,
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

// Procedures and Guidelines endpoint - uses PROCEDURESGUIDELINESTEMPLATE with ##MESSAGE##
app.post('/api/proceduresguidelines', bodyParser.raw({ type: '*/*', limit: '50mb' }), async (req, res) => {
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

        // Handle raw binary Word document
        if (!isJSON) {
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
                        ].join("\n"),
                        convertImage: mammoth.images.imgElement(function (image) {
                            return image.read("base64").then(function (imageBuffer) {
                                return {
                                    src: "data:" + image.contentType + ";base64," + imageBuffer
                                };
                            });
                        })
                    });
                    const htmlContent = result.value || '';

                    // Check for optional title filter
                    const rawTitle = req.query.title || jsonBody?.title || jsonBody?.['$title'];
                    const titleFilter = rawTitle ? rawTitle.trim() : undefined;
                    const h1Sections = extractAllH1Sections(htmlContent);

                    let contentToInsert;

                    if (titleFilter) {
                        // Return only the specific H1 section that matches the title
                        const matchedSection = h1Sections.find(section =>
                            section.title.toLowerCase().includes(titleFilter.toLowerCase())
                        );

                        if (matchedSection) {
                            contentToInsert = matchedSection.content;
                        } else {
                            return res.status(404).json({
                                status: 'error',
                                message: `Section with title "${titleFilter}" not found`
                            });
                        }
                    } else {
                        // Return all H1 sections concatenated together
                        contentToInsert = h1Sections.map(section => section.content).join('');
                    }

                    // Apply employee paragraph styles to content
                    contentToInsert = applyEmployeeParagraphStyles(contentToInsert);

                    // Escape the HTML for JSON string (preserve newlines and formatting)
                    const escapedMessage = contentToInsert
                        .replace(/\\/g, '\\\\')
                        .replace(/"/g, '\\"')
                        .replace(/\n/g, '\\n');

                    // Replace ##MESSAGE## placeholder with converted HTML
                    let canvasContent = PROCEDURESGUIDELINESTEMPLATE.CanvasContent1
                        .replace('##MESSAGE##', escapedMessage);

                    const response = {
                        ...PROCEDURESGUIDELINESTEMPLATE,
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
                        styleMap: [
                            "b => strong",
                            "i => em",
                            "u => u",
                            "p[style-name='Heading 3'] => strong",
                            "p[style-name='Heading 4'] => strong",
                            "p[style-name='Heading 5'] => strong",
                            "p[style-name='List Paragraph'] => p:fresh > strong"
                        ].join("\n"),
                        convertImage: mammoth.images.imgElement(function (image) {
                            return image.read("base64").then(function (imageBuffer) {
                                return {
                                    src: "data:" + image.contentType + ";base64," + imageBuffer
                                };
                            });
                        })
                    });
                    const htmlContent = result.value || '';

                    // Check for optional title filter
                    const rawTitle = req.query.title || jsonBody.title || jsonBody['$title'];
                    const titleFilter = rawTitle ? rawTitle.trim() : undefined;
                    const h1Sections = extractAllH1Sections(htmlContent);

                    let contentToInsert;

                    if (titleFilter) {
                        // Return only the specific H1 section that matches the title
                        const matchedSection = h1Sections.find(section =>
                            section.title.toLowerCase().includes(titleFilter.toLowerCase())
                        );

                        if (matchedSection) {
                            contentToInsert = matchedSection.content;
                        } else {
                            return res.status(404).json({
                                status: 'error',
                                message: `Section with title "${titleFilter}" not found`
                            });
                        }
                    } else {
                        // Return all H1 sections concatenated together
                        contentToInsert = h1Sections.map(section => section.content).join('');
                    }

                    // Apply employee paragraph styles to content
                    contentToInsert = applyEmployeeParagraphStyles(contentToInsert);

                    // Escape the HTML for JSON string
                    const escapedMessage = contentToInsert
                        .replace(/\\/g, '\\\\')
                        .replace(/"/g, '\\"')
                        .replace(/\n/g, '\\n');

                    // Replace ##MESSAGE## placeholder
                    let canvasContent = PROCEDURESGUIDELINESTEMPLATE.CanvasContent1
                        .replace('##MESSAGE##', escapedMessage);

                    const response = {
                        ...PROCEDURESGUIDELINESTEMPLATE,
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

// Employee Acknowledgment endpoint with H1 extraction
app.post('/api/employeeacknowledgment', bodyParser.raw({ type: '*/*', limit: '50mb' }), async (req, res) => {
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

        // Handle raw binary Word document
        if (!isJSON) {
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
                        ].join("\n"),
                        convertImage: mammoth.images.imgElement(function (image) {
                            return image.read("base64").then(function (imageBuffer) {
                                return {
                                    src: "data:" + image.contentType + ";base64," + imageBuffer
                                };
                            });
                        })
                    });
                    const htmlContent = result.value || '';

                    // Check for optional title filter
                    const rawTitle = req.query.title || jsonBody?.title || jsonBody?.['$title'];
                    const titleFilter = rawTitle ? rawTitle.trim() : undefined;
                    const h1Sections = extractAllH1Sections(htmlContent);

                    let contentToInsert;

                    if (titleFilter) {
                        // Return only the specific H1 section that matches the title
                        const matchedSection = h1Sections.find(section =>
                            section.title.toLowerCase().includes(titleFilter.toLowerCase())
                        );

                        if (matchedSection) {
                            contentToInsert = matchedSection.content;
                        } else {
                            return res.status(404).json({
                                status: 'error',
                                message: `Section with title "${titleFilter}" not found`
                            });
                        }
                    } else {
                        // Return all H1 sections concatenated together
                        contentToInsert = h1Sections.map(section => section.content).join('');
                    }

                    // Apply employee paragraph styles to content
                    contentToInsert = applyEmployeeParagraphStyles(contentToInsert);

                    // Escape the HTML for JSON string (preserve newlines and formatting)
                    const escapedMessage = contentToInsert
                        .replace(/\\/g, '\\\\')
                        .replace(/"/g, '\\"')
                        .replace(/\n/g, '\\n');

                    // Replace ##MESSAGE## placeholder with converted HTML
                    let canvasContent = EMPLOYEEACKNOWLEDGMENTTEMPLATE.CanvasContent1
                        .replace('##MESSAGE##', escapedMessage);

                    const response = {
                        ...EMPLOYEEACKNOWLEDGMENTTEMPLATE,
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
                        styleMap: [
                            "b => strong",
                            "i => em",
                            "u => u",
                            "p[style-name='Heading 3'] => strong",
                            "p[style-name='Heading 4'] => strong",
                            "p[style-name='Heading 5'] => strong",
                            "p[style-name='List Paragraph'] => p:fresh > strong"
                        ].join("\n"),
                        convertImage: mammoth.images.imgElement(function (image) {
                            return image.read("base64").then(function (imageBuffer) {
                                return {
                                    src: "data:" + image.contentType + ";base64," + imageBuffer
                                };
                            });
                        })
                    });
                    const htmlContent = result.value || '';

                    // Check for optional title filter
                    const rawTitle = req.query.title || jsonBody.title || jsonBody['$title'];
                    const titleFilter = rawTitle ? rawTitle.trim() : undefined;
                    const h1Sections = extractAllH1Sections(htmlContent);

                    let contentToInsert;

                    if (titleFilter) {
                        // Return only the specific H1 section that matches the title
                        const matchedSection = h1Sections.find(section =>
                            section.title.toLowerCase().includes(titleFilter.toLowerCase())
                        );

                        if (matchedSection) {
                            contentToInsert = matchedSection.content;
                        } else {
                            return res.status(404).json({
                                status: 'error',
                                message: `Section with title "${titleFilter}" not found`
                            });
                        }
                    } else {
                        // Return all H1 sections concatenated together
                        contentToInsert = h1Sections.map(section => section.content).join('');
                    }

                    // Apply employee paragraph styles to content
                    contentToInsert = applyEmployeeParagraphStyles(contentToInsert);

                    // Escape the HTML for JSON string
                    const escapedMessage = contentToInsert
                        .replace(/\\/g, '\\\\')
                        .replace(/"/g, '\\"')
                        .replace(/\n/g, '\\n');

                    // Replace ##MESSAGE## placeholder
                    let canvasContent = EMPLOYEEACKNOWLEDGMENTTEMPLATE.CanvasContent1
                        .replace('##MESSAGE##', escapedMessage);

                    const response = {
                        ...EMPLOYEEACKNOWLEDGMENTTEMPLATE,
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

// SharePoint page data endpoint
app.get('/api/page-data', (req, res) => {
    res.status(200).json(SHAREPOINT_TEMPLATE);
});

// SharePoint section template endpoint
app.get('/api/section-template', (req, res) => {
    res.status(200).json(SHAREPOINT_SECTION_TEMPLATE);
});

// Company Policy template endpoint
app.get('/api/companypolicy-template', (req, res) => {
    res.status(200).json(COMPANYPOLICYTEMPLATE);
});

// Message from CEO template endpoint
app.get('/api/messagefromceo-template', (req, res) => {
    res.status(200).json(MESSAGEFROMCEO);
});

// Compensation and Benefits template endpoint
app.get('/api/compensationbenefits-template', (req, res) => {
    res.status(200).json(COMPENSATIONBENEFITSTEMPLATE);
});

// Work Performance template endpoint
app.get('/api/workperformance-template', (req, res) => {
    res.status(200).json(WORKPERFORMANCETEMPLATE);
});

// Procedures and Guidelines template endpoint
app.get('/api/proceduresguidelines-template', (req, res) => {
    res.status(200).json(PROCEDURESGUIDELINESTEMPLATE);
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
    console.log(`POST /api/decode - Process Word doc with SharePoint template`);
    console.log(`POST /api/introduction - Uses MESSAGEFROMCEO template with ##MESSAGE##`);
    console.log(`POST /api/section - Process Word doc with section template`);
    console.log(`POST /api/companypolicy - Process Word doc with company policy template`);
    console.log(`POST /api/compensationbenefits - Process Word doc with compensation benefits template`);
    console.log(`POST /api/workperformance - Process Word doc with work performance template`);
    console.log(`POST /api/proceduresguidelines - Process Word doc with procedures and guidelines template`);
    console.log(`GET /api/page-data - Get SharePoint template`);
    console.log(`GET /api/section-template - Get SharePoint section template`);
    console.log(`GET /api/companypolicy-template - Get Company Policy template`);
    console.log(`GET /api/messagefromceo-template - Get Message from CEO template`);
    console.log(`GET /api/compensationbenefits-template - Get Compensation and Benefits template`);
    console.log(`GET /api/workperformance-template - Get Work Performance template`);
    console.log(`GET /api/proceduresguidelines-template - Get Procedures and Guidelines template`);
    console.log(`GET /health - Health check`);
});
