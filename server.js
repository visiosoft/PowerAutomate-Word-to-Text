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

// SharePoint page template with placeholder
const SHAREPOINT_TEMPLATE = {
    "__metadata": {
        "type": "SP.Publishing.SitePage"
    },
    "PageRenderingState": 0,
    "CanvasContent1": "[{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"c12d16a8-30de-46cb-8d44-f22d081aace5\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":1},\"zoneHeight\":550,\"id\":\"3dbcc936-0d52-4658-91f8-dc3cf63925c1\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":0,\"y\":0,\"w\":70,\"h\":7,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<h2 class=\\\"headingSpacingAbove headingSpacingBelow lineHeight1_0\\\"><span class=\\\"fontSizeMega rte-fontscale-font-max\\\">1.1 A&nbsp;Message&nbsp;from the CEO</span></h2>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"c12d16a8-30de-46cb-8d44-f22d081aace5\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":2},\"zoneHeight\":550,\"id\":\"1857d513-039d-4197-8e49-4bd8e6f378c4\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":0,\"y\":6,\"w\":70,\"h\":6,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><span class=\\\"fontSizeXLarge\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\">##CEOMESSAGE##</span></span></p><p class=\\\"noSpacingAbove noSpacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" aria-hidden=\\\"true\\\" data-text-type=\\\"noSpacing\\\">&nbsp;</p><p class=\\\"noSpacingAbove noSpacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" aria-hidden=\\\"true\\\" data-text-type=\\\"noSpacing\\\">&nbsp;</p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"c12d16a8-30de-46cb-8d44-f22d081aace5\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":3},\"zoneHeight\":550,\"id\":\"bd7c16e0-b868-43ce-9742-bd272082eb3c\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":25,\"y\":29,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"bd7c16e0-b868-43ce-9742-bd272082eb3c\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Home\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/HANDBOOK.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Center\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"c12d16a8-30de-46cb-8d44-f22d081aace5\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":4},\"zoneHeight\":550,\"id\":\"2a4fe9c1-c46f-4664-ac1b-b0179329f01b\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":38,\"y\":29,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"2a4fe9c1-c46f-4664-ac1b-b0179329f01b\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Next Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Employment%20Policies%20and%20Practices.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Center\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"0ba22869-c07a-4c44-bccc-521aae856fe8\",\"sectionIndex\":1,\"sectionFactor\":8,\"controlIndex\":1},\"id\":\"48047102-2fb5-4efc-ac96-034716e6f250\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"innerHTML\":\"<p><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\">##COMPANYINFO##</span></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"0ba22869-c07a-4c44-bccc-521aae856fe8\",\"sectionIndex\":2,\"sectionFactor\":4,\"controlIndex\":1},\"id\":\"4696d89e-ec85-4c9b-8ebf-e04f59ab5343\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"innerHTML\":\"<p class=\\\"noSpacingAbove noSpacingBelow\\\" data-text-type=\\\"noSpacing\\\"><span class=\\\"fontSizeXLargePlus\\\"><span class=\\\"fontColorThemeDarker\\\"><strong>Section Contents</strong></span></span></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"0ba22869-c07a-4c44-bccc-521aae856fe8\",\"sectionIndex\":2,\"sectionFactor\":4,\"controlIndex\":2},\"id\":\"da32caa7-5244-4649-8293-35d10e746a80\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"innerHTML\":\"<p class=\\\"lineHeight1_2\\\" style=\\\"margin-left:0px;\\\"><a href=\\\"/sites/InnovationArtsEntertainment/SitePages/Introduction.aspx#1.1-a-message-from-the-ceo\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>1.1 </strong>CEO</span> Message</span></a></p><p class=\\\"lineHeight1_2\\\" style=\\\"margin-left:0px;\\\"><a href=\\\"/sites/InnovationArtsEntertainment/SitePages/Introduction.aspx#1.2-the-company\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>1.2</strong> The Company</span>&nbsp;</span></a></p><p class=\\\"lineHeight1_2\\\" style=\\\"margin-left:0px;\\\"><a href=\\\"/sites/InnovationArtsEntertainment/SitePages/Introduction.aspx#1.3-change-in-policy\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>1.3 </strong>Change in Policy</span>&nbsp;</span></a></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":3,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":1},\"zoneHeight\":1600,\"id\":\"9a58ab09-29e4-4d76-b0f5-8b8a114cb40e\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":36,\"y\":87,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"9a58ab09-29e4-4d76-b0f5-8b8a114cb40e\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Next Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Employment%20Policies%20and%20Practices.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Right\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":3,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":2},\"zoneHeight\":1600,\"id\":\"602be227-1a60-4fd8-a4bd-3f69c27a0fa8\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":3,\"y\":87,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"602be227-1a60-4fd8-a4bd-3f69c27a0fa8\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Home\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/HANDBOOK.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Center\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":3,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":3},\"zoneHeight\":1600,\"id\":\"8048e3e4-74fd-4e6b-9d4f-cfb5649079d8\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":5,\"y\":85,\"w\":40,\"h\":0,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"2161a1c6-db61-4731-b97c-3cdb303f7cbb\",\"instanceId\":\"8048e3e4-74fd-4e6b-9d4f-cfb5649079d8\",\"title\":\"Divider\",\"description\":\"Add a line to divide areas on your page.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{},\"links\":{}},\"dataVersion\":\"1.2\",\"properties\":{\"minimumLayoutWidth\":1,\"length\":100,\"weight\":1},\"containsDynamicDataSource\":false},\"webPartId\":\"2161a1c6-db61-4731-b97c-3cdb303f7cbb\",\"reservedWidth\":688,\"reservedHeight\":1},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isAIGeneratedDescription\":false,\"isDefaultThumbnail\":true,\"isSpellCheckEnabled\":true,\"globalRichTextStylingVersion\":1,\"rtePageSettings\":{\"contentVersion\":5,\"indentationVersion\":2},\"isEmailReady\":true,\"webPartsPageSettings\":{\"isTitleHeadingLevelsEnabled\":true,\"isLowQualityImagePlaceholderEnabled\":true}}},{\"controlType\":14,\"webPartData\":{\"properties\":{\"zoneBackground\":{\"c12d16a8-30de-46cb-8d44-f22d081aace5\":{\"type\":\"image\",\"imageData\":{\"fileName\":\"68234-IAE_Display_28x22_FINAL.jpg\",\"width\":2016,\"height\":720,\"source\":0,\"siteId\":\"44583a03-7429-4de6-9641-ae749e56727f\",\"webId\":\"454df1df-1248-4cb7-a41d-a4e482d69fbb\",\"listId\":\"2ab24772-b566-4fe3-abab-a9edde27e5ea\",\"uniqueId\":\"df6e1993-3198-4e40-9f91-db508a198a82\"},\"overlay\":{\"color\":\"#000000\",\"opacity\":0},\"useLightText\":true},\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\":{\"type\":\"image\",\"imageData\":{\"fileName\":\"template11Image0003.png\",\"width\":144,\"height\":144,\"source\":1},\"overlay\":{\"color\":\"#FFFFFF\",\"opacity\":100},\"useLightText\":false}},\"zoneThemeIndex\":{\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\":{\"lightIndex\":0,\"darkIndex\":0}}},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{\"zoneBackground.c12d16a8-30de-46cb-8d44-f22d081aace5.imageData.url\":\"/sites/InnovationArtsEntertainment/SiteAssets/SitePages/VolunteerCenter(1)/68234-IAE_Display_28x22_FINAL.jpg\",\"zoneBackground.54f8a2d2-b737-4a6e-96a2-54ae0a01c7da.imageData.url\":\"https://cdn.hubblecontent.osi.office.net/m365content/publish/4b99b582-026b-4490-ab73-b0091ab721cc/image.png\"},\"links\":{}},\"dataVersion\":\"1.0\"}}]"
}

// SharePoint section template with ##MESSAGE## placeholder
const SHAREPOINT_SECTION_TEMPLATE = {
    "__metadata": {
        "type": "SP.Publishing.SitePage"
    },
    "PageRenderingState": 0,
    "CanvasContent1": "[{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"56eb74b5-aa11-4125-b709-c075b5fb9b35\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":1},\"zoneHeight\":241,\"id\":\"3dbcc936-0d52-4658-91f8-dc3cf63925c1\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":0,\"y\":2,\"w\":70,\"h\":7,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;text-align:center;\\\" data-text-type=\\\"withSpacing\\\"><span class=\\\"fontSizeMega rte-fontscale-font-max\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>Employment Policies and Practices</strong></span><strong>&nbsp;</strong></span></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"56eb74b5-aa11-4125-b709-c075b5fb9b35\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":2},\"zoneHeight\":241,\"id\":\"064d4a8b-0c85-415e-b306-31b3f5d411d4\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":30,\"y\":10,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"064d4a8b-0c85-415e-b306-31b3f5d411d4\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Home\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/HANDBOOK.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Center\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"56eb74b5-aa11-4125-b709-c075b5fb9b35\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":3},\"zoneHeight\":241,\"id\":\"abf4a27b-49df-4fbf-82d9-868ef2a43fc5\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":17,\"y\":10,\"w\":11,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"abf4a27b-49df-4fbf-82d9-868ef2a43fc5\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Previous Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Introduction.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Left\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":189,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"56eb74b5-aa11-4125-b709-c075b5fb9b35\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":4},\"zoneHeight\":241,\"id\":\"048cb962-5cba-402b-892d-d9dc44421c1d\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":42,\"y\":10,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"048cb962-5cba-402b-892d-d9dc44421c1d\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Next Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Company%20Policies%20and%20Practices.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Right\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":1},\"zoneHeight\":3612,\"id\":\"48047102-2fb5-4efc-ac96-034716e6f250\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":4,\"y\":1,\"w\":44,\"h\":3,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<p><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\">##MESSAGE##</span></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":2},\"zoneHeight\":3612,\"id\":\"4696d89e-ec85-4c9b-8ebf-e04f59ab5343\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":52,\"y\":1,\"w\":14,\"h\":3,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<p class=\\\"noSpacingAbove noSpacingBelow\\\" data-text-type=\\\"noSpacing\\\"><span class=\\\"fontSizeXLargePlus\\\"><span class=\\\"fontColorThemeDarker\\\"><strong>Section Contents</strong></span></span></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":3},\"zoneHeight\":3612,\"id\":\"da32caa7-5244-4649-8293-35d10e746a80\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":52,\"y\":3,\"w\":17,\"h\":16,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"/sites/InnovationArtsEntertainment/SitePages/Employment%20Policies%20and%20Practices.aspx#2.1.-equal-employment-opportunity-%28illinois%29\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>2.1</strong> Equal Employment Opportunity (Illinois employees)</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"/sites/InnovationArtsEntertainment/SitePages/Employment%20Policies%20and%20Practices.aspx#2.2.-anti-discrimination-and-anti-sexual-harassment-policies\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>2.2</strong> Anti-Discrimination and Anti-Sexual Harassment Policies</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"/sites/InnovationArtsEntertainment/SitePages/Employment%20Policies%20and%20Practices.aspx#2.3.-at-will-employment\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>2.3</strong> At-Will Employment</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"/sites/InnovationArtsEntertainment/SitePages/Employment%20Policies%20and%20Practices.aspx#2.4.-employee-classifications\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>2.4</strong> Employee Classifications (exempt, full-time)</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"/sites/InnovationArtsEntertainment/SitePages/Employment%20Policies%20and%20Practices.aspx#2.5.-confidentiality-and-trade-secrets\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>2.5</strong> Confidentiality and Trade Secrets</span>&nbsp;</span></a></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":4},\"zoneHeight\":3612,\"id\":\"9a58ab09-29e4-4d76-b0f5-8b8a114cb40e\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":38,\"y\":208,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"9a58ab09-29e4-4d76-b0f5-8b8a114cb40e\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Next Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Company%20Policies%20and%20Practices.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Right\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":5},\"zoneHeight\":3612,\"id\":\"602be227-1a60-4fd8-a4bd-3f69c27a0fa8\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":21,\"y\":208,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"602be227-1a60-4fd8-a4bd-3f69c27a0fa8\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Home\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/HANDBOOK.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Center\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":6},\"zoneHeight\":3612,\"id\":\"22a99bb7-d8fb-4e17-96cf-2c0d90d5988f\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":5,\"y\":208,\"w\":11,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"22a99bb7-d8fb-4e17-96cf-2c0d90d5988f\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Previous Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Introduction.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Left\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":189,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":7},\"zoneHeight\":3612,\"id\":\"8048e3e4-74fd-4e6b-9d4f-cfb5649079d8\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":6,\"y\":186,\"w\":40,\"h\":0,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"2161a1c6-db61-4731-b97c-3cdb303f7cbb\",\"instanceId\":\"8048e3e4-74fd-4e6b-9d4f-cfb5649079d8\",\"title\":\"Divider\",\"description\":\"Add a line to divide areas on your page.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{},\"links\":{}},\"dataVersion\":\"1.2\",\"properties\":{\"minimumLayoutWidth\":1,\"length\":100,\"weight\":1},\"containsDynamicDataSource\":false},\"webPartId\":\"2161a1c6-db61-4731-b97c-3cdb303f7cbb\",\"reservedWidth\":688,\"reservedHeight\":1},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isAIGen eratedDescription\":false,\"isDefaultThumbnail\":true,\"isSpellCheckEnabled\":true,\"globalRichTextStylingVersion\":1,\"rtePageSettings\":{\"contentVersion\":5,\"indentationVersion\":2},\"isEmailReady\":true,\"webPartsPageSettings\":{\"isTitleHeadingLevelsEnabled\":true,\"isLowQualityImagePlaceholderEnabled\":true}}},{\"controlType\":14,\"webPartData\":{\"properties\":{\"zoneBackground\":{\"56eb74b5-aa11-4125-b709-c075b5fb9b35\":{\"type\":\"image\",\"imageData\":{\"fileName\":\"68234-IAE_Display_28x22_FINAL.jpg\",\"width\":2016,\"height\":720,\"source\":0,\"siteId\":\"44583a03-7429-4de6-9641-ae749e56727f\",\"webId\":\"454df1df-1248-4cb7-a41d-a4e482d69fbb\",\"listId\":\"2ab24772-b566-4fe3-abab-a9edde27e5ea\",\"uniqueId\":\"df6e1993-3198-4e40-9f91-db508a198a82\"},\"overlay\":{\"color\":\"#000000\",\"opacity\":0},\"useLightText\":true},\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\":{\"type\":\"image\",\"imageData\":{\"fileName\":\"template11Image0003.png\",\"width\":144,\"height\":144,\"source\":1},\"overlay\":{\"color\":\"#FFFFFF\",\"opacity\":100},\"useLightText\":false}},\"zoneThemeIndex\":{\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\":{\"lightIndex\":0,\"darkIndex\":0}}},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{\"zoneBackground.56eb74b5-aa11-4125-b709-c075b5fb9b35.imageData.url\":\"/sites/InnovationArtsEntertainment/SiteAssets/SitePages/VolunteerCenter(1)/68234-IAE_Display_28x22_FINAL.jpg\",\"zoneBackground.54f8a2d2-b737-4a6e-96a2-54ae0a01c7da.imageData.url\":\"https://cdn.hubblecontent.osi.office.net/m365content/publish/4b99b582-026b-4490-ab73-b0091ab721cc/image.png\"},\"links\":{}},\"dataVersion\":\"1.0\"}}]"
}

// Company Policy template with ##MESSAGE## placeholder
const COMPANYPOLICYTEMPLATE = {
    "__metadata": {
        "type": "SP.Publishing.SitePage"
    },
    "PageRenderingState": 0,
    "CanvasContent1": "[{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"f374e17d-f830-40c9-859a-b641fed7e09c\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":1},\"zoneHeight\":344,\"id\":\"3dbcc936-0d52-4658-91f8-dc3cf63925c1\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":0,\"y\":3,\"w\":70,\"h\":13,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<h2 class=\\\"headingSpacingAbove headingSpacingBelow lineHeight2_4\\\" style=\\\"text-align:center;\\\"><span class=\\\"fontSizeMega rte-fontscale-font-max\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\">Company Policies and Practices</span>&nbsp;</span></h2><p class=\\\"noSpacingAbove spacingBelow lineHeight1_3\\\" style=\\\"margin-left:0px;text-align:center;\\\" data-text-type=\\\"withSpacing\\\"><span class=\\\"fontSizeXLarge\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\">The Innovation Arts &amp; Entertainment Employee Handbook is intended to provide&nbsp;employees with a clear understanding of company policies and procedures related to a variety of&nbsp;business areas</span>&nbsp;</span></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;text-align:center;\\\" data-text-type=\\\"withSpacing\\\"><span class=\\\"fontSizeXLarge\\\">&nbsp;</span></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"f374e17d-f830-40c9-859a-b641fed7e09c\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":2},\"zoneHeight\":344,\"id\":\"17459e36-060a-4ca2-99ce-ffbd8d4b59c2\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":19,\"y\":15,\"w\":10,\"h\":3}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"17459e36-060a-4ca2-99ce-ffbd8d4b59c2\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Previous Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Employment%20Policies%20and%20Practices.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Left\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":172,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"f374e17d-f830-40c9-859a-b641fed7e09c\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":3},\"zoneHeight\":344,\"id\":\"ff4e4cae-e93d-4692-af6f-eafd6f0a6083\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":31,\"y\":15,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"ff4e4cae-e93d-4692-af6f-eafd6f0a6083\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Home\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/HANDBOOK.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Center\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"f374e17d-f830-40c9-859a-b641fed7e09c\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":4},\"zoneHeight\":344,\"id\":\"3d2aa629-fe8c-4209-84bb-a6382cee019c\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":41,\"y\":15,\"w\":9,\"h\":3,\"dataVersion\":\" 1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"3d2aa629-fe8c-4209-84bb-a6382cee019c\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Next Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Compensation%20and%20Benefits.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Right\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":1},\"zoneHeight\":13089,\"id\":\"48047102-2fb5-4efc-ac96-034716e6f250\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":4,\"y\":1,\"w\":44,\"h\":3,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<p>##MESSAGE##</p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":2},\"zoneHeight\":13089,\"id\":\"4696d89e-ec85-4c9b-8ebf-e04f59ab5343\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":52,\"y\":1,\"w\":14,\"h\":3,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<p class=\\\"noSpacingAbove noSpacingBelow\\\" data-text-type=\\\"noSpacing\\\"><span class=\\\"fontSizeXLargePlus\\\"><span class=\\\"fontColorThemeDarker\\\"><strong>Section Contents</strong></span></span></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":3},\"zoneHeight\":13089,\"id\":\"da32caa7-5244-4649-8293-35d10e746a80\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":52,\"y\":3,\"w\":17,\"h\":37,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"/sites/InnovationArtsEntertainment/SitePages/Company%20Policies%20and%20Practices.aspx#3.1.-workweek-and-work-schedules\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>3.1 </strong>Workweek and Work Schedules</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"/sites/InnovationArtsEntertainment/SitePages/Company%20Policies%20and%20Practices.aspx#3.2.-access-card-policy\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>3.2</strong> Access Card Policy</span></span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"/sites/InnovationArtsEntertainment/SitePages/Company%20Policies%20and%20Practices.aspx#3.3.-use-of-company-property\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>3.3</strong> Use of Company Property</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"/sites/InnovationArtsEntertainment/SitePages/Company%20Policies%20and%20Practices.aspx#3.4.-technology-systems-and-software\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>3.4 </strong>Technology Systems and Software</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"#expectation-of-privacy\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>3.5</strong> Expectation of Privacy</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"#cell-phone-policy\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>3.6</strong> Cell Phone Policy</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"/sites/InnovationArtsEntertainment/SitePages/Company%20Policies%20and%20Practices.aspx#3.7.-social-media-policy\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>3.7</strong> Social Media Policy</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"/sites/InnovationArtsEntertainment/SitePages/Company%20Policies%20and%20Practices.aspx#3.8.-employee-dating-policy\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>3.8</strong> Employee Dating Policy</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"/sites/InnovationArtsEntertainment/SitePages/Company%20Policies%20and%20Practices.aspx#3.9.-office-voicemail-policy\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>3.9</strong> Office&nbsp;Voicemail Policy</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"/sites/InnovationArtsEntertainment/SitePages/Company%20Policies%20and%20Practices.aspx#3.10.-email-signature\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>3.10</strong> Email Signature&nbsp;Policy</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"/sites/InnovationArtsEntertainment/SitePages/Company%20Policies%20and%20Practices.aspx#3.11.-out-of-office-policy\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>3.11 </strong>Out-of-Office Policy</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"/sites/InnovationArtsEntertainment/SitePages/Company%20Policies%20and%20Practices.aspx#3.12.-distribution-lists%2c-groups%2c-and-teams\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>3.12 </strong>Distribution Lists, Groups and Teams&nbsp;Policy</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"/sites/InnovationArtsEntertainment/SitePages/Company%20Policies%20and%20Practices.aspx#3.13.-labeling-and-coding-policy\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>3.13 </strong>Labeling and Coding Policy</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"/sites/InnovationArtsEntertainment/SitePages/Company%20Policies%20and%20Practices.aspx#3.14.-work-from-home-policy\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>3.14 </strong>Work From Home Policy</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"/sites/InnovationArtsEntertainment/SitePages/Company%20Policies%20and%20Practices.aspx#3.15.-credit-card-policy\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>3.15</strong> Credit Card&nbsp;Policy</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"/sites/InnovationArtsEntertainment/SitePages/Company%20Policies%20and%20Practices.aspx#3.16.-travel-policy\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>3.16</strong> Travel&nbsp;Policy</span>&nbsp;</span></a></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><a href=\\\"/sites/InnovationArtsEntertainment/SitePages/Company%20Policies%20and%20Practices.aspx#3.17.-expense-reporting-and-reimbursement\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>3.17 </strong>Expense Reporting and Reimbursement&nbsp;Policy</span>&nbsp;</span></a></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":4},\"zoneHeight\":13089,\"id\":\"9a58ab09-29e4-4d76-b0f5-8b8a114cb40e\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":31,\"y\":752,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"9a58ab09-29e4-4d76-b0f5-8b8a114cb40e\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Next Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Compensation%20and%20Benefits.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Right\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":5},\"zoneHeight\":13089,\"id\":\"602be227-1a60-4fd8-a4bd-3f69c27a0fa8\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":19,\"y\":752,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"602be227-1a60-4fd8-a4bd-3f69c27a0fa8\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Home\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/HANDBOOK.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Center\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":6},\"zoneHeight\":13089,\"id\":\"22a99bb7-d8fb-4e17-96cf-2c0d90d5988f\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":6,\"y\":752,\"w\":10,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"22a99bb7-d8fb-4e17-96cf-2c0d90d5988f\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Previous Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Employment%20Policies%20and%20Practices.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Left\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":172,\"reservedHeight\":40},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isAIGeneratedDescription\":false,\"isDefaultThumbnail\":true,\"isSpellCheckEnabled\":true,\"globalRichTextStylingVersion\":1,\"rtePageSettings\":{\"contentVersion\":5,\"indentationVersion\":2},\"isEmailReady\":true,\"webPartsPageSettings\":{\"isTitleHeadingLevelsEnabled\":true,\"isLowQualityImagePlaceholderEnabled\":true}}},{\"controlType\":14,\"webPartData\":{\"properties\":{\"zoneBackground\":{\"f374e17d-f830-40c9-859a-b641fed7e09c\":{\"type\":\"image\",\"imageData\":{\"fileName\":\"68234-IAE_Display_28x22_FINAL.jpg\",\"width\":2016,\"height\":720,\"source\":0,\"siteId\":\"44583a03-7429-4de6-9641-ae749e56727f\",\"webId\":\"454df1df-1248-4cb7-a41d-a4e482d69fbb\",\"listId\":\"2ab24772-b566-4fe3-abab-a9edde27e5ea\",\"uniqueId\":\"df6e1993-3198-4e40-9f91-db508a198a82\"},\"overlay\":{\"color\":\"#000000\",\"opacity\":0},\"useLightText\":true},\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\":{\"type\":\"image\",\"imageData\":{\"fileName\":\"template11Image0003.png\",\"width\":144,\"height\":144,\"source\":1},\"overlay\":{\"color\":\"#FFFFFF\",\"opacity\":100},\"useLightText\":false}},\"zoneThemeIndex\":{\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\":{\"lightIndex\":0,\"darkIndex\":0}}},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{\"zoneBackground.f374e17d-f830-40c9-859a-b641fed7e09c.imageData.url\":\"/sites/InnovationArtsEntertainment/SiteAssets/SitePages/VolunteerCenter(1)/68234-IAE_Display_28x22_FINAL.jpg\",\"zoneBackground.54f8a2d2-b737-4a6e-96a2-54ae0a01c7da.imageData.url\":\"https://cdn.hubblecontent.osi.office.net/m365content/publish/4b99b582-026b-4490-ab73-b0091ab721cc/image.png\"},\"links\":{}},\"dataVersion\":\"1.0\"}}]"
}

// Message from CEO template with ##MESSAGE## placeholder
const MESSAGEFROMCEO = {
    "__metadata": {
        "type": "SP.Publishing.SitePage"
    },
    "PageRenderingState": 0,
    "CanvasContent1": "[{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"c12d16a8-30de-46cb-8d44-f22d081aace5\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":1},\"zoneHeight\":550,\"id\":\"3dbcc936-0d52-4658-91f8-dc3cf63925c1\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":0,\"y\":0,\"w\":70,\"h\":7,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<h2 class=\\\"headingSpacingAbove headingSpacingBelow lineHeight1_0\\\"><span class=\\\"fontSizeMega rte-fontscale-font-max\\\">1.1 A&nbsp;Message&nbsp;from the CEO</span></h2>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"c12d16a8-30de-46cb-8d44-f22d081aace5\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":2},\"zoneHeight\":550,\"id\":\"1857d513-039d-4197-8e49-4bd8e6f378c4\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":0,\"y\":6,\"w\":70,\"h\":23,\"dataVersion\":\"1.0\"}},\"innerHTML\":\"<p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><span class=\\\"fontSizeXLarge\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\">Welcome!&nbsp;</span>&nbsp;</span></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><span class=\\\"fontSizeXLarge\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\">You have just joined a dynamic and rapidly growing company. We hope that your employment with Innovation Arts &amp; Entertainment&nbsp;will be&nbsp;both&nbsp;challenging&nbsp;and rewarding.&nbsp;We are&nbsp;proud of the&nbsp;professional&nbsp;services we provide – and&nbsp;we are&nbsp;equally&nbsp;proud of our employees,&nbsp;who enable and enhance our ability to perform&nbsp;our work at&nbsp;a high level.</span> &nbsp;</span></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><span class=\\\"fontSizeXLarge\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\">Our&nbsp;employee&nbsp;handbook is intended to provide employees with important policy&nbsp;and procedure&nbsp;guidelines to learn and follow, as well as a reference source for many aspects of life&nbsp;at&nbsp;Innovation Arts &amp; Entertainment.</span> &nbsp;</span></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><span class=\\\"fontSizeXLarge\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\">Our&nbsp;people are energetic, and our&nbsp;office&nbsp;life&nbsp;is&nbsp;fast paced.&nbsp;Be prepared to learn, grow, and perform important work,&nbsp;daily – and your days here will fly by.</span>&nbsp;</span></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"withSpacing\\\"><span class=\\\"fontSizeXLarge\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\">We&nbsp;wish you&nbsp;remarkable&nbsp;success!</span>&nbsp;</span></p><p class=\\\"noSpacingAbove spacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" aria-hidden=\\\"true\\\" data-text-type=\\\"withSpacing\\\">&nbsp;</p><p class=\\\"noSpacingAbove noSpacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"noSpacing\\\"><span class=\\\"fontSizeXLarge\\\"><i><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>Adam Epstein</strong></span></i></span></p><p class=\\\"noSpacingAbove noSpacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"noSpacing\\\"><span class=\\\"fontSizeXLarge\\\"><i><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\">Founder &amp; CEO</span></i></span></p><p class=\\\"noSpacingAbove noSpacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"noSpacing\\\" aria-hidden=\\\"true\\\">&nbsp;</p><p class=\\\"noSpacingAbove noSpacingBelow lineHeight1_2\\\" style=\\\"margin-left:0px;\\\" data-text-type=\\\"noSpacing\\\" aria-hidden=\\\"true\\\">&nbsp;</p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"c12d16a8-30de-46cb-8d44-f22d081aace5\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":3},\"zoneHeight\":550,\"id\":\"bd7c16e0-b868-43ce-9742-bd272082eb3c\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":25,\"y\":29,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"bd7c16e0-b868-43ce-9742-bd272082eb3c\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Home\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/HANDBOOK.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Center\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":1,\"zoneId\":\"c12d16a8-30de-46cb-8d44-f22d081aace5\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":4},\"zoneHeight\":550,\"id\":\"2a4fe9c1-c46f-4664-ac1b-b0179329f01b\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":38,\"y\":29,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"2a4fe9c1-c46f-4664-ac1b-b0179329f01b\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Next Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Employment%20Policies%20and%20Practices.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Center\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"0ba22869-c07a-4c44-bccc-521aae856fe8\",\"sectionIndex\":1,\"sectionFactor\":8,\"controlIndex\":1},\"id\":\"48047102-2fb5-4efc-ac96-034716e6f250\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"innerHTML\":\"<p class=\\\"noSpacingAbove spacingBelow\\\" data-text-type=\\\"withSpacing\\\"><span style=\\\"color:#0451a5;\\\">##MESSAGE##</span></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"0ba22869-c07a-4c44-bccc-521aae856fe8\",\"sectionIndex\":2,\"sectionFactor\":4,\"controlIndex\":1},\"id\":\"4696d89e-ec85-4c9b-8ebf-e04f59ab5343\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"innerHTML\":\"<p class=\\\"noSpacingAbove noSpacingBelow\\\" data-text-type=\\\"noSpacing\\\"><span class=\\\"fontSizeXLargePlus\\\"><span class=\\\"fontColorThemeDarker\\\"><strong>Section Contents</strong></span></span></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":2,\"zoneId\":\"0ba22869-c07a-4c44-bccc-521aae856fe8\",\"sectionIndex\":2,\"sectionFactor\":4,\"controlIndex\":2},\"id\":\"da32caa7-5244-4649-8293-35d10e746a80\",\"controlType\":4,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"innerHTML\":\"<p class=\\\"lineHeight1_2\\\" style=\\\"margin-left:0px;\\\"><a href=\\\"/sites/InnovationArtsEntertainment/SitePages/Introduction.aspx#1.1-a-message-from-the-ceo\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>1.1 </strong>CEO</span> Message</span></a></p><p class=\\\"lineHeight1_2\\\" style=\\\"margin-left:0px;\\\"><a href=\\\"/sites/InnovationArtsEntertainment/SitePages/Introduction.aspx#1.2-the-company\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>1.2</strong> The Company</span>&nbsp;</span></a></p><p class=\\\"lineHeight1_2\\\" style=\\\"margin-left:0px;\\\"><a href=\\\"/sites/InnovationArtsEntertainment/SitePages/Introduction.aspx#1.3-change-in-policy\\\"><span class=\\\"fontSizeMediumPlus\\\"><span lang=\\\"EN-US\\\" dir=\\\"ltr\\\"><strong>1.3 </strong>Change in Policy</span>&nbsp;</span></a></p>\",\"contentVersion\":5},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":3,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":1},\"zoneHeight\":1600,\"id\":\"9a58ab09-29e4-4d76-b0f5-8b8a114cb40e\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":36,\"y\":87,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"9a58ab09-29e4-4d76-b0f5-8b8a114cb40e\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Next Section\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/Employment%20Policies%20and%20Practices.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Right\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":3,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":2},\"zoneHeight\":1600,\"id\":\"602be227-1a60-4fd8-a4bd-3f69c27a0fa8\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":3,\"y\":87,\"w\":9,\"h\":3,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"instanceId\":\"602be227-1a60-4fd8-a4bd-3f69c27a0fa8\",\"title\":\"Button\",\"description\":\"Add a clickable button with a custom label and link.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{\"label\":\"Home\"},\"imageSources\":{},\"links\":{\"linkUrl\":\"/sites/InnovationArtsEntertainment/SitePages/HANDBOOK.aspx\"}},\"dataVersion\":\"1.1\",\"properties\":{\"alignment\":\"Center\",\"minimumLayoutWidth\":9},\"containsDynamicDataSource\":false},\"webPartId\":\"0f087d7f-520e-42b7-89c0-496aaf979d58\",\"reservedWidth\":155,\"reservedHeight\":40},{\"position\":{\"layoutIndex\":1,\"zoneIndex\":3,\"zoneId\":\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\",\"sectionIndex\":1,\"sectionFactor\":100,\"controlIndex\":3},\"zoneHeight\":1600,\"id\":\"8048e3e4-74fd-4e6b-9d4f-cfb5649079d8\",\"controlType\":3,\"isFromSectionTemplate\":false,\"addedFromPersistedData\":true,\"flexibleLayoutPosition\":{\"lg\":{\"x\":5,\"y\":85,\"w\":40,\"h\":0,\"dataVersion\":\"1.0\"}},\"webPartData\":{\"id\":\"2161a1c6-db61-4731-b97c-3cdb303f7cbb\",\"instanceId\":\"8048e3e4-74fd-4e6b-9d4f-cfb5649079d8\",\"title\":\"Divider\",\"description\":\"Add a line to divide areas on your page.\",\"audiences\":[],\"hideOn\":{\"mobile\":false},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{},\"links\":{}},\"dataVersion\":\"1.2\",\"properties\":{\"minimumLayoutWidth\":1,\"length\":100,\"weight\":1},\"containsDynamicDataSource\":false},\"webPartId\":\"2161a1c6-db61-4731-b97c-3cdb303f7cbb\",\"reservedWidth\":688,\"reservedHeight\":1},{\"controlType\":0,\"pageSettingsSlice\":{\"isDefaultDescription\":true,\"isAIGeneratedDescription\":false,\"isDefaultThumbnail\":true,\"isSpellCheckEnabled\":true,\"globalRichTextStylingVersion\":1,\"rtePageSettings\":{\"contentVersion\":5,\"indentationVersion\":2},\"isEmailReady\":true,\"webPartsPageSettings\":{\"isTitleHeadingLevelsEnabled\":true,\"isLowQualityImagePlaceholderEnabled\":true}}},{\"controlType\":14,\"webPartData\":{\"properties\":{\"zoneBackground\":{\"c12d16a8-30de-46cb-8d44-f22d081aace5\":{\"type\":\"image\",\"imageData\":{\"fileName\":\"68234-IAE_Display_28x22_FINAL.jpg\",\"width\":2016,\"height\":720,\"source\":0,\"siteId\":\"44583a03-7429-4de6-9641-ae749e56727f\",\"webId\":\"454df1df-1248-4cb7-a41d-a4e482d69fbb\",\"listId\":\"2ab24772-b566-4fe3-abab-a9edde27e5ea\",\"uniqueId\":\"df6e1993-3198-4e40-9f91-db508a198a82\"},\"overlay\":{\"color\":\"#000000\",\"opacity\":0},\"useLightText\":true},\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\":{\"type\":\"image\",\"imageData\":{\"fileName\":\"template11Image0003.png\",\"width\":144,\"height\":144,\"source\":1},\"overlay\":{\"color\":\"#FFFFFF\",\"opacity\":100},\"useLightText\":false}},\"zoneThemeIndex\":{\"54f8a2d2-b737-4a6e-96a2-54ae0a01c7da\":{\"lightIndex\":0,\"darkIndex\":0}}},\"serverProcessedContent\":{\"htmlStrings\":{},\"searchablePlainTexts\":{},\"imageSources\":{\"zoneBackground.c12d16a8-30de-46cb-8d44-f22d081aace5.imageData.url\":\"/sites/InnovationArtsEntertainment/SiteAssets/SitePages/VolunteerCenter(1)/68234-IAE_Display_28x22_FINAL.jpg\",\"zoneBackground.54f8a2d2-b737-4a6e-96a2-54ae0a01c7da.imageData.url\":\"https://cdn.hubblecontent.osi.office.net/m365content/publish/4b99b582-026b-4490-ab73-b0091ab721cc/image.png\"},\"links\":{}},\"dataVersion\":\"1.0\"}}]"
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

                    // Escape the HTML for JSON string (preserve newlines and formatting)
                    const escapedMessage = contentToInsert
                        .replace(/\\/g, '\\\\')
                        .replace(/"/g, '\\"')
                        .replace(/\n/g, '\\n');

                    // Replace ##MESSAGE## placeholder with converted HTML
                    let canvasContent = selectedTemplate.CanvasContent1
                        .replace('##MESSAGE##', escapedMessage);

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

                    // Escape the HTML for JSON string
                    const escapedMessage = contentToInsert
                        .replace(/\\/g, '\\\\')
                        .replace(/"/g, '\\"')
                        .replace(/\n/g, '\\n');

                    // Replace ##MESSAGE## placeholder
                    let canvasContent = selectedTemplate.CanvasContent1
                        .replace('##MESSAGE##', escapedMessage);

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
    console.log(`GET /api/page-data - Get SharePoint template`);
    console.log(`GET /api/section-template - Get SharePoint section template`);
    console.log(`GET /api/companypolicy-template - Get Company Policy template`);
    console.log(`GET /api/messagefromceo-template - Get Message from CEO template`);
    console.log(`GET /health - Health check`);
});
