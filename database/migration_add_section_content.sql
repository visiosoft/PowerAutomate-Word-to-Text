/*
===============================================================================
  Migration: Add SectionContent table
  Description:
    Stores processed section API responses so content can be served from DB
    instead of re-processing Word documents on every request. Supports upsert
    by SectionTitle for idempotent saves.
===============================================================================
*/

IF NOT EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = 'dbo' AND TABLE_NAME = 'SectionContent')
BEGIN
    CREATE TABLE SectionContent (
        SectionContentID    INT IDENTITY(1,1) PRIMARY KEY,
        SectionTitle        NVARCHAR(500)   NOT NULL,
        SectionNumber       INT             NULL,
        HtmlContent         NVARCHAR(MAX)   NULL,
        CanvasContent       NVARCHAR(MAX)   NULL,
        APIResponse         NVARCHAR(MAX)   NULL,
        NavigationConfig    NVARCHAR(MAX)   NULL,
        IsActive            BIT             DEFAULT 1,
        CreatedAt           DATETIME2       DEFAULT GETUTCDATE(),
        UpdatedAt           DATETIME2       DEFAULT GETUTCDATE(),

        CONSTRAINT UQ_SectionContent_Title UNIQUE (SectionTitle)
    );

    CREATE INDEX IX_SectionContent_IsActive ON SectionContent(IsActive);

    PRINT '  ✓ SectionContent table created';
END
ELSE
BEGIN
    PRINT '  - SectionContent table already exists';
END;
