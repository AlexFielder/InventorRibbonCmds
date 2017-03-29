CREATE TABLE [dbo].[DETAILS] (
  [INDEX] [int] IDENTITY (1, 1) NOT NULL,
  [CERTIFIED] [varchar] (MAX) NULL,
  [CHECKED] [varchar] (MAX) NULL,
  [DIMS IN] [varchar] (MAX) NULL,
  [DRAWING NUMBER 1] [varchar] (MAX) NULL,
  [DRAWING NUMBER 2] [varchar] (MAX) NULL,
  [DRAWN] [varchar] (MAX) NULL,
  [SHEET NUMBER 1] [varchar] (MAX) NULL,
  [SHEET NUMBER 2] [varchar] (MAX) NULL,
  [NUMBER OF SHEETS 1] [varchar] (MAX) NULL,
  [NUMBER OF SHEETS 2] [varchar] (MAX) NULL,
  [PROTECTIVE FINISH] [varchar] (MAX) NULL,
  [CHANGE NO 1] [varchar] (MAX) NULL,
  [ISSUE 1] [varchar] (MAX) NULL,
  [DATE 1] [varchar] (MAX) NULL,
  [CHANGE NO 2] [varchar] (MAX) NULL,
  [ISSUE 2] [varchar] (MAX) NULL,
  [DATE 2] [varchar] (MAX) NULL,
  [CHANGE NO 3] [varchar] (MAX) NULL,
  [ISSUE 3] [varchar] (MAX) NULL,
  [DATE 3] [varchar] (MAX) NULL,
  [CHANGE NO 4] [varchar] (MAX) NULL,
  [ISSUE 4] [varchar] (MAX) NULL,
  [DATE 4] [varchar] (MAX) NULL,
  [CHANGE NO 5] [varchar] (MAX) NULL,
  [ISSUE 5] [varchar] (MAX) NULL,
  [DATE 5] [varchar] (MAX) NULL,
  [SCALE] [varchar] (MAX) NULL,
  [TITLE 1] [varchar] (MAX) NULL,
  [TITLE 2] [varchar] (MAX) NULL,
  [TITLE 3] [varchar] (MAX) NULL,
  [SURFACE TEXTURE] [varchar] (MAX) NULL,
  [SUB-CONTRACTOR] [varchar] (MAX) NULL,
  [TOLERANCE 1] [varchar] (MAX) NULL,
  [TOLERANCE 2] [varchar] (MAX) NULL,
  [TOLERANCE 3] [varchar] (MAX) NULL,
  [CAD REF] [varchar] (MAX) NULL,
  [USED ON 01 - DRG NUMBER] [varchar] (MAX) NULL,
  [USED ON 01 - PREFIX] [varchar] (MAX) NULL,
  [USED ON 02 - DRG NUMBER] [varchar] (MAX) NULL,
  [USED ON 02 - PREFIX] [varchar] (MAX) NULL,
  [USED ON 03 - DRG NUMBER] [varchar] (MAX) NULL,
  [USED ON 03 - PREFIX] [varchar] (MAX) NULL,
  [DATETIME] [datetime] NULL,
  CONSTRAINT [PK__DETAILS__E61849D86E01572D]
  PRIMARY KEY CLUSTERED
  (
    [INDEX] ASC
  )
  WITH (PAD_INDEX = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, IGNORE_DUP_KEY = OFF, STATISTICS_NORECOMPUTE = OFF, DATA_COMPRESSION = NONE)
  ON [PRIMARY]
) ON [PRIMARY]
TEXTIMAGE_ON [PRIMARY];
GO
