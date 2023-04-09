if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Remark]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Remark]
GO

CREATE TABLE [dbo].[Remark] (
	[Remarks] [text] COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DateFrom] [datetime] NULL ,
	[DateTo] [datetime] NULL ,
	[StudentNo] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

