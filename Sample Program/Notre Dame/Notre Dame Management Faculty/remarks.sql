if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Remarks]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Remarks]
GO

CREATE TABLE [dbo].[Remarks] (
	[RemarksID] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ,
	[SSN] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[DTFrom] [datetime] NULL ,
	[DTTo] [datetime] NULL ,
	[Remarks] [nvarchar] (250) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

