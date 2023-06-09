if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AttLog]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[AttLog]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Userst]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Userst]
GO

CREATE TABLE [dbo].[AttLog] (
	[IDNumber] [numeric](18, 0) IDENTITY (1, 1) NOT NULL ,
	[ACNumber] [numeric](18, 0) NOT NULL ,
	[LogDT] [datetime] NOT NULL ,
	[ReaderIP] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Userst] (
	[ACNumber] [numeric](18, 0) NULL ,
	[StudentNumber] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[LName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[FName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[MName] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[GradeYear] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Notes] [varchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

