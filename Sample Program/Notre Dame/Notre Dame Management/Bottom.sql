if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BottomScroll]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BottomScroll]
GO

CREATE TABLE [dbo].[BottomScroll] (
	[ScrollText] [nvarchar] (250) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

