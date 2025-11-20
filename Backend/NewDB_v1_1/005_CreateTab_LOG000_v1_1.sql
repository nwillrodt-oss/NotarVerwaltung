if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[LOG000]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[LOG000]
GO

CREATE TABLE [dbo].[LOG000] (
	[ID000] [uniqueidentifier] NOT NULL ,
	[Host000] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[User000] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[AppName000] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[Message000] [varchar] (1000) COLLATE Latin1_General_CI_AS NULL ,
	[Timestamp000] [datetime] NULL 
) ON [PRIMARY]
GO

