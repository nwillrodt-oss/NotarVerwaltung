if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USER001]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[USER001]
GO

CREATE TABLE [dbo].[USER001] (
	[ID001] [uniqueidentifier] NOT NULL ,
	[USERNAME001] [varchar] (255) COLLATE Latin1_General_CI_AS NOT NULL ,
	[PWD001] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[FK002001] [uniqueidentifier] NULL ,
	[VORNAME001] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[NACHNAME001] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[TEL001] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[FAX001] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[EMAIL001] [varchar] (100) COLLATE Latin1_General_CI_AS NULL ,
	[SYSTEM001] [smallint] NULL ,
	[CREATE001] [datetime] NOT NULL ,
	[MODIFY001] [datetime] NULL ,
	[CFROM001] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[MFROM001] [varchar] (255) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[USER001] ADD 
	CONSTRAINT [DF_USER001_ID001] DEFAULT (newid()) FOR [ID001],
	CONSTRAINT [DF_USER001_SYSTEM001] DEFAULT (0) FOR [SYSTEM001],
	CONSTRAINT [DF_USER001_CREATE001] DEFAULT (getdate()) FOR [CREATE001],
	CONSTRAINT [DF_USER001_MODIFY001] DEFAULT (getdate()) FOR [MODIFY001]
GO

