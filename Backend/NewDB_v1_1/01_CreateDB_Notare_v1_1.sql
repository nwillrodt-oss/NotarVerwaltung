IF EXISTS (SELECT name FROM master.dbo.sysdatabases WHERE name = N'Notare')
	DROP DATABASE [Notare]
GO

CREATE DATABASE [Notare]  
 COLLATE Latin1_General_CI_AS
GO

exec sp_dboption N'Notare', N'autoclose', N'false'
GO

exec sp_dboption N'Notare', N'bulkcopy', N'false'
GO

exec sp_dboption N'Notare', N'trunc. log', N'false'
GO

exec sp_dboption N'Notare', N'torn page detection', N'true'
GO

exec sp_dboption N'Notare', N'read only', N'false'
GO

exec sp_dboption N'Notare', N'dbo use', N'false'
GO

exec sp_dboption N'Notare', N'single', N'false'
GO

exec sp_dboption N'Notare', N'autoshrink', N'false'
GO

exec sp_dboption N'Notare', N'ANSI null default', N'false'
GO

exec sp_dboption N'Notare', N'recursive triggers', N'false'
GO

exec sp_dboption N'Notare', N'ANSI nulls', N'false'
GO

exec sp_dboption N'Notare', N'concat null yields null', N'false'
GO

exec sp_dboption N'Notare', N'cursor close on commit', N'false'
GO

exec sp_dboption N'Notare', N'default to local cursor', N'false'
GO

exec sp_dboption N'Notare', N'quoted identifier', N'false'
GO

exec sp_dboption N'Notare', N'ANSI warnings', N'false'
GO

exec sp_dboption N'Notare', N'auto create statistics', N'true'
GO

exec sp_dboption N'Notare', N'auto update statistics', N'true'
GO

if( (@@microsoftversion / power(2, 24) = 8) and (@@microsoftversion & 0xffff >= 724) )
	exec sp_dboption N'Notare', N'db chaining', N'false'
GO

use [Notare]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_AFORT014_FORT011]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[AFORT014] DROP CONSTRAINT FK_AFORT014_FORT011
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_AG004_LG003]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[AG004] DROP CONSTRAINT FK_AG004_LG003
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_LG003_OLG002]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[LG003] DROP CONSTRAINT FK_LG003_OLG002
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_AFORT014_RA010]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[AFORT014] DROP CONSTRAINT FK_AFORT014_RA010
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_AKTENORT017_RA010]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[AKTENORT017] DROP CONSTRAINT FK_AKTENORT017_RA010
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_BEWERB013_RA010]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[BEWERB013] DROP CONSTRAINT FK_BEWERB013_RA010
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_DOC018_RA010]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[DOC018] DROP CONSTRAINT FK_DOC018_RA010
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_BEWERB013_STELLEN012]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[BEWERB013] DROP CONSTRAINT FK_BEWERB013_STELLEN012
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AFORT014]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[AFORT014]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AG004]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[AG004]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AKTENORT017]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[AKTENORT017]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AUSSCHREIBUNG020]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[AUSSCHREIBUNG020]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BERECHNUNGEN016]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BERECHNUNGEN016]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BEWERB013]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[BEWERB013]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DISZ019]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[DISZ019]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DOC018]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[DOC018]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FORD022]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[FORD022]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FORT011]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[FORT011]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FRIST024]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[FRIST024]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[LAND005]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[LAND005]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[LG003]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[LG003]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[LOG000]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[LOG000]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[OLG002]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[OLG002]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[OPTIONS023]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[OPTIONS023]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[RA010]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[RA010]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[STELLEN012]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[STELLEN012]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[USER001]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[USER001]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[UpdateHistory025]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[UpdateHistory025]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[VALUES015]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[VALUES015]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[VORGANG021]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[VORGANG021]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[WORKFLOW006]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[WORKFLOW006]
GO

CREATE TABLE [dbo].[AFORT014] (
	[ID014] [uniqueidentifier] NOT NULL ,
	[FK010014] [uniqueidentifier] NOT NULL ,
	[FK011014] [uniqueidentifier] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[AG004] (
	[ID004] [uniqueidentifier] NOT NULL ,
	[AGNAME004] [varchar] (255) COLLATE Latin1_General_CI_AS NOT NULL ,
	[AGSTR004] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[AGPLZ004] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[AGORT004] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[AGMAIL004] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[AGTEL004] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[AGFAX004] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[FK003004] [uniqueidentifier] NULL ,
	[MODIFY004] [datetime] NULL ,
	[CREATE004] [datetime] NULL ,
	[MFROM004] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[CFROM004] [varchar] (50) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[AKTENORT017] (
	[ID017] [uniqueidentifier] NOT NULL ,
	[FK010017] [uniqueidentifier] NULL ,
	[FK013017] [uniqueidentifier] NULL ,
	[AKTENORT017] [varchar] (255) COLLATE Latin1_General_CI_AS NOT NULL ,
	[CREATE017] [datetime] NOT NULL ,
	[MODIFY017] [datetime] NULL ,
	[CFROM017] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[MFROM017] [varchar] (255) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[AUSSCHREIBUNG020] (
	[ID020] [uniqueidentifier] NOT NULL ,
	[JAHR020] [int] NOT NULL ,
	[AZ020] [varchar] (50) COLLATE Latin1_General_CI_AS NOT NULL ,
	[WORKFLOW020] [varchar] (5) COLLATE Latin1_General_CI_AS NULL ,
	[CREATE020] [datetime] NOT NULL ,
	[MODIFY020] [datetime] NULL ,
	[CFROM020] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[MFROM020] [varchar] (255) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[BERECHNUNGEN016] (
	[ID016] [uniqueidentifier] NOT NULL ,
	[FAKTOR016] [float] NULL ,
	[VALUETYPE016] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[ORDER016] [int] NULL ,
	[MAXWERT016] [int] NULL ,
	[CAPTION016] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[CAPTIONSQL016] [varchar] (1000) COLLATE Latin1_General_CI_AS NULL ,
	[VALUESQL016] [varchar] (1000) COLLATE Latin1_General_CI_AS NULL ,
	[MAXVALUE016] [int] NULL ,
	[LOCKED016] [smallint] NOT NULL ,
	[SAVEFIELD016] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[PUNKTESAVEFIELD016] [varchar] (255) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[BEWERB013] (
	[ID013] [uniqueidentifier] NOT NULL ,
	[FK012013] [uniqueidentifier] NOT NULL ,
	[FK010013] [uniqueidentifier] NOT NULL ,
	[BEM013] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[EINGANG013] [datetime] NOT NULL ,
	[PUNKTE01013] [float] NULL ,
	[PUNKTE02013] [float] NULL ,
	[PUNKTE03013] [float] NULL ,
	[PUNKTE04013] [float] NULL ,
	[PUNKTE05013] [float] NULL ,
	[PUNKTE06013] [float] NULL ,
	[PUNKTE07013] [float] NULL ,
	[PUNKTE08013] [float] NULL ,
	[PUNKTE09013] [float] NULL ,
	[PUNKTE10013] [float] NULL ,
	[PUNKTE11013] [float] NULL ,
	[PUNKTE12013] [float] NULL ,
	[PUNKTE13013] [float] NULL ,
	[PUNKTE14013] [float] NULL ,
	[PUNKTE15013] [float] NULL ,
	[PUNKTE16013] [float] NULL ,
	[PUNKTE17013] [float] NULL ,
	[PUNKTE18013] [float] NULL ,
	[PUNKTE19013] [float] NULL ,
	[PUNKTE20013] [float] NULL ,
	[PUNKTE21013] [float] NULL ,
	[ZUSAGE013] [smallint] NULL ,
	[RANG013] [int] NULL ,
	[CREATE013] [datetime] NOT NULL ,
	[MODIFY013] [datetime] NULL ,
	[CFROM013] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[MFROM013] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[PUNKTESUM013] [float] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[DISZ019] (
	[ID019] [uniqueidentifier] NOT NULL ,
	[FK010019] [uniqueidentifier] NOT NULL ,
	[DATUM019] [datetime] NOT NULL ,
	[MASSNAHME019] [varchar] (255) COLLATE Latin1_General_CI_AS NOT NULL ,
	[VERSTOSS019] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[MUMSTAENDE019] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[EUMSTAENDE019] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[ERGEBNISS019] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[CREATE019] [datetime] NOT NULL ,
	[MODIFY019] [datetime] NULL ,
	[CFROM019] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[MFROM019] [varchar] (50) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[DOC018] (
	[ID018] [uniqueidentifier] NOT NULL ,
	[FK020018] [uniqueidentifier] NULL ,
	[FK010018] [uniqueidentifier] NULL ,
	[FK012018] [uniqueidentifier] NULL ,
	[DOCPATH018] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[TEMPLATE018] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[ALIAS018] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[DOCNAME018] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[CREATE018] [datetime] NOT NULL ,
	[MODIFY018] [datetime] NULL ,
	[CFROM018] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[MFROM018] [varchar] (155) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[FORD022] (
	[ID022] [uniqueidentifier] NOT NULL ,
	[ORDER022] [int] NOT NULL ,
	[GLAEUBIGER022] [varchar] (255) COLLATE Latin1_General_CI_AS NOT NULL ,
	[BETRAG022] [money] NULL ,
	[FK010022] [uniqueidentifier] NOT NULL ,
	[CREATE022] [datetime] NOT NULL ,
	[MODIFY022] [datetime] NULL ,
	[CFROM022] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[MFROM022] [varchar] (255) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[FORT011] (
	[ID011] [uniqueidentifier] NOT NULL ,
	[DATUM011] [datetime] NULL ,
	[VERANSTALTER011] [varchar] (255) COLLATE Latin1_General_CI_AS NOT NULL ,
	[THEMA011] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[ANZHT011] [int] NULL ,
	[ANERKANNT011] [smallint] NULL ,
	[CREATE011] [datetime] NOT NULL ,
	[MODIFY011] [datetime] NULL ,
	[CFROM011] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[MFROM011] [varchar] (255) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[FRIST024] (
	[ID024] [uniqueidentifier] NOT NULL ,
	[FK010024] [uniqueidentifier] NULL ,
	[FK013024] [uniqueidentifier] NULL ,
	[FRIST024] [datetime] NULL ,
	[BEMERKUNG024] [varchar] (55) COLLATE Latin1_General_CI_AS NULL ,
	[CREATE024] [datetime] NOT NULL ,
	[MODIFY024] [datetime] NULL ,
	[CFROM024] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[MFROM024] [varchar] (255) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[LAND005] (
	[ID005] [uniqueidentifier] NOT NULL ,
	[LAND005] [varchar] (255) COLLATE Latin1_General_CI_AS NOT NULL ,
	[LANDKURZ005] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[LANDKFZ005] [varchar] (50) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[LG003] (
	[ID003] [uniqueidentifier] NOT NULL ,
	[LGNAME003] [varchar] (255) COLLATE Latin1_General_CI_AS NOT NULL ,
	[LGSTR003] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[LGPLZ003] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[LGORT003] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[LGMAIL003] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[LGTEL003] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[LGFAX003] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[FK002003] [uniqueidentifier] NULL ,
	[MODIFY003] [datetime] NULL ,
	[CREATE003] [datetime] NULL ,
	[MFROM003] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[CFROM003] [varchar] (50) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
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

CREATE TABLE [dbo].[OLG002] (
	[ID002] [uniqueidentifier] NOT NULL ,
	[OLGNAME002] [varchar] (255) COLLATE Latin1_General_CI_AS NOT NULL ,
	[OLGSTR002] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[OLGPLZ002] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[OLGORT002] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[OLGMAIL002] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[OLGTEL002] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[OLGFAX002] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[MODIFY002] [datetime] NULL ,
	[CREATE002] [datetime] NULL ,
	[MFROM002] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[CFROM002] [varchar] (50) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[OPTIONS023] (
	[ID23] [uniqueidentifier] NOT NULL ,
	[Option023] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[USER023] [varchar] (60) COLLATE Latin1_General_CI_AS NULL ,
	[COMPUTER023] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[KATEGORIE023] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[WERTTEXT023] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[WERTNUM023] [int] NULL ,
	[WERTMEMO023] [text] COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[RA010] (
	[ID010] [uniqueidentifier] NOT NULL ,
	[VORNAME010] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[NACHNAME010] [varchar] (255) COLLATE Latin1_General_CI_AS NOT NULL ,
	[ANREDE010] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[TITEL010] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[NAMEZUS010] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[GESCHL010] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[AZ010] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[GEB010] [datetime] NULL ,
	[VERSTORBEN010] [smallint] NULL ,
	[STAAT010] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[AMTPLZ010] [varchar] (5) COLLATE Latin1_General_CI_AS NULL ,
	[AMTORT010] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[STR010] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[PLZ010] [varchar] (5) COLLATE Latin1_General_CI_AS NULL ,
	[ORT010] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[KSTR010] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[KPLZ010] [varchar] (5) COLLATE Latin1_General_CI_AS NULL ,
	[KORT010] [varchar] (155) COLLATE Latin1_General_CI_AS NULL ,
	[TEL010] [varchar] (20) COLLATE Latin1_General_CI_AS NULL ,
	[FAX010] [varchar] (20) COLLATE Latin1_General_CI_AS NULL ,
	[EMAIL010] [varchar] (100) COLLATE Latin1_General_CI_AS NULL ,
	[KTEL010] [varchar] (20) COLLATE Latin1_General_CI_AS NULL ,
	[KFAX010] [varchar] (20) COLLATE Latin1_General_CI_AS NULL ,
	[KEMAIL010] [varchar] (100) COLLATE Latin1_General_CI_AS NULL ,
	[ANWALTSEIT010] [datetime] NULL ,
	[EXANOTE010] [float] NULL ,
	[WORKFLOW010] [int] NULL ,
	[BEM010] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[STATUS010] [varchar] (50) COLLATE Latin1_General_CI_AS NOT NULL ,
	[AG010] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[LG010] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[BESTELLT010] [datetime] NULL ,
	[AUSGESCH010] [datetime] NULL ,
	[CREATE010] [datetime] NOT NULL ,
	[MODIFY010] [datetime] NULL ,
	[CFROM010] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[MFROM010] [varchar] (255) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[STELLEN012] (
	[ID012] [uniqueidentifier] NOT NULL ,
	[FK020012] [uniqueidentifier] NOT NULL ,
	[BEZIRK012] [varchar] (255) COLLATE Latin1_General_CI_AS NOT NULL ,
	[BESCH012] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[ANZ012] [int] NOT NULL ,
	[FRIST012] [datetime] NOT NULL ,
	[WORKFLOW012] [int] NULL ,
	[CAPTION006] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[CREATE012] [datetime] NOT NULL ,
	[MODIFY012] [datetime] NULL ,
	[CFROM012] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[MFROM012] [varchar] (255) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
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

CREATE TABLE [dbo].[UpdateHistory025] (
	[UpdateDate025] [datetime] NOT NULL ,
	[DBVersion_Vor025] [varchar] (5) COLLATE Latin1_General_CI_AS NULL ,
	[DBVersion_Nach025] [varchar] (5) COLLATE Latin1_General_CI_AS NULL ,
	[Beschreibung025] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[Benutzer025] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[SQL_SRV_Version025] [varchar] (255) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[VALUES015] (
	[ID015] [uniqueidentifier] NOT NULL ,
	[Fieldname015] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[Order015] [int] NULL ,
	[Value015] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[UseValue015] [smallint] NULL ,
	[DestField015] [varchar] (255) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[VORGANG021] (
	[ID021] [uniqueidentifier] NOT NULL ,
	[AZ021] [varchar] (50) COLLATE Latin1_General_CI_AS NULL ,
	[JAHR021] [int] NULL ,
	[CAPTION021] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[STEP021] [int] NULL ,
	[NOTARID021] [uniqueidentifier] NULL ,
	[CREATE021] [datetime] NOT NULL ,
	[MODIFY021] [datetime] NOT NULL ,
	[CFROM021] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[MFROM021] [varchar] (255) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[WORKFLOW006] (
	[ID006] [uniqueidentifier] NOT NULL ,
	[STEP006] [int] NOT NULL ,
	[STEPTITLE006] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[ROOTKEY006] [varchar] (255) COLLATE Latin1_General_CI_AS NOT NULL ,
	[STATUSFIELD006] [varchar] (255) COLLATE Latin1_General_CI_AS NOT NULL ,
	[DESC006] [varchar] (1000) COLLATE Latin1_General_CI_AS NULL ,
	[ACTION006] [varchar] (1000) COLLATE Latin1_General_CI_AS NULL ,
	[CONDITONVALUE006] [varchar] (1000) COLLATE Latin1_General_CI_AS NULL ,
	[CONDITION006] [varchar] (1000) COLLATE Latin1_General_CI_AS NULL ,
	[CONDITIONFROM006] [varchar] (1000) COLLATE Latin1_General_CI_AS NULL ,
	[CONDITIONFAILDMSG006] [varchar] (1000) COLLATE Latin1_General_CI_AS NULL ,
	[CREATE006] [datetime] NOT NULL ,
	[MODIFY006] [datetime] NOT NULL ,
	[CFROM006] [varchar] (255) COLLATE Latin1_General_CI_AS NULL ,
	[MFROM006] [varchar] (255) COLLATE Latin1_General_CI_AS NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[AFORT014] WITH NOCHECK ADD 
	CONSTRAINT [PK_AFORT014] PRIMARY KEY  CLUSTERED 
	(
		[ID014]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[AG004] WITH NOCHECK ADD 
	CONSTRAINT [PK_AG004] PRIMARY KEY  CLUSTERED 
	(
		[ID004]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[AKTENORT017] WITH NOCHECK ADD 
	CONSTRAINT [PK_AKTENORT017] PRIMARY KEY  CLUSTERED 
	(
		[ID017]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[AUSSCHREIBUNG020] WITH NOCHECK ADD 
	CONSTRAINT [PK_AUSSCHREIBUNG020] PRIMARY KEY  CLUSTERED 
	(
		[ID020]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[BERECHNUNGEN016] WITH NOCHECK ADD 
	CONSTRAINT [PK_Berechnungen] PRIMARY KEY  CLUSTERED 
	(
		[ID016]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[BEWERB013] WITH NOCHECK ADD 
	CONSTRAINT [PK_BEWERB013] PRIMARY KEY  CLUSTERED 
	(
		[ID013]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[DISZ019] WITH NOCHECK ADD 
	CONSTRAINT [PK_DISZ019] PRIMARY KEY  CLUSTERED 
	(
		[ID019]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[DOC018] WITH NOCHECK ADD 
	CONSTRAINT [PK_DOC018] PRIMARY KEY  CLUSTERED 
	(
		[ID018]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[FORD022] WITH NOCHECK ADD 
	CONSTRAINT [PK_FORD018] PRIMARY KEY  CLUSTERED 
	(
		[ID022]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[FORT011] WITH NOCHECK ADD 
	CONSTRAINT [PK_FORT011] PRIMARY KEY  CLUSTERED 
	(
		[ID011]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[FRIST024] WITH NOCHECK ADD 
	CONSTRAINT [PK_FRIST024] PRIMARY KEY  CLUSTERED 
	(
		[ID024]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[LG003] WITH NOCHECK ADD 
	CONSTRAINT [PK_LG003] PRIMARY KEY  CLUSTERED 
	(
		[ID003]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[OLG002] WITH NOCHECK ADD 
	CONSTRAINT [PK_OLG002] PRIMARY KEY  CLUSTERED 
	(
		[ID002]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[RA010] WITH NOCHECK ADD 
	CONSTRAINT [PK_RA010] PRIMARY KEY  CLUSTERED 
	(
		[ID010]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[STELLEN012] WITH NOCHECK ADD 
	CONSTRAINT [PK_STELLEN012] PRIMARY KEY  CLUSTERED 
	(
		[ID012]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[VALUES015] WITH NOCHECK ADD 
	CONSTRAINT [PK_VALUES015] PRIMARY KEY  CLUSTERED 
	(
		[ID015]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[AFORT014] ADD 
	CONSTRAINT [DF_AFORT014_ID014] DEFAULT (newid()) FOR [ID014],
	CONSTRAINT [IX_AFORT014] UNIQUE  NONCLUSTERED 
	(
		[FK010014],
		[FK011014]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] ,
	CONSTRAINT [IX_AFORT014_1] UNIQUE  NONCLUSTERED 
	(
		[FK010014],
		[FK011014]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[AG004] ADD 
	CONSTRAINT [DF_AG004_ID004] DEFAULT (newid()) FOR [ID004],
	CONSTRAINT [DF_AG004_MODIFY004] DEFAULT (getdate()) FOR [MODIFY004],
	CONSTRAINT [DF_AG004_CREATE004] DEFAULT (getdate()) FOR [CREATE004],
	CONSTRAINT [IX_AG004] UNIQUE  NONCLUSTERED 
	(
		[AGNAME004]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[AKTENORT017] ADD 
	CONSTRAINT [DF_AKTENORT017_ID017] DEFAULT (newid()) FOR [ID017],
	CONSTRAINT [DF_AKTENORT017_CREATE017] DEFAULT (getdate()) FOR [CREATE017],
	CONSTRAINT [DF_AKTENORT017_MODIFY017] DEFAULT (getdate()) FOR [MODIFY017]
GO

ALTER TABLE [dbo].[AUSSCHREIBUNG020] ADD 
	CONSTRAINT [DF_AUSSCHREIBUNG020_ID020] DEFAULT (newid()) FOR [ID020],
	CONSTRAINT [DF_AUSSCHREIBUNG020_WORKFLOW020] DEFAULT (0) FOR [WORKFLOW020],
	CONSTRAINT [DF_AUSSCHREIBUNG020_CREATE020] DEFAULT (getdate()) FOR [CREATE020],
	CONSTRAINT [DF_AUSSCHREIBUNG020_MODIFY020] DEFAULT (getdate()) FOR [MODIFY020]
GO

 CREATE  UNIQUE  INDEX [IX_AUSSCHREIBUNG020] ON [dbo].[AUSSCHREIBUNG020]([JAHR020]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

ALTER TABLE [dbo].[BERECHNUNGEN016] ADD 
	CONSTRAINT [DF_Berechnungen_ID016] DEFAULT (newid()) FOR [ID016],
	CONSTRAINT [DF_Berechnungen_FAKTOR016] DEFAULT (0.0) FOR [FAKTOR016],
	CONSTRAINT [DF_BERECHNUNGEN016_MAXWERT016] DEFAULT (0) FOR [MAXWERT016],
	CONSTRAINT [DF_BERECHNUNGEN016_MAXVALUE016] DEFAULT (0) FOR [MAXVALUE016],
	CONSTRAINT [DF_BERECHNUNGEN016_LOCKED016] DEFAULT (0) FOR [LOCKED016]
GO

ALTER TABLE [dbo].[BEWERB013] ADD 
	CONSTRAINT [DF_BEWERB013_ID013] DEFAULT (newid()) FOR [ID013],
	CONSTRAINT [DF_BEWERB013_EINGANG013] DEFAULT (getdate()) FOR [EINGANG013],
	CONSTRAINT [DF_BEWERB013_PUNKTE01013] DEFAULT (0) FOR [PUNKTE01013],
	CONSTRAINT [DF_BEWERB013_PUNKTE02013] DEFAULT (0) FOR [PUNKTE02013],
	CONSTRAINT [DF_BEWERB013_PUNKTE03013] DEFAULT (0) FOR [PUNKTE03013],
	CONSTRAINT [DF_BEWERB013_PUNKTE04013] DEFAULT (0) FOR [PUNKTE04013],
	CONSTRAINT [DF_BEWERB013_PUNKTE05013] DEFAULT (0) FOR [PUNKTE05013],
	CONSTRAINT [DF_BEWERB013_PUNKTE06013] DEFAULT (0) FOR [PUNKTE06013],
	CONSTRAINT [DF_BEWERB013_PUNKTE07013] DEFAULT (0) FOR [PUNKTE07013],
	CONSTRAINT [DF_BEWERB013_PUNKTE08013] DEFAULT (0) FOR [PUNKTE08013],
	CONSTRAINT [DF_BEWERB013_PUNKTE09013] DEFAULT (0) FOR [PUNKTE09013],
	CONSTRAINT [DF_BEWERB013_PUNKTE10013] DEFAULT (0) FOR [PUNKTE10013],
	CONSTRAINT [DF_BEWERB013_PUNKTE11013] DEFAULT (0) FOR [PUNKTE11013],
	CONSTRAINT [DF_BEWERB013_PUNKTE12013] DEFAULT (0) FOR [PUNKTE12013],
	CONSTRAINT [DF_BEWERB013_PUNKTE13013] DEFAULT (0) FOR [PUNKTE13013],
	CONSTRAINT [DF_BEWERB013_PUNKTE14013] DEFAULT (0) FOR [PUNKTE14013],
	CONSTRAINT [DF_BEWERB013_PUNKZE15013] DEFAULT (0) FOR [PUNKTE15013],
	CONSTRAINT [DF_BEWERB013_PUNKTE16013] DEFAULT (0) FOR [PUNKTE16013],
	CONSTRAINT [DF_BEWERB013_PUNKTE17013] DEFAULT (0) FOR [PUNKTE17013],
	CONSTRAINT [DF_BEWERB013_PUNKTE18013] DEFAULT (0) FOR [PUNKTE18013],
	CONSTRAINT [DF_BEWERB013_PUNKTE19013] DEFAULT (0) FOR [PUNKTE19013],
	CONSTRAINT [DF_BEWERB013_PUNKTE20013] DEFAULT (0) FOR [PUNKTE20013],
	CONSTRAINT [DF_BEWERB013_PUNKTE21013] DEFAULT (0) FOR [PUNKTE21013],
	CONSTRAINT [DF_BEWERB013_RANG013] DEFAULT (0) FOR [RANG013],
	CONSTRAINT [DF_BEWERB013_CREATE013] DEFAULT (getdate()) FOR [CREATE013],
	CONSTRAINT [DF_BEWERB013_MODIFY013] DEFAULT (getdate()) FOR [MODIFY013]
GO

ALTER TABLE [dbo].[DISZ019] ADD 
	CONSTRAINT [DF_DISZ019_ID019] DEFAULT (newid()) FOR [ID019],
	CONSTRAINT [DF_DISZ019_CREATE019] DEFAULT (getdate()) FOR [CREATE019],
	CONSTRAINT [DF_DISZ019_MODIFY019] DEFAULT (getdate()) FOR [MODIFY019]
GO

ALTER TABLE [dbo].[DOC018] ADD 
	CONSTRAINT [DF_DOC018_ID018] DEFAULT (newid()) FOR [ID018],
	CONSTRAINT [DF_DOC018_CREATE018] DEFAULT (getdate()) FOR [CREATE018],
	CONSTRAINT [DF_DOC018_MODIFY018] DEFAULT (getdate()) FOR [MODIFY018]
GO

ALTER TABLE [dbo].[FORD022] ADD 
	CONSTRAINT [DF_FORD018_ID018] DEFAULT (newid()) FOR [ID022],
	CONSTRAINT [DF_FORD022_ORDER022] DEFAULT (0) FOR [ORDER022],
	CONSTRAINT [DF_FORD022_BETRAG022] DEFAULT (0) FOR [BETRAG022],
	CONSTRAINT [DF_FORD018_CREATE018] DEFAULT (getdate()) FOR [CREATE022],
	CONSTRAINT [DF_FORD018_MODIFY018] DEFAULT (getdate()) FOR [MODIFY022]
GO

ALTER TABLE [dbo].[FORT011] ADD 
	CONSTRAINT [DF_FORT011_ID011] DEFAULT (newid()) FOR [ID011],
	CONSTRAINT [DF_FORT011_CREATE011] DEFAULT (getdate()) FOR [CREATE011],
	CONSTRAINT [DF_FORT011_MODIFY011] DEFAULT (getdate()) FOR [MODIFY011]
GO

ALTER TABLE [dbo].[FRIST024] ADD 
	CONSTRAINT [DF_FRIST024_ID024] DEFAULT (newid()) FOR [ID024],
	CONSTRAINT [DF_FRIST024_CREATE024] DEFAULT (getdate()) FOR [CREATE024],
	CONSTRAINT [DF_FRIST024_MODIFY024] DEFAULT (getdate()) FOR [MODIFY024]
GO

ALTER TABLE [dbo].[LAND005] ADD 
	CONSTRAINT [DF_LAND005_ID005] DEFAULT (newid()) FOR [ID005]
GO

ALTER TABLE [dbo].[LG003] ADD 
	CONSTRAINT [DF_LG003_ID003] DEFAULT (newid()) FOR [ID003],
	CONSTRAINT [DF_LG003_MODIFY003] DEFAULT (getdate()) FOR [MODIFY003],
	CONSTRAINT [DF_LG003_CREATE003] DEFAULT (getdate()) FOR [CREATE003],
	CONSTRAINT [IX_LG003] UNIQUE  NONCLUSTERED 
	(
		[LGNAME003]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

 CREATE  INDEX [IX_LG003_1] ON [dbo].[LG003]([ID003], [LGNAME003]) WITH  FILLFACTOR = 90 ON [PRIMARY]
GO

ALTER TABLE [dbo].[OLG002] ADD 
	CONSTRAINT [DF_OLG002_ID002] DEFAULT (newid()) FOR [ID002],
	CONSTRAINT [DF_OLG002_MODIFY002] DEFAULT (getdate()) FOR [MODIFY002],
	CONSTRAINT [DF_OLG002_CREATE002] DEFAULT (getdate()) FOR [CREATE002],
	CONSTRAINT [IX_OLG002] UNIQUE  NONCLUSTERED 
	(
		[OLGNAME002]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[OPTIONS023] ADD 
	CONSTRAINT [DF_OPTIONS023_ID23] DEFAULT (newid()) FOR [ID23]
GO

ALTER TABLE [dbo].[RA010] ADD 
	CONSTRAINT [DF_RA010_ID010] DEFAULT (newid()) FOR [ID010],
	CONSTRAINT [DF_RA010_VORNAME010] DEFAULT ('') FOR [VORNAME010],
	CONSTRAINT [DF_RA010_VERSTORBEN010] DEFAULT (0) FOR [VERSTORBEN010],
	CONSTRAINT [DF_RA010_STATUS010] DEFAULT ('Bewerber') FOR [STATUS010],
	CONSTRAINT [DF_RA010_CREATE010] DEFAULT (getdate()) FOR [CREATE010],
	CONSTRAINT [DF_RA010_MODIFY010] DEFAULT (getdate()) FOR [MODIFY010]
GO

ALTER TABLE [dbo].[STELLEN012] ADD 
	CONSTRAINT [DF_STELLEN012_ID012] DEFAULT (newid()) FOR [ID012],
	CONSTRAINT [DF_STELLEN012_ANZ012] DEFAULT (1) FOR [ANZ012],
	CONSTRAINT [DF_STELLEN012_CREATE012] DEFAULT (getdate()) FOR [CREATE012],
	CONSTRAINT [DF_STELLEN012_MODIFY012] DEFAULT (getdate()) FOR [MODIFY012]
GO

ALTER TABLE [dbo].[USER001] ADD 
	CONSTRAINT [DF_USER001_ID001] DEFAULT (newid()) FOR [ID001],
	CONSTRAINT [DF_USER001_SYSTEM001] DEFAULT (0) FOR [SYSTEM001],
	CONSTRAINT [DF_USER001_CREATE001] DEFAULT (getdate()) FOR [CREATE001],
	CONSTRAINT [DF_USER001_MODIFY001] DEFAULT (getdate()) FOR [MODIFY001]
GO

ALTER TABLE [dbo].[VALUES015] ADD 
	CONSTRAINT [DF_VALUES015_ID015] DEFAULT (newid()) FOR [ID015]
GO

ALTER TABLE [dbo].[VORGANG021] ADD 
	CONSTRAINT [DF_Vorgang_ID021] DEFAULT (newid()) FOR [ID021],
	CONSTRAINT [DF_Vorgang_CREATE021] DEFAULT (getdate()) FOR [CREATE021],
	CONSTRAINT [DF_Vorgang_MODIFY021] DEFAULT (getdate()) FOR [MODIFY021]
GO

ALTER TABLE [dbo].[WORKFLOW006] ADD 
	CONSTRAINT [DF_WORKFLOW006_ID006_1] DEFAULT (newid()) FOR [ID006],
	CONSTRAINT [DF_WORKFLOW006_CREATE006_1] DEFAULT (getdate()) FOR [CREATE006],
	CONSTRAINT [DF_WORKFLOW006_MODIFY006_1] DEFAULT (getdate()) FOR [MODIFY006]
GO

ALTER TABLE [dbo].[AFORT014] ADD 
	CONSTRAINT [FK_AFORT014_FORT011] FOREIGN KEY 
	(
		[FK011014]
	) REFERENCES [dbo].[FORT011] (
		[ID011]
	),
	CONSTRAINT [FK_AFORT014_RA010] FOREIGN KEY 
	(
		[FK010014]
	) REFERENCES [dbo].[RA010] (
		[ID010]
	)
GO

ALTER TABLE [dbo].[AG004] ADD 
	CONSTRAINT [FK_AG004_LG003] FOREIGN KEY 
	(
		[FK003004]
	) REFERENCES [dbo].[LG003] (
		[ID003]
	)
GO

ALTER TABLE [dbo].[AKTENORT017] ADD 
	CONSTRAINT [FK_AKTENORT017_RA010] FOREIGN KEY 
	(
		[FK010017]
	) REFERENCES [dbo].[RA010] (
		[ID010]
	)
GO

ALTER TABLE [dbo].[BEWERB013] ADD 
	CONSTRAINT [FK_BEWERB013_RA010] FOREIGN KEY 
	(
		[FK010013]
	) REFERENCES [dbo].[RA010] (
		[ID010]
	),
	CONSTRAINT [FK_BEWERB013_STELLEN012] FOREIGN KEY 
	(
		[FK012013]
	) REFERENCES [dbo].[STELLEN012] (
		[ID012]
	)
GO

ALTER TABLE [dbo].[DOC018] ADD 
	CONSTRAINT [FK_DOC018_RA010] FOREIGN KEY 
	(
		[FK010018]
	) REFERENCES [dbo].[RA010] (
		[ID010]
	)
GO

ALTER TABLE [dbo].[LG003] ADD 
	CONSTRAINT [FK_LG003_OLG002] FOREIGN KEY 
	(
		[FK002003]
	) REFERENCES [dbo].[OLG002] (
		[ID002]
	)
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MitbewerberDaten]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[MitbewerberDaten]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Empf‰ngerDaten]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[Empf‰ngerDaten]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[StellenDaten]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[StellenDaten]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[BewerberDaten]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[BewerberDaten]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MitbewerberAbgesagtDaten]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[MitbewerberAbgesagtDaten]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MitbewerberZugesagtDaten]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[MitbewerberZugesagtDaten]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PunkteBewerber]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[PunkteBewerber]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AbsenderDaten]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[AbsenderDaten]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AnzahlBewerber]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[AnzahlBewerber]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AusschreibungsDaten]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[AusschreibungsDaten]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Ausgeschriebene Stellen]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[Ausgeschriebene Stellen]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[AusgeschriebeneStellen]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[AusgeschriebeneStellen]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE VIEW dbo.[Ausgeschriebene Stellen]
AS
SELECT DISTINCT YEAR(FRIST012) AS Jahr, BEZIRK012 AS Bezirk, ANZ012 AS [Ausgeschriebene Stellen]
FROM         dbo.STELLEN012


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.AusgeschriebeneStellen
AS
SELECT DISTINCT YEAR(FRIST012) AS Jahr, BEZIRK012 AS Bezirk, ANZ012 AS 'Ausgeschriebene Stellen'
FROM         dbo.STELLEN012

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



CREATE VIEW dbo.AbsenderDaten
AS
SELECT     ISNULL(TEL001, '') AS AbsenderTel, ISNULL(FAX001, '') AS AbsenderFax, ISNULL(EMAIL001, '') AS AbsenderEmail, ID001, ISNULL(VORNAME001, '') 
                      + ', ' + ISNULL(NACHNAME001, '') AS Absender
FROM         dbo.USER001



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



CREATE VIEW dbo.AnzahlBewerber
AS
SELECT     dbo.RA010.ID010, COUNT(*) AS AnzBewerber
FROM         dbo.RA010 INNER JOIN
                      dbo.BEWERB013 ON dbo.RA010.ID010 = dbo.BEWERB013.FK010013 AND dbo.RA010.STATUS010 = 'Bewerber' LEFT OUTER JOIN
                      dbo.STELLEN012 ON dbo.STELLEN012.ID012 = dbo.BEWERB013.FK012013 LEFT OUTER JOIN
                      dbo.BEWERB013 Mitbewerber ON Mitbewerber.FK012013 = dbo.STELLEN012.ID012
GROUP BY dbo.RA010.ID010



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE VIEW dbo.AusschreibungsDaten
AS
SELECT     JAHR020 AS AusschreibungJahr, AZ020 AS AusschreibungAZ, ID020 AS AusschreibungID
FROM         dbo.AUSSCHREIBUNG020


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE VIEW [dbo].[BewerberDaten]
AS
SELECT     dbo.RA010.ID010 AS PersID, dbo.BEWERB013.EINGANG013 AS BewerberBewDatum, ISNULL(dbo.BEWERB013.RANG013, '') AS BewerberRang, 
                      ISNULL(CAST(dbo.BEWERB013.PUNKTE01013 AS Varchar(8)), '-') AS BewerberNoteStaatsEx, 
					  ISNULL(dbo.BEWERB013.PUNKTE03013, 0) AS BewerberPunkteStaatsexamen,  
						ISNULL(CAST(dbo.BEWERB013.PUNKTE02013 AS Varchar(8)), '-') AS BewerberNoteFachpr¸fung, 
					  ISNULL(dbo.BEWERB013.PUNKTE04013, 0) AS BewerberPunkteFachpr¸fung,  
                      ISNULL(dbo.BEWERB013.PUNKTESUM013, 0) AS BewerberPunkte,  
                      dbo.RA010.ID010 AS BewerberPersID, 
                      dbo.BEWERB013.FK012013 AS StellenID 

FROM         dbo.RA010 INNER JOIN
                      dbo.BEWERB013 ON dbo.RA010.ID010 = dbo.BEWERB013.FK010013



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW [dbo].[MitbewerberAbgesagtDaten]
AS
SELECT     TOP 100 PERCENT dbo.RA010.ID010 AS PersID, CASE WHEN ISNULL(dbo.RA010.TITEL010, '') = '' THEN ISNULL(dbo.RA010.NACHNAME010, '') 
                      + ', ' + ISNULL(dbo.RA010.VORNAME010, '') ELSE ISNULL(dbo.RA010.TITEL010, '') + ' ' + ISNULL(dbo.RA010.NACHNAME010, '') 
                      + ', ' + ISNULL(dbo.RA010.VORNAME010, '') END AS AMitbewerberVoll, CASE WHEN ISNULL(dbo.RA010.TITEL010, '') 
                      = '' THEN dbo.RA010.NACHNAME010 ELSE ISNULL(dbo.RA010.TITEL010, '') + ' ' + dbo.RA010.NACHNAME010 END AS AMitbewerberNachname, 
                      ISNULL(dbo.RA010.ANREDE010, '') AS AMitbewerberAnrede, ISNULL(dbo.RA010.TITEL010, '') AS AMitbewerberTitel, ISNULL(dbo.RA010.NAMEZUS010, 
                      '') AS AMitbewerberNamenszusatz, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Sehr geehrte Frau' ELSE 'Sehr geehrter Herr' END AS AMitbewerberAnredeVoll, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Frau Rechtsanw‰ltin' ELSE 'Herrn Rechtsanwalt' END AS AMitbewerberAnredeFormal, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Sehr geerhrte Frau ' ELSE 'Sehr geehrter Herr ' END + CASE WHEN ISNULL(TITEL010, '') 
                      = '' THEN dbo.RA010.NACHNAME010 ELSE ISNULL(TITEL010, '') + ' ' + dbo.RA010.NACHNAME010 END AS AMitbewerberAnredeVollName, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Frau Rechtsanw‰ltin ' ELSE 'Herr Rechtsanwalt ' END + CASE WHEN ISNULL(TITEL010, '') 
                      = '' THEN ISNULL(dbo.RA010.VORNAME010, '') + ' ' + dbo.RA010.NACHNAME010 ELSE ISNULL(TITEL010, '') + ' ' + ISNULL(dbo.RA010.VORNAME010, '') 
                      + ' ' + dbo.RA010.NACHNAME010 END AS AMitbewerberAnredeFormalName, ISNULL(dbo.RA010.AMTORT010, '-') AS AMitbewerberAmtssitz, 
                      dbo.BEWERB013.EINGANG013 AS AMitbewerberBewDatum, dbo.RA010.AZ010 AS AMitbewerberAZVI, ISNULL(dbo.BEWERB013.RANG013, '') 
                      AS AMitbewerberRang, dbo.RA010.AG010 AS AMitbewerberAGzugelassen, dbo.RA010.LG010 AS AMitbewerberLGZugelassen, 
                      ISNULL(dbo.RA010.VORNAME010, '') AS AMitbewerberVorname, ISNULL(CAST(dbo.BEWERB013.PUNKTE01013 AS Varchar(8)), '-') 
                      AS AMitbewerberNoteStaatsEx, ISNULL(dbo.BEWERB013.PUNKTE03013, 0) AS AMitbewerberPunkteStaatsexamen, 
                       CAST(CONVERT(money,ISNULL(dbo.BEWERB013.PUNKTESUM013, 0)) AS varchar(30)) AS AMitbewerberPunkte,  
                      dbo.RA010.ID010 AS AMitbewerberPersID, dbo.BEWERB013.FK012013 AS StellenID, 
						CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Rechtsanw«œltin ' ELSE 'Rechtsanwalt ' END AS AMitbewerberAmtsbez,
                       CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'RA in ' ELSE 'RA ' END AS AMitbewerberAmtsbezKurz, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Notarin' ELSE 'Notar ' END AS AMitbewerberNotar
FROM         dbo.RA010 INNER JOIN
                      dbo.BEWERB013 ON dbo.RA010.ID010 = dbo.BEWERB013.FK010013
ORDER BY ISNULL(dbo.BEWERB013.RANG013, '')


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW [dbo].[MitbewerberZugesagtDaten]
AS
SELECT     TOP 100 PERCENT dbo.RA010.ID010 AS PersID, CASE WHEN ISNULL(dbo.RA010.TITEL010, '') = '' THEN ISNULL(dbo.RA010.NACHNAME010, '') 
                      + ', ' + ISNULL(dbo.RA010.VORNAME010, '') ELSE ISNULL(dbo.RA010.TITEL010, '') + ' ' + ISNULL(dbo.RA010.NACHNAME010, '') 
                      + ', ' + ISNULL(dbo.RA010.VORNAME010, '') END AS ZMitbewerberVoll, CASE WHEN ISNULL(dbo.RA010.TITEL010, '') 
                      = '' THEN dbo.RA010.NACHNAME010 ELSE ISNULL(dbo.RA010.TITEL010, '') + ' ' + dbo.RA010.NACHNAME010 END AS ZMitbewerberNachname, 
                      ISNULL(dbo.RA010.ANREDE010, '') AS ZMitbewerberAnrede, ISNULL(dbo.RA010.TITEL010, '') AS ZMitbewerberTitel, ISNULL(dbo.RA010.NAMEZUS010, 
                      '') AS ZMitbewerberNamenszusatz, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Sehr geehrte Frau' ELSE 'Sehr geehrter Herr' END AS ZMitbewerberAnredeVoll, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Frau Rechtsanw‰ltin' ELSE 'Herrn Rechtsanwalt' END AS ZMitbewerberAnredeFormal, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Sehr geerhrte Frau ' ELSE 'Sehr geehrter Herr ' END + CASE WHEN ISNULL(dbo.RA010.TITEL010, '') 
                      = '' THEN dbo.RA010.NACHNAME010 ELSE ISNULL(dbo.RA010.TITEL010, '') + ' ' + dbo.RA010.NACHNAME010 END AS ZMitbewerberAnredeVollName, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Frau Rechtsanw‰ltin ' ELSE 'Herr Rechtsanwalt ' END + CASE WHEN ISNULL(dbo.RA010.TITEL010, '') 
                      = '' THEN ISNULL(dbo.RA010.VORNAME010, '') + ' ' + dbo.RA010.NACHNAME010 ELSE ISNULL(dbo.RA010.TITEL010, '') 
                      + ' ' + ISNULL(dbo.RA010.VORNAME010, '') + ' ' + dbo.RA010.NACHNAME010 END AS ZMitbewerberAnredeFormalName, 
                      ISNULL(dbo.RA010.AMTORT010, '-') AS ZMitbewerberAmtssitz, dbo.BEWERB013.EINGANG013 AS ZMitbewerberBewDatum, 
                      dbo.RA010.AZ010 AS ZMitbewerberAZVI, ISNULL(dbo.BEWERB013.RANG013, '') AS ZMitbewerberRang, 
                      dbo.RA010.AG010 AS ZMitbewerberAGzugelassen, dbo.RA010.LG010 AS ZMitbewerberLGZugelassen, ISNULL(dbo.RA010.VORNAME010, '') 
                      AS ZMitbewerberVorname, ISNULL(CAST(dbo.BEWERB013.PUNKTE01013 AS Varchar(8)), '-') AS ZMitbewerberNoteStaatsEx, 
						ISNULL(dbo.BEWERB013.PUNKTE03013, 0)  AS ZMitbewerberPunkteStaatsexamen, 
					 CAST(CONVERT(money, ISNULL(dbo.BEWERB013.PUNKTESUM013, 0)) AS varchar(30)) AS ZMitbewerberPunkte, 
                      dbo.RA010.ID010 AS ZMitbewerberPersID,  dbo.BEWERB013.FK012013 AS StellenID, 
                      ISNULL(CAST(dbo.BEWERB013.PUNKTE02013 AS Varchar(8)), '-') AS ZMitbewerberNoteFachpr¸fung, 
						ISNULL(dbo.BEWERB013.PUNKTE04013, 0)  AS ZMitbewerberPunkteFachpr¸fung,
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Rechtsanw«œltin ' ELSE 'Rechtsanwalt ' END AS ZMitbewerberAmtsbez, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'RA in ' ELSE 'RA ' END AS ZMitbewerberAmtsbezKurz, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Notarin' ELSE 'Notar ' END AS ZMitbewerberKurz
FROM         dbo.RA010 INNER JOIN
                      dbo.BEWERB013 ON dbo.RA010.ID010 = dbo.BEWERB013.FK010013
ORDER BY ISNULL(dbo.BEWERB013.RANG013, '')


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



CREATE VIEW [dbo].[PunkteBewerber]
AS
SELECT     ISNULL(CAST(dbo.BEWERB013.PUNKTE01013 AS Varchar(8)), '-') AS NoteStaatsEx, ISNULL(dbo.BEWERB013.PUNKTE03013, 0) AS PunkteStaatsexamen, 
                      ISNULL(CAST(dbo.BEWERB013.PUNKTE01013 AS Varchar(8)), '-') AS NoteFachpr¸fung, 
					  ISNULL(dbo.BEWERB013.PUNKTE03013, 0) AS PunkteFachpr¸fung,
                      ISNULL(dbo.BEWERB013.PUNKTESUM013, 0) AS Punkte, dbo.STELLEN012.FRIST012 AS Bewebungsfrist, 
					  dbo.RA010.ID010 AS PersID 
                      
FROM         dbo.RA010 LEFT OUTER JOIN
                      dbo.BEWERB013 ON dbo.RA010.ID010 = dbo.BEWERB013.FK010013 LEFT OUTER JOIN
                      dbo.STELLEN012 ON dbo.STELLEN012.ID012 = dbo.BEWERB013.FK012013




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.Empf‰ngerDaten
AS
SELECT     dbo.RA010.ID010 AS PersID, CASE WHEN ISNULL(TITEL010, '') = '' THEN '' + NACHNAME010 ELSE ISNULL(TITEL010, '') 
                      + ' ' + NACHNAME010 END AS RANachname, ISNULL(dbo.RA010.VORNAME010, '') AS RAVorname, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Sehr geehrte Frau' ELSE 'Sehr geehrter Herr' END AS RAAnredeVoll, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Frau Rechtsanw‰ltin' ELSE 'Herrn Rechtsanwalt' END AS RAAnredeFormal, 
                      dbo.RA010.KSTR010 + ' ' + dbo.RA010.KPLZ010 + ' ' + dbo.RA010.KORT010 AS RAAnschriftKanzlei, ISNULL(dbo.RA010.KSTR010, '') AS RAStrKanzlei, 
                      dbo.RA010.KPLZ010 + ' ' + dbo.RA010.KORT010 AS RAPLZOrtKanzlei, ISNULL(dbo.RA010.AZ010, '-') AS RAAZVI, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Sehr geehrte Frau ' ELSE 'Sehr geehrter Herr ' END + CASE WHEN ISNULL(TITEL010, '') 
                      = '' THEN '' + NACHNAME010 ELSE ISNULL(TITEL010, '') + ' ' + NACHNAME010 END AS RAAnredeVollName, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Frau Rechtsanw‰ltin ' ELSE 'Herr Rechtsanwalt ' END + CASE WHEN ISNULL(TITEL010, '') 
                      = '' THEN ISNULL(VORNAME010, '') + ' ' + NACHNAME010 ELSE ISNULL(TITEL010, '') + ' ' + ISNULL(VORNAME010, '') 
                      + ' ' + NACHNAME010 END AS RAAnredeFormalName, CASE WHEN ISNULL(TITEL010, '') = '' THEN ISNULL(VORNAME010, '') 
                      + ' ' + NACHNAME010 ELSE ISNULL(TITEL010, '') + ' ' + ISNULL(VORNAME010, '') + ' ' + NACHNAME010 END AS RAVornameNachname, 
                      CASE WHEN ISNULL(TITEL010, '') = '' THEN '' + NACHNAME010 ELSE ISNULL(TITEL010, '') 
                      + ' ' + NACHNAME010 END + ', ' + ISNULL(dbo.RA010.VORNAME010, '') AS RANachnameVorname, 
                      CASE WHEN RA010.ANREDE010 = 'Fau' THEN 'Frau' ELSE 'Herr' END AS RAAnrede, dbo.RA010.GEB010 AS RAGebDat, 
                      ISNULL(dbo.RA010.AMTORT010, '-') AS RAAmtssitz, dbo.RA010.ANWALTSEIT010 AS RAAnwaltSeit, dbo.RA010.AG010 AS RAAGzugelassen, 
                      dbo.RA010.LG010 AS RALGZugelassen, CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Rechtsanw‰ltin ' ELSE 'Rechtsanwalt ' END AS RAAmtsbez, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'RA in ' ELSE 'RA ' END AS RAAmtsbezKurz, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Notarin' ELSE 'Notar ' END AS RANotar, ISNULL(CAST(dbo.RA010.EXANOTE010 AS Varchar(8)), '-') 
                      AS RANoteStaatsexamen, ISNULL(dbo.RA010.EXANOTE010, 0) * 5 AS RAPunkteStaasexamen, ISNULL(dbo.LG003.LGPLZ003, '-') 
                      + ' ' + ISNULL(dbo.LG003.LGORT003, '-') AS RALGPLZORT, ISNULL(dbo.AG004.AGPLZ004, '-') + ' ' + ISNULL(dbo.AG004.AGORT004, '-') 
                      AS RAAGPLZORT, ISNULL(dbo.LG003.LGSTR003, '-') AS RALGSTR, ISNULL(dbo.AG004.AGSTR004, '-') AS RAAGSTR
FROM         dbo.RA010 LEFT OUTER JOIN
                      dbo.AG004 ON dbo.RA010.AG010 = dbo.AG004.AGNAME004 LEFT OUTER JOIN
                      dbo.LG003 ON dbo.RA010.LG010 = dbo.LG003.LGNAME003

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE VIEW dbo.StellenDaten
AS
SELECT     dbo.STELLEN012.ID012 AS StellenID, dbo.STELLEN012.FRIST012 AS StellenBewebungsfrist, dbo.STELLEN012.BEZIRK012 AS StellenAGBezirk, 
                      dbo.LG003.LGNAME003 AS StellenLGBezirk, ISNULL(dbo.LG003.LGPLZ003, '') + ' ' + ISNULL(dbo.LG003.LGORT003, '') AS StellenLGPLZOrt, 
                      dbo.STELLEN012.ANZ012 AS StellenAnzStellen, COUNT('dbo.BEWERB013.*') AS StellenAnzBewerber, ISNULL(dbo.LG003.LGSTR003, '') 
                      AS StellenLGStrasse, dbo.STELLEN012.FK020012, dbo.AUSSCHREIBUNG020.AZ020 AS StellenAZ, 
                      dbo.AUSSCHREIBUNG020.JAHR020 AS StellenBewFristJahr
FROM         dbo.AUSSCHREIBUNG020 INNER JOIN
                      dbo.STELLEN012 ON dbo.AUSSCHREIBUNG020.ID020 = dbo.STELLEN012.FK020012 LEFT OUTER JOIN
                      dbo.BEWERB013 ON dbo.STELLEN012.ID012 = dbo.BEWERB013.FK012013 LEFT OUTER JOIN
                      dbo.LG003 RIGHT OUTER JOIN
                      dbo.AG004 ON dbo.LG003.ID003 = dbo.AG004.FK003004 ON dbo.STELLEN012.BEZIRK012 = dbo.AG004.AGNAME004
GROUP BY dbo.STELLEN012.ID012, dbo.STELLEN012.FRIST012, dbo.STELLEN012.BEZIRK012, dbo.LG003.LGNAME003, ISNULL(dbo.LG003.LGPLZ003, '') 
                      + ' ' + ISNULL(dbo.LG003.LGORT003, ''), dbo.STELLEN012.ANZ012, ISNULL(dbo.LG003.LGSTR003, ''), dbo.STELLEN012.FK020012, 
                      dbo.AUSSCHREIBUNG020.AZ020, dbo.AUSSCHREIBUNG020.JAHR020


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW [dbo].[MitbewerberDaten]
AS
SELECT     TOP 100 PERCENT dbo.Empf‰ngerDaten.PersID, dbo.Empf‰ngerDaten.PersID AS MitbewerberPersID, dbo.BEWERB013.ID013, 
                      dbo.Empf‰ngerDaten.RAVornameNachname AS MitbewerberVoll, dbo.Empf‰ngerDaten.RANachname AS MitbewerberNachname, 
                      dbo.Empf‰ngerDaten.RAVorname AS MitbewerberVorname, dbo.Empf‰ngerDaten.RAAnrede AS MitbewerberAnrede, 
                      dbo.Empf‰ngerDaten.RAAnredeVoll AS MitbewerberAnredeVoll, dbo.Empf‰ngerDaten.RAAnredeFormal AS MitbewerberAnredeFormal, 
                      dbo.Empf‰ngerDaten.RAAnredeVollName AS MitbewerberAnredeVollName, 
                      dbo.Empf‰ngerDaten.RAAnredeFormalName AS MitbewerberAnredeFormalName, dbo.Empf‰ngerDaten.RAAmtssitz AS MitbewerberAmtssitz, 
                      dbo.BEWERB013.EINGANG013 AS MitbewerberBewDatum, dbo.Empf‰ngerDaten.RAAZVI AS MitbewerberAZVI, ISNULL(dbo.BEWERB013.RANG013, '') 
                      AS MitbewerberRang, dbo.Empf‰ngerDaten.RAAGzugelassen AS MitbewerberAGzugelassen, 
                      dbo.Empf‰ngerDaten.RALGZugelassen AS MitbewerberLGZugelassen,  
                       ISNULL(dbo.BEWERB013.PUNKTESUM013, 0) AS MitbewerberPunkte, 
                      dbo.BEWERB013.FK012013 AS StellenID,  IsNull(BEWERB013.PUNKTE02013,0) AS MitbewerberNoteFachpr¸fung,
						IsNull(BEWERB013.PUNKTE04013,0) AS MitbewerberPunkteFachpr¸fung,
                      dbo.Empf‰ngerDaten.RAAmtsbez AS MitbewerberAmtsbez, dbo.Empf‰ngerDaten.RAAmtsbezKurz AS MitbewerberAmtsbezKurz, 
                      dbo.Empf‰ngerDaten.RANotar AS MitbewerberNotar, dbo.BEWERB013.PUNKTE01013 AS MitbewerberNoteStaatsexamen, 
                      dbo.BEWERB013.PUNKTE03013 AS MitbewerberPunkteStaatsexamen
FROM         dbo.BEWERB013 RIGHT OUTER JOIN
                      dbo.Empf‰ngerDaten ON dbo.BEWERB013.FK010013 = dbo.Empf‰ngerDaten.PersID
ORDER BY ISNULL(dbo.BEWERB013.RANG013, '')



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

/****** Object:  Login [NotareUser]    Script Date: 10/31/2011 11:39:57 ******/
IF  EXISTS (SELECT * FROM sys.server_principals WHERE name = N'NotareUser')
DROP LOGIN [NotareUser]
GO

/* For security reasons the login is created disabled and with a random password. */
/****** Object:  Login [NotareUser]    Script Date: 10/31/2011 11:39:57 ******/
CREATE LOGIN [NotareUser] WITH PASSWORD=N'''∆√¢À?Nn¥mÌo$?èQπ%QÿÃ‘≥ﬁL„y', DEFAULT_DATABASE=[Notare], DEFAULT_LANGUAGE=[Deutsch], CHECK_EXPIRATION=OFF, CHECK_POLICY=OFF
GO

ALTER LOGIN [NotareUser] DISABLE
GO

CREATE USER [NotareUser] FOR LOGIN [NotareUser] WITH DEFAULT_SCHEMA=[dbo]
GO
