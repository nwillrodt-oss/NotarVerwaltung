
-- Tab Ausschreibung020 anlegen
if exists (select * from dbo.sysobjects where id = object_id(N'[AUSSCHREIBUNG020]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [AUSSCHREIBUNG020]
GO

CREATE TABLE [AUSSCHREIBUNG020] (
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

ALTER TABLE [AUSSCHREIBUNG020] WITH NOCHECK ADD 
	CONSTRAINT [PK_AUSSCHREIBUNG020] PRIMARY KEY  CLUSTERED 
	(
		[ID020]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [AUSSCHREIBUNG020] ADD 
	CONSTRAINT [DF_AUSSCHREIBUNG020_ID020] DEFAULT (newid()) FOR [ID020],
	CONSTRAINT [DF_AUSSCHREIBUNG020_WORKFLOW020] DEFAULT (0) FOR [WORKFLOW020],
	CONSTRAINT [DF_AUSSCHREIBUNG020_CREATE020] DEFAULT (getdate()) FOR [CREATE020],
	CONSTRAINT [DF_AUSSCHREIBUNG020_MODIFY020] DEFAULT (getdate()) FOR [MODIFY020]
GO

 CREATE  UNIQUE  INDEX [IX_AUSSCHREIBUNG020] ON [AUSSCHREIBUNG020]([JAHR020]) ON [PRIMARY]
GO



-- Tab Berechnungen016 löschen
if exists (select * from dbo.sysobjects where id = object_id(N'[BERECHNUNGEN016]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [BERECHNUNGEN016]
GO

-- Tab Berechnungen016 Neuanlegen
CREATE TABLE [BERECHNUNGEN016] (
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

ALTER TABLE [BERECHNUNGEN016] WITH NOCHECK ADD 
	CONSTRAINT [PK_Berechnungen] PRIMARY KEY  CLUSTERED 
	(
		[ID016]
	) WITH  FILLFACTOR = 90  ON [PRIMARY] 
GO

ALTER TABLE [BERECHNUNGEN016] ADD 
	CONSTRAINT [DF_Berechnungen_ID016] DEFAULT (newid()) FOR [ID016],
	CONSTRAINT [DF_Berechnungen_FAKTOR016] DEFAULT (0.0) FOR [FAKTOR016],
	CONSTRAINT [DF_BERECHNUNGEN016_MAXWERT016] DEFAULT (0) FOR [MAXWERT016],
	CONSTRAINT [DF_BERECHNUNGEN016_MAXVALUE016] DEFAULT (0) FOR [MAXVALUE016],
	CONSTRAINT [DF_BERECHNUNGEN016_LOCKED016] DEFAULT (0) FOR [LOCKED016]
GO

-- Tab Berechnungen016 DS Anlegen
DELETE FROM [BERECHNUNGEN016]
GO
INSERT INTO [BERECHNUNGEN016] ([ID016], [FAKTOR016], [VALUETYPE016], [ORDER016], [MAXWERT016], [CAPTION016], [CAPTIONSQL016], [VALUESQL016], [MAXVALUE016], [LOCKED016], [SAVEFIELD016], [PUNKTESAVEFIELD016])
VALUES ('96ebfc16-ca87-48e5-9cd9-16e821433eb2', 0, 'int', 3, 0, '2. a) Angerechnete Monate Wehrdienst / Elternzeit', NULL, NULL, 0, 0, 'PUNKTE15013', NULL)
INSERT INTO [BERECHNUNGEN016] ([ID016], [FAKTOR016], [VALUETYPE016], [ORDER016], [MAXWERT016], [CAPTION016], [CAPTIONSQL016], [VALUESQL016], [MAXVALUE016], [LOCKED016], [SAVEFIELD016], [PUNKTESAVEFIELD016])
VALUES ('e68becad-5479-4397-b210-34dd9ea20b4c', 0.3, 'float', 15, NULL, '    b) Sonstige Fortbildungen (Halbtage)', NULL, '', NULL, 0, 'PUNKTE14013', 'PUNKTE07013')
INSERT INTO [BERECHNUNGEN016] ([ID016], [FAKTOR016], [VALUETYPE016], [ORDER016], [MAXWERT016], [CAPTION016], [CAPTIONSQL016], [VALUESQL016], [MAXVALUE016], [LOCKED016], [SAVEFIELD016], [PUNKTESAVEFIELD016])
VALUES ('5ef32adf-5af4-46fe-ab85-550a1479b9de', 5, 'float', 1, NULL, '1. Note Staatsexamen', NULL, 'SELECT EXANOTE010 AS WERT FROM RA010', NULL, -1, '', 'PUNKTE04013')
INSERT INTO [BERECHNUNGEN016] ([ID016], [FAKTOR016], [VALUETYPE016], [ORDER016], [MAXWERT016], [CAPTION016], [CAPTIONSQL016], [VALUESQL016], [MAXVALUE016], [LOCKED016], [SAVEFIELD016], [PUNKTESAVEFIELD016])
VALUES ('3062a045-4378-42a4-985d-5c281ffaabd5', 0.05, 'int', 30, 25, '   bb) Sonst. Niederschriften (201 - 500)', NULL, 'SELECT CASE WHEN ((ISNULL(PUNKTE01013,0)-ISNULL(PUNKTE02013,0))-ISNULL(punkte18013,0)) > 300 THEN 300 WHEN ((ISNULL(PUNKTE01013,0)-ISNULL(PUNKTE02013,0))-ISNULL(punkte18013,0)) < 0 THEN 0 ELSE ((ISNULL(PUNKTE01013,0) -ISNULL(PUNKTE02013,0))-ISNULL(punkte18013,0))  END FROM BEWERB013 INNER JOIN RA010 on ID010 = FK010013', 300, -1, 'PUNKTE19013', 'PUNKTE10013')
INSERT INTO [BERECHNUNGEN016] ([ID016], [FAKTOR016], [VALUETYPE016], [ORDER016], [MAXWERT016], [CAPTION016], [CAPTIONSQL016], [VALUESQL016], [MAXVALUE016], [LOCKED016], [SAVEFIELD016], [PUNKTESAVEFIELD016])
VALUES ('2b773851-879b-4ac6-b1e7-7e3bdf91cdd6', 0.01, 'int', 40, NULL, '   dd) Sonst. Niederschriften (ab 1001)', NULL, 'SELECT  CASE WHEN (ISNULL(PUNKTE01013,0)-ISNULL(PUNKTE02013,0) - ISNULL(Punkte18013,0)- ISNULL(punkte19013,0)- ISNULL(punkte20013,0)) <0 Then 0 ELSE (ISNULL(PUNKTE01013,0)-ISNULL(PUNKTE02013,0) - ISNULL(Punkte18013,0)- ISNULL(punkte19013,0)- ISNULL(punkte20013,0)) END FROM BEWERB013 INNER JOIN RA010 on ID010 = FK010013', NULL, -1, NULL, 'PUNKTE12013')
INSERT INTO [BERECHNUNGEN016] ([ID016], [FAKTOR016], [VALUETYPE016], [ORDER016], [MAXWERT016], [CAPTION016], [CAPTIONSQL016], [VALUESQL016], [MAXVALUE016], [LOCKED016], [SAVEFIELD016], [PUNKTESAVEFIELD016])
VALUES ('900295e6-5baa-44b0-a385-8f837bca2e73', 0, 'int', 19, NULL, '4. Niederschriften insgesamt', NULL, '', NULL, 0, 'PUNKTE01013', '')
INSERT INTO [BERECHNUNGEN016] ([ID016], [FAKTOR016], [VALUETYPE016], [ORDER016], [MAXWERT016], [CAPTION016], [CAPTIONSQL016], [VALUESQL016], [MAXVALUE016], [LOCKED016], [SAVEFIELD016], [PUNKTESAVEFIELD016])
VALUES ('06650e3f-aec9-4ef5-892d-9840e2b97353', 0.6, 'float', 10, NULL, '3. a) Fortbildungen (Halbtage) innerhalb der letzen 3 Jahre', NULL, '', NULL, 0, 'PUNKTE13013', 'PUNKTE06013')
INSERT INTO [BERECHNUNGEN016] ([ID016], [FAKTOR016], [VALUETYPE016], [ORDER016], [MAXWERT016], [CAPTION016], [CAPTIONSQL016], [VALUESQL016], [MAXVALUE016], [LOCKED016], [SAVEFIELD016], [PUNKTESAVEFIELD016])
VALUES ('94b57c06-94ce-4838-ae3f-9df97ba3bb89', 0.02, 'int', 35, 20, '   cc) Sonst. Niederschriften (501 - 1000)', NULL, 'SELECT CASE WHEN ((ISNULL(PUNKTE01013,0)-ISNULL(PUNKTE02013,0))-ISNULL(punkte18013,0) - ISNULL(PUNKTE19013,0)) > 500 THEN 500 WHEN ((ISNULL(PUNKTE01013,0)-ISNULL(PUNKTE02013,0))-ISNULL(punkte18013,0) - ISNULL(PUNKTE19013,0)) < 0 THEN 0 ELSE ((ISNULL(PUNKTE01013,0)-ISNULL(PUNKTE02013,0))-ISNULL(punkte18013,0) - ISNULL(PUNKTE19013,0)) END FROM BEWERB013 INNER JOIN RA010 on ID010 = FK010013', 500, -1, 'PUNKTE20013', 'PUNKTE11013')
INSERT INTO [BERECHNUNGEN016] ([ID016], [FAKTOR016], [VALUETYPE016], [ORDER016], [MAXWERT016], [CAPTION016], [CAPTIONSQL016], [VALUESQL016], [MAXVALUE016], [LOCKED016], [SAVEFIELD016], [PUNKTESAVEFIELD016])
VALUES ('44c096ea-a73f-4c1e-958e-c097c6b4e5dd', 0, 'float', 50, NULL, 'Summe', NULL, NULL, NULL, -1, '', 'Punkte')
INSERT INTO [BERECHNUNGEN016] ([ID016], [FAKTOR016], [VALUETYPE016], [ORDER016], [MAXWERT016], [CAPTION016], [CAPTIONSQL016], [VALUESQL016], [MAXVALUE016], [LOCKED016], [SAVEFIELD016], [PUNKTESAVEFIELD016])
VALUES ('4d433973-fdc9-408a-8ce0-d6b9282711fc', 0, 'float', 22, 0, '   b) Sonst. Niederschriften', NULL, 'SELECT CASE WHEN ISNULL(PUNKTE01013,0) - IsNull(PUNKTE02013,0) <0 THEN 0 ELSE ISNULL(PUNKTE01013,0) - IsNull(PUNKTE02013,0) END FROM BEWERB013 INNER JOIN RA010 on ID010 = FK010013', 0, -1, NULL, NULL)


-- Sichten Aktualisieren
if exists (select * from dbo.sysobjects where id = object_id(N'[StellenDaten]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [StellenDaten]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[HTFortbildungen3Jahre]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [HTFortbildungen3Jahre]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[HTSonstFortbildungen]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [HTSonstFortbildungen]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[MitbewerberDaten]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [MitbewerberDaten]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[AbsenderDaten]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [AbsenderDaten]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[AnzahlBewerber]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [AnzahlBewerber]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[AusschreibungsDaten]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [AusschreibungsDaten]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[BewerberDaten]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [BewerberDaten]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[EmpfÃ¤ngerDaten]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [EmpfÃ¤ngerDaten]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[MitbewerberAbgesagtDaten]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [MitbewerberAbgesagtDaten]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[MitbewerberDaten_Alt]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [MitbewerberDaten_Alt]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[MitbewerberZugesagtDaten]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [MitbewerberZugesagtDaten]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[PunkteBewerber]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [PunkteBewerber]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[Ausgeschriebene Stellen]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [Ausgeschriebene Stellen]
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

CREATE VIEW dbo.BewerberDaten
AS
SELECT     dbo.RA010.ID010 AS PersID, dbo.BEWERB013.EINGANG013 AS BewerberBewDatum, ISNULL(dbo.BEWERB013.RANG013, '') AS BewerberRang, 
                      ISNULL(CAST(dbo.RA010.EXANOTE010 AS Varchar(8)), '-') AS BewerberNoteStaatsEx, ISNULL(dbo.RA010.EXANOTE010, 0) * 5 AS BewerberPunkteStaatsexamen, ISNULL(dbo.BEWERB013.PUNKTE05013, 0) AS BewerberPunkteAnwalttätigkeit, 
                      ISNULL(dbo.BEWERB013.PUNKTE06013, 0) AS BewerberPunkteFortbildungen3Jahre, ISNULL(dbo.BEWERB013.PUNKTE07013, 0) 
                      AS BewerberPunkteSonstigeFortbildungen, ISNULL(dbo.BEWERB013.PUNKTE06013, 0) + ISNULL(dbo.BEWERB013.PUNKTE07013, 0) 
                      AS BewerberPunkteFortbildungenGesamt, ISNULL(dbo.BEWERB013.PUNKTE08013, 0) AS BewerberPunkteNS3Jahren, 
                      ISNULL(dbo.BEWERB013.PUNKTE09013, 0) AS BewerberPunkteNS200, ISNULL(dbo.BEWERB013.PUNKTE10013, 0) AS BewerberPunkteNS500, 
                      ISNULL(dbo.BEWERB013.PUNKTE11013, 0) AS BewerberPunkteNS1000, ISNULL(dbo.BEWERB013.PUNKTE12013, 0) AS BewerberPunkteNSab1000, 
                      ISNULL(dbo.BEWERB013.PUNKTE03013, 0) AS BewerberPunkte, ISNULL(dbo.BEWERB013.PUNKTE02013, 0) AS BewerberNSin3Jahren, 
                      ISNULL(dbo.BEWERB013.PUNKTE18013, 0) AS BewerberSonstNS200, ISNULL(dbo.BEWERB013.PUNKTE19013, 0) AS BewerberSonstNS500, 
                      ISNULL(dbo.BEWERB013.PUNKTE20013, 0) AS BewerberSonstNS1000, CASE WHEN (ISNULL(PUNKTE01013, 0) - ISNULL(PUNKTE02013, 0) 
                      - ISNULL(Punkte18013, 0) - ISNULL(punkte19013, 0) - ISNULL(punkte20013, 0)) < 0 THEN 0 ELSE (ISNULL(PUNKTE01013, 0) - ISNULL(PUNKTE02013, 0) 
                      - ISNULL(Punkte18013, 0) - ISNULL(punkte19013, 0) - ISNULL(punkte20013, 0)) END AS BewerberSonstNSab1000, 
                      dbo.RA010.ID010 AS BewerberPersID, ISNULL(dbo.BEWERB013.PUNKTE08013, 0) + ISNULL(dbo.BEWERB013.PUNKTE09013, 0) 
                      + ISNULL(dbo.BEWERB013.PUNKTE10013, 0) + ISNULL(dbo.BEWERB013.PUNKTE11013, 0) + ISNULL(dbo.BEWERB013.PUNKTE12013, 0) 
                      AS BewerberPunkteNSGesamt, ISNULL(dbo.BEWERB013.PUNKTE16013, 0) AS BewerberSonderpunkte, ISNULL(dbo.BEWERB013.PUNKTE13013, 0) 
                      AS BewerberHTFortbildungen3Jahre, ISNULL(dbo.BEWERB013.PUNKTE14013, 0) AS BewerberHTSonstFortbildungen, 
                      dbo.BEWERB013.FK012013 AS StellenID, ISNULL(dbo.BEWERB013.PUNKTE21013, '') AS BewerberMonateAnwalt
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

CREATE VIEW dbo.EmpfängerDaten
AS
SELECT     ID010 AS PersID, NACHNAME010 AS RANachname, ISNULL(VORNAME010, '') AS RAVorname, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Sehr geehrte Frau' ELSE 'Sehr geehrter Herr' END AS RAAnredeVoll, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Frau Rechtsanwältin' ELSE 'Herrn Rechtsanwalt' END AS RAAnredeFormal, 
                      KSTR010 + ' ' + KPLZ010 + ' ' + KORT010 AS RAAnschriftKanzlei, ISNULL(KSTR010, '') AS RAStrKanzlei, KPLZ010 + ' ' + KORT010 AS RAPLZOrtKanzlei, 
                      ISNULL(AZ010, '-') AS RAAZVI, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Sehr geerhrte Frau ' ELSE 'Sehr geehrter Herr ' END + NACHNAME010 AS RAAnredeVollName, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Frau Rechtsanwältin ' ELSE 'Herr Rechtsanwalt ' END + VORNAME010 + ' ' + NACHNAME010 AS RAAnredeFormalName,
                       ISNULL(VORNAME010, '') + ' ' + NACHNAME010 AS RAVornameNachname, NACHNAME010 + ', ' + ISNULL(VORNAME010, '') AS RANachnameVorname, 
                      CASE WHEN RA010.ANREDE010 = 'Fau' THEN 'Frau' ELSE 'Herr' END AS RAAnrede, GEB010 AS RAGebDat, ISNULL(AMTORT010, '-') AS RAAmtssitz, 
                      ANWALTSEIT010 AS RAAnwaltSeit, AG010 AS RAAGzugelassen, LG010 AS RALGZugelassen, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Rechtsanwältin ' ELSE 'Rechtsanwalt ' END AS RAAmtsbez, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'RA in ' ELSE 'RA ' END AS RAAmtsbezKurz, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Notarin' ELSE 'Notar ' END AS RANotar, ISNULL(CAST(EXANOTE010 AS Varchar(8)), '-') 
                      AS RANoteStaatsexamen, ISNULL(EXANOTE010, 0) * 5 AS RAPunkteStaasexamen
FROM         dbo.RA010

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.MitbewerberAbgesagtDaten
AS
SELECT     TOP 100 PERCENT dbo.RA010.ID010 AS PersID, ISNULL(dbo.RA010.NACHNAME010, '') + ', ' + ISNULL(dbo.RA010.VORNAME010, '') 
                      AS AMitbewerberVoll, dbo.RA010.NACHNAME010 AS AMitbewerberNachname, ISNULL(dbo.RA010.ANREDE010, '') AS AMitbewerberAnrede, 
                      ISNULL(dbo.RA010.TITEL010, '') AS AMitbewerberTitel, ISNULL(dbo.RA010.NAMEZUS010, '') AS AMitbewerberNamenszusatz, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Sehr geehrte Frau' ELSE 'Sehr geehrter Herr' END AS AMitbewerberAnredeVoll, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Frau Rechtsanwältin' ELSE 'Herrn Rechtsanwalt' END AS AMitbewerberAnredeFormal, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Sehr geerhrte Frau ' ELSE 'Sehr geehrter Herr ' END + dbo.RA010.NACHNAME010 AS AMitbewerberAnredeVollName,
                       CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Frau Rechtsanwältin ' ELSE 'Herr Rechtsanwalt ' END + dbo.RA010.VORNAME010 + ' ' + dbo.RA010.NACHNAME010
                       AS AMitbewerberAnredeFormalName, ISNULL(dbo.RA010.AMTORT010, '-') AS AMitbewerberAmtssitz, 
                      dbo.BEWERB013.EINGANG013 AS AMitbewerberBewDatum, dbo.RA010.AZ010 AS AMitbewerberAZVI, ISNULL(dbo.BEWERB013.RANG013, '') 
                      AS AMitbewerberRang, dbo.RA010.AG010 AS AMitbewerberAGzugelassen, dbo.RA010.LG010 AS AMitbewerberLGZugelassen, 
                      ISNULL(dbo.RA010.VORNAME010, '') AS AMitbewerberVorname, ISNULL(CAST(dbo.RA010.EXANOTE010 AS Varchar(8)), '-') 
                      AS AMitbewerberNoteStaatsEx, ISNULL(dbo.RA010.EXANOTE010, 0) * 5 AS AMitbewerberPunkteStaatsexamen, 
                      ISNULL(dbo.BEWERB013.PUNKTE05013, 0) AS AMitbewerberPunkteAnwalttätigkeit, ISNULL(dbo.BEWERB013.PUNKTE06013, 0) 
                      AS AMitbewerberPunkteFortbildungen3Jahre, ISNULL(dbo.BEWERB013.PUNKTE07013, 0) AS AMitbewerberPunkteSonstigeFortbildungen, 
                      ISNULL(dbo.BEWERB013.PUNKTE06013, 0) + ISNULL(dbo.BEWERB013.PUNKTE07013, 0) AS AMitbewerberPunkteFortbildungenGesamt, 
                      ISNULL(dbo.BEWERB013.PUNKTE08013, 0) AS AMitbewerberPunkteNS3Jahren, ISNULL(dbo.BEWERB013.PUNKTE09013, 0) 
                      AS AMitbewerberPunkteNS200, ISNULL(dbo.BEWERB013.PUNKTE10013, 0) AS AMitbewerberPunkteNS500, ISNULL(dbo.BEWERB013.PUNKTE11013, 
                      0) AS AMitbewerberPunkteNS1000, ISNULL(dbo.BEWERB013.PUNKTE12013, 0) AS AMitbewerberPunkteNSab1000, 
                      ISNULL(dbo.BEWERB013.PUNKTE03013, 0) AS AMitbewerberPunkte, ISNULL(dbo.BEWERB013.PUNKTE02013, 0) AS AMitbewerberNSin3Jahren, 
                      ISNULL(dbo.BEWERB013.PUNKTE18013, 0) AS AMitbewerberSonstNS200, ISNULL(dbo.BEWERB013.PUNKTE19013, 0) AS AMitbewerberSonstNS500, 
                      ISNULL(dbo.BEWERB013.PUNKTE20013, 0) AS AMitbewerberSonstNS1000, CASE WHEN (ISNULL(PUNKTE01013, 0) - ISNULL(PUNKTE02013, 0) 
                      - ISNULL(Punkte18013, 0) - ISNULL(punkte19013, 0) - ISNULL(punkte20013, 0)) < 0 THEN 0 ELSE (ISNULL(PUNKTE01013, 0) - ISNULL(PUNKTE02013, 0) 
                      - ISNULL(Punkte18013, 0) - ISNULL(punkte19013, 0) - ISNULL(punkte20013, 0)) END AS AMitbewerberSonstNSab1000, 
                      dbo.RA010.ID010 AS AMitbewerberPersID, ISNULL(dbo.BEWERB013.PUNKTE08013, 0) + ISNULL(dbo.BEWERB013.PUNKTE09013, 0) 
                      + ISNULL(dbo.BEWERB013.PUNKTE10013, 0) + ISNULL(dbo.BEWERB013.PUNKTE11013, 0) + ISNULL(dbo.BEWERB013.PUNKTE12013, 0) 
                      AS AMitbewerberPunkteNSGesamt, ISNULL(dbo.BEWERB013.PUNKTE16013, 0) AS AMitbewerberSonderpunkte, 
                      ISNULL(dbo.BEWERB013.PUNKTE13013, 0) AS AMitbewerberHTFortbildungen3Jahre, ISNULL(dbo.BEWERB013.PUNKTE14013, 0) 
                      AS AMitbewerberHTSonstFortbildungen, dbo.BEWERB013.FK012013 AS StellenID, ISNULL(dbo.BEWERB013.PUNKTE21013, 0) 
                      AS AMitbewerberMonateAnwalt, CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'RechtsanwÃ¤ltin ' ELSE 'Rechtsanwalt ' END AS AMitbewerberAmtsbez, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'RA in ' ELSE 'RA ' END AS AMitbewerberAmtsbezKurz, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Notarin' ELSE 'Notar ' END AS AMitbewerberNotar
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

CREATE VIEW dbo.MitbewerberDaten
AS
SELECT     TOP 100 PERCENT dbo.RA010.ID010 AS PersID, dbo.BEWERB013.ID013, ISNULL(dbo.RA010.NACHNAME010, '') 
                      + ', ' + ISNULL(dbo.RA010.VORNAME010, '') AS MitbewerberVoll, dbo.RA010.NACHNAME010 AS MitbewerberNachname, 
                      ISNULL(dbo.RA010.ANREDE010, '') AS MitbewerberAnrede, ISNULL(dbo.RA010.TITEL010, '') AS MitbewerberTitel, ISNULL(dbo.RA010.NAMEZUS010, '') 
                      AS MitbewerberNamenszusatz, dbo.RA010.GESCHL010, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Sehr geehrte Frau' ELSE 'Sehr geehrter Herr' END AS MitbewerberAnredeVoll, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Frau Rechtsanwältin' ELSE 'Herrn Rechtsanwalt' END AS MitbewerberAnredeFormal, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Sehr geerhrte Frau ' ELSE 'Sehr geehrter Herr ' END + dbo.RA010.NACHNAME010 AS MitbewerberAnredeVollName,
                       CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Frau Rechtsanwältin ' ELSE 'Herr Rechtsanwalt ' END + dbo.RA010.VORNAME010 + ' ' + dbo.RA010.NACHNAME010
                       AS MitbewerberAnredeFormalName, ISNULL(dbo.RA010.AMTORT010, '-') AS MitbewerberAmtssitz, 
                      dbo.BEWERB013.EINGANG013 AS MitbewerberBewDatum, dbo.RA010.AZ010 AS MitbewerberAZVI, ISNULL(dbo.BEWERB013.RANG013, '') 
                      AS MitbewerberRang, dbo.RA010.AG010 AS MitbewerberAGzugelassen, dbo.RA010.LG010 AS MitbewerberLGZugelassen, 
                      ISNULL(dbo.RA010.VORNAME010, '') AS MitbewerberVorname, ISNULL(CAST(dbo.RA010.EXANOTE010 AS Varchar(8)), '-') AS MitbewerberNoteStaatsEx, 
                      ISNULL(dbo.RA010.EXANOTE010, 0) * 5 AS MitbewerberPunkteStaatsexamen, ISNULL(dbo.BEWERB013.PUNKTE05013, 0) 
                      AS MitbewerberPunkteAnwalttätigkeit, ISNULL(dbo.BEWERB013.PUNKTE06013, 0) AS MitbewerberPunkteFortbildungen3Jahre, 
                      ISNULL(dbo.BEWERB013.PUNKTE07013, 0) AS MitbewerberPunkteSonstigeFortbildungen, ISNULL(dbo.BEWERB013.PUNKTE06013, 0) 
                      + ISNULL(dbo.BEWERB013.PUNKTE07013, 0) AS MitbewerberPunkteFortbildungenGesamt, ISNULL(dbo.BEWERB013.PUNKTE08013, 0) 
                      AS MitbewerberPunkteNS3Jahren, ISNULL(dbo.BEWERB013.PUNKTE09013, 0) AS MitbewerberPunkteNS200, ISNULL(dbo.BEWERB013.PUNKTE10013, 
                      0) AS MitbewerberPunkteNS500, ISNULL(dbo.BEWERB013.PUNKTE11013, 0) AS MitbewerberPunkteNS1000, ISNULL(dbo.BEWERB013.PUNKTE12013, 
                      0) AS MitbewerberPunkteNSab1000, ISNULL(dbo.BEWERB013.PUNKTE03013, 0) AS MitbewerberPunkte, ISNULL(dbo.BEWERB013.PUNKTE02013, 0) 
                      AS MitbewerberNSin3Jahren, ISNULL(dbo.BEWERB013.PUNKTE18013, 0) AS MitbewerberSonstNS200, ISNULL(dbo.BEWERB013.PUNKTE19013, 0) 
                      AS MitbewerberSonstNS500, ISNULL(dbo.BEWERB013.PUNKTE20013, 0) AS MitbewerberSonstNS1000, CASE WHEN (ISNULL(PUNKTE01013, 0) 
                      - ISNULL(PUNKTE02013, 0) - ISNULL(Punkte18013, 0) - ISNULL(punkte19013, 0) - ISNULL(punkte20013, 0)) < 0 THEN 0 ELSE (ISNULL(PUNKTE01013, 0) 
                      - ISNULL(PUNKTE02013, 0) - ISNULL(Punkte18013, 0) - ISNULL(punkte19013, 0) - ISNULL(punkte20013, 0)) END AS MitbewerberSonstNSab1000, 
                      dbo.RA010.ID010 AS MitbewerberPersID, ISNULL(dbo.BEWERB013.PUNKTE08013, 0) + ISNULL(dbo.BEWERB013.PUNKTE09013, 0) 
                      + ISNULL(dbo.BEWERB013.PUNKTE10013, 0) + ISNULL(dbo.BEWERB013.PUNKTE11013, 0) + ISNULL(dbo.BEWERB013.PUNKTE12013, 0) 
                      AS MitbewerberPunkteNSGesamt, ISNULL(dbo.BEWERB013.PUNKTE16013, 0) AS MitbewerberSonderpunkte, ISNULL(dbo.BEWERB013.PUNKTE13013, 
                      0) AS MitbewerberHTFortbildungen3Jahre, ISNULL(dbo.BEWERB013.PUNKTE14013, 0) AS MitbewerberHTSonstFortbildungen, 
                      dbo.BEWERB013.FK012013 AS StellenID, ISNULL(dbo.BEWERB013.PUNKTE21013, '') AS MitbewerberMonateAnwalt, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Rechtsanwältin ' ELSE 'Rechtsanwalt ' END AS MitbewerberAmtsbez, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'RA in ' ELSE 'RA ' END AS MitbewerberAmtsbezKurz, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Notarin' ELSE 'Notar ' END AS MitbewerberNotar
FROM         dbo.RA010 INNER JOIN
                      dbo.BEWERB013 ON dbo.RA010.ID010 = dbo.BEWERB013.FK010013
ORDER BY dbo.BEWERB013.RANG013

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE VIEW dbo.MitbewerberZugesagtDaten
AS
SELECT     TOP 100 PERCENT dbo.RA010.ID010 AS PersID, ISNULL(dbo.RA010.NACHNAME010, '') + ', ' + ISNULL(dbo.RA010.VORNAME010, '') 
                      AS ZMitbewerberVoll, dbo.RA010.NACHNAME010 AS ZMitbewerberNachname, ISNULL(dbo.RA010.ANREDE010, '') AS ZMitbewerberAnrede, 
                      ISNULL(dbo.RA010.TITEL010, '') AS ZMitbewerberTitel, ISNULL(dbo.RA010.NAMEZUS010, '') AS ZMitbewerberNamenszusatz, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Sehr geehrte Frau' ELSE 'Sehr geehrter Herr' END AS ZMitbewerberAnredeVoll, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Frau Rechtsanwältin' ELSE 'Herrn Rechtsanwalt' END AS ZMitbewerberAnredeFormal, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Sehr geerhrte Frau ' ELSE 'Sehr geehrter Herr ' END + dbo.RA010.NACHNAME010 AS ZMitbewerberAnredeVollName,
                       CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Frau Rechtsanwältin ' ELSE 'Herr Rechtsanwalt ' END + dbo.RA010.VORNAME010 + ' ' + dbo.RA010.NACHNAME010
                       AS ZMitbewerberAnredeFormalName, ISNULL(dbo.RA010.AMTORT010, '-') AS ZMitbewerberAmtssitz, 
                      dbo.BEWERB013.EINGANG013 AS ZMitbewerberBewDatum, dbo.RA010.AZ010 AS ZMitbewerberAZVI, ISNULL(dbo.BEWERB013.RANG013, '') 
                      AS ZMitbewerberRang, dbo.RA010.AG010 AS ZMitbewerberAGzugelassen, dbo.RA010.LG010 AS ZMitbewerberLGZugelassen, 
                      ISNULL(dbo.RA010.VORNAME010, '') AS ZMitbewerberVorname, ISNULL(CAST(dbo.RA010.EXANOTE010 AS Varchar(8)), '-') 
                      AS ZMitbewerberNoteStaatsEx, ISNULL(dbo.RA010.EXANOTE010, 0) * 5 AS ZMitbewerberPunkteStaatsexamen, 
                      ISNULL(dbo.BEWERB013.PUNKTE05013, 0) AS ZMitbewerberPunkteAnwalttätigkeit, ISNULL(dbo.BEWERB013.PUNKTE06013, 0) 
                      AS ZMitbewerberPunkteFortbildungen3Jahre, ISNULL(dbo.BEWERB013.PUNKTE07013, 0) AS ZMitbewerberPunkteSonstigeFortbildungen, 
                      ISNULL(dbo.BEWERB013.PUNKTE06013, 0) + ISNULL(dbo.BEWERB013.PUNKTE07013, 0) AS ZMitbewerberPunkteFortbildungenGesamt, 
                      ISNULL(dbo.BEWERB013.PUNKTE08013, 0) AS ZMitbewerberPunkteNS3Jahren, ISNULL(dbo.BEWERB013.PUNKTE09013, 0) 
                      AS ZMitbewerberPunkteNS200, ISNULL(dbo.BEWERB013.PUNKTE10013, 0) AS ZMitbewerberPunkteNS500, ISNULL(dbo.BEWERB013.PUNKTE11013, 
                      0) AS ZMitbewerberPunkteNS1000, ISNULL(dbo.BEWERB013.PUNKTE12013, 0) AS ZMitbewerberPunkteNSab1000, 
                      ISNULL(dbo.BEWERB013.PUNKTE03013, 0) AS ZMitbewerberPunkte, ISNULL(dbo.BEWERB013.PUNKTE02013, 0) AS ZMitbewerberNSin3Jahren, 
                      ISNULL(dbo.BEWERB013.PUNKTE18013, 0) AS ZMitbewerberSonstNS200, ISNULL(dbo.BEWERB013.PUNKTE19013, 0) AS ZMitbewerberSonstNS500, 
                      ISNULL(dbo.BEWERB013.PUNKTE20013, 0) AS ZMitbewerberSonstNS1000, CASE WHEN (ISNULL(PUNKTE01013, 0) - ISNULL(PUNKTE02013, 0) 
                      - ISNULL(Punkte18013, 0) - ISNULL(punkte19013, 0) - ISNULL(punkte20013, 0)) < 0 THEN 0 ELSE (ISNULL(PUNKTE01013, 0) - ISNULL(PUNKTE02013, 0) 
                      - ISNULL(Punkte18013, 0) - ISNULL(punkte19013, 0) - ISNULL(punkte20013, 0)) END AS ZMitbewerberSonstNSab1000, 
                      dbo.RA010.ID010 AS ZMitbewerberPersID, ISNULL(dbo.BEWERB013.PUNKTE08013, 0) + ISNULL(dbo.BEWERB013.PUNKTE09013, 0) 
                      + ISNULL(dbo.BEWERB013.PUNKTE10013, 0) + ISNULL(dbo.BEWERB013.PUNKTE11013, 0) + ISNULL(dbo.BEWERB013.PUNKTE12013, 0) 
                      AS ZMitbewerberPunkteNSGesamt, ISNULL(dbo.BEWERB013.PUNKTE16013, 0) AS ZMitbewerberSonderpunkte, 
                      ISNULL(dbo.BEWERB013.PUNKTE13013, 0) AS ZMitbewerberHTFortbildungen3Jahre, ISNULL(dbo.BEWERB013.PUNKTE14013, 0) 
                      AS ZMitbewerberHTSonstFortbildungen, dbo.BEWERB013.FK012013 AS StellenID, ISNULL(dbo.BEWERB013.PUNKTE21013, 0) 
                      AS ZMitbewerberMonateAnwalt, CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'RechtsanwÃ¤ltin ' ELSE 'Rechtsanwalt ' END AS ZMitbewerberAmtsbez, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'RA in ' ELSE 'RA ' END AS ZMitbewerberAmtsbezKurz, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Notarin' ELSE 'Notar ' END AS ZMitbewerberKurz
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


CREATE VIEW dbo.PunkteBewerber
AS
SELECT     ISNULL(CAST(dbo.RA010.EXANOTE010 AS Varchar(8)), '-') AS NoteStaatsEx, ISNULL(dbo.RA010.EXANOTE010, 0) * 5 AS PunkteStaatsexamen, 
                      ISNULL(dbo.BEWERB013.PUNKTE05013, 0) AS PunkteAnwalttätigkeit, ISNULL(dbo.BEWERB013.PUNKTE06013, 0) AS PunkteFortbildungen3Jahre, 
                      ISNULL(dbo.BEWERB013.PUNKTE07013, 0) AS PunkteSonstigeFortbildungen, ISNULL(dbo.BEWERB013.PUNKTE06013, 0) 
                      + ISNULL(dbo.BEWERB013.PUNKTE07013, 0) AS PunkteFortbildungenGesamt, ISNULL(dbo.BEWERB013.PUNKTE08013, 0) AS PunkteNS3Jahren, 
                      ISNULL(dbo.BEWERB013.PUNKTE09013, 0) AS PunkteNS200, ISNULL(dbo.BEWERB013.PUNKTE10013, 0) AS PunkteNS500, 
                      ISNULL(dbo.BEWERB013.PUNKTE11013, 0) AS PunkteNS1000, ISNULL(dbo.BEWERB013.PUNKTE12013, 0) AS PunkteNSab1000, 
                      ISNULL(dbo.BEWERB013.PUNKTE03013, 0) AS Punkte, dbo.STELLEN012.FRIST012 AS Bewebungsfrist, ISNULL(dbo.BEWERB013.PUNKTE02013, 0) 
                      AS NSin3Jahren, ISNULL(DATEDIFF([month], dbo.RA010.ANWALTSEIT010, dbo.STELLEN012.FRIST012), 0) AS MonateAnwalt, 
                      ISNULL(dbo.BEWERB013.PUNKTE18013, 0) AS SonstNS200, ISNULL(dbo.BEWERB013.PUNKTE19013, 0) AS SonstNS500, 
                      ISNULL(dbo.BEWERB013.PUNKTE20013, 0) AS SonstNS1000, CASE WHEN (ISNULL(PUNKTE01013, 0) - ISNULL(PUNKTE02013, 0) 
                      - ISNULL(Punkte18013, 0) - ISNULL(punkte19013, 0) - ISNULL(punkte20013, 0)) < 0 THEN 0 ELSE (ISNULL(PUNKTE01013, 0) - ISNULL(PUNKTE02013, 0) 
                      - ISNULL(Punkte18013, 0) - ISNULL(punkte19013, 0) - ISNULL(punkte20013, 0)) END AS SonstNSab1000, dbo.RA010.ID010 AS PersID, 
                      ISNULL(dbo.BEWERB013.PUNKTE08013, 0) + ISNULL(dbo.BEWERB013.PUNKTE09013, 0) + ISNULL(dbo.BEWERB013.PUNKTE10013, 0) 
                      + ISNULL(dbo.BEWERB013.PUNKTE11013, 0) + ISNULL(dbo.BEWERB013.PUNKTE12013, 0) AS PunkteNSGesamt, 
                      ISNULL(dbo.BEWERB013.PUNKTE16013, 0) AS Sonderpunkte, ISNULL(dbo.BEWERB013.PUNKTE13013, 0) AS HTFortbildungen3Jahre, 
                      ISNULL(dbo.BEWERB013.PUNKTE14013, 0) AS HTSonstFortbildungen
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


CREATE VIEW dbo.HTFortbildungen3Jahre
AS
SELECT     ISNULL(SUM(dbo.FORT011.ANZHT011), 0) AS HTFortbildungen3Jahre, dbo.RA010.ID010
FROM         dbo.RA010 LEFT OUTER JOIN
                      dbo.BEWERB013 ON dbo.RA010.ID010 = dbo.BEWERB013.FK010013 LEFT OUTER JOIN
                      dbo.AFORT014 ON dbo.RA010.ID010 = dbo.AFORT014.FK010014 LEFT OUTER JOIN
                      dbo.FORT011 ON dbo.FORT011.ID011 = dbo.AFORT014.FK011014 LEFT OUTER JOIN
                      dbo.STELLEN012 ON dbo.STELLEN012.ID012 = dbo.BEWERB013.FK012013 AND dbo.FORT011.DATUM011 >= DATEADD([Year], - 3, 
                      dbo.STELLEN012.FRIST012)
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

CREATE VIEW dbo.
SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE VIEW dbo.HTSonstFortbildungen
AS
SELECT     ISNULL(SUM(dbo.FORT011.ANZHT011), 0) AS HTSonstFortbildungen, dbo.RA010.ID010
FROM         dbo.RA010 LEFT OUTER JOIN
                      dbo.BEWERB013 ON dbo.RA010.ID010 = dbo.BEWERB013.FK010013 LEFT OUTER JOIN
                      dbo.STELLEN012 ON dbo.STELLEN012.ID012 = dbo.BEWERB013.FK012013 LEFT OUTER JOIN
                      dbo.AFORT014 ON dbo.RA010.ID010 = dbo.AFORT014.FK011014 LEFT OUTER JOIN
                      dbo.FORT011 ON dbo.FORT011.ID011 = dbo.AFORT014.FK011014 AND dbo.FORT011.DATUM011 < DATEADD([Year], - 3, dbo.STELLEN012.FRIST012)
GROUP BY dbo.RA010.ID010


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO
CREATE VIEW dbo.MitbewerberDaten
AS
SELECT     TOP 100 PERCENT dbo.EmpfängerDaten.PersID, dbo.EmpfängerDaten.PersID AS MitbewerberPersID, dbo.BEWERB013.ID013, 
                      dbo.EmpfängerDaten.RAVornameNachname AS MitbewerberVoll, dbo.EmpfängerDaten.RANachname AS MitbewerberNachname, 
                      dbo.EmpfängerDaten.RAVorname AS MitbewerberVorname, dbo.EmpfängerDaten.RAAnrede AS MitbewerberAnrede, 
                      dbo.EmpfängerDaten.RAAnredeVoll AS MitbewerberAnredeVoll, dbo.EmpfängerDaten.RAAnredeFormal AS MitbewerberAnredeFormal, 
                      dbo.EmpfängerDaten.RAAnredeVollName AS MitbewerberAnredeVollName, 
                      dbo.EmpfängerDaten.RAAnredeFormalName AS MitbewerberAnredeFormalName, dbo.EmpfängerDaten.RAAmtssitz AS MitbewerberAmtssitz, 
                      dbo.BEWERB013.EINGANG013 AS MitbewerberBewDatum, dbo.EmpfängerDaten.RAAZVI AS MitbewerberAZVI, ISNULL(dbo.BEWERB013.RANG013, '') 
                      AS MitbewerberRang, dbo.EmpfängerDaten.RAAGzugelassen AS MitbewerberAGzugelassen, 
                      dbo.EmpfängerDaten.RALGZugelassen AS MitbewerberLGZugelassen, ISNULL(dbo.BEWERB013.PUNKTE05013, 0) 
                      AS MitbewerberPunkteAnwalttätigkeit, ISNULL(dbo.BEWERB013.PUNKTE06013, 0) AS MitbewerberPunkteFortbildungen3Jahre, 
                      ISNULL(dbo.BEWERB013.PUNKTE07013, 0) AS MitbewerberPunkteSonstigeFortbildungen, ISNULL(dbo.BEWERB013.PUNKTE06013, 0) 
                      + ISNULL(dbo.BEWERB013.PUNKTE07013, 0) AS MitbewerberPunkteFortbildungenGesamt, ISNULL(dbo.BEWERB013.PUNKTE08013, 0) 
                      AS MitbewerberPunkteNS3Jahren, ISNULL(dbo.BEWERB013.PUNKTE09013, 0) AS MitbewerberPunkteNS200, ISNULL(dbo.BEWERB013.PUNKTE10013, 
                      0) AS MitbewerberPunkteNS500, ISNULL(dbo.BEWERB013.PUNKTE11013, 0) AS MitbewerberPunkteNS1000, ISNULL(dbo.BEWERB013.PUNKTE12013, 
                      0) AS MitbewerberPunkteNSab1000, ISNULL(dbo.BEWERB013.PUNKTE03013, 0) AS MitbewerberPunkte, ISNULL(dbo.BEWERB013.PUNKTE02013, 0) 
                      AS MitbewerberNSin3Jahren, ISNULL(dbo.BEWERB013.PUNKTE18013, 0) AS MitbewerberSonstNS200, ISNULL(dbo.BEWERB013.PUNKTE19013, 0) 
                      AS MitbewerberSonstNS500, ISNULL(dbo.BEWERB013.PUNKTE20013, 0) AS MitbewerberSonstNS1000, CASE WHEN (ISNULL(PUNKTE01013, 0) 
                      - ISNULL(PUNKTE02013, 0) - ISNULL(Punkte18013, 0) - ISNULL(punkte19013, 0) - ISNULL(punkte20013, 0)) < 0 THEN 0 ELSE (ISNULL(PUNKTE01013, 0) 
                      - ISNULL(PUNKTE02013, 0) - ISNULL(Punkte18013, 0) - ISNULL(punkte19013, 0) - ISNULL(punkte20013, 0)) END AS MitbewerberSonstNSab1000, 
                      ISNULL(dbo.BEWERB013.PUNKTE08013, 0) + ISNULL(dbo.BEWERB013.PUNKTE09013, 0) + ISNULL(dbo.BEWERB013.PUNKTE10013, 0) 
                      + ISNULL(dbo.BEWERB013.PUNKTE11013, 0) + ISNULL(dbo.BEWERB013.PUNKTE12013, 0) AS MitbewerberPunkteNSGesamt, 
                      ISNULL(dbo.BEWERB013.PUNKTE16013, 0) AS MitbewerberSonderpunkte, ISNULL(dbo.BEWERB013.PUNKTE13013, 0) 
                      AS MitbewerberHTFortbildungen3Jahre, ISNULL(dbo.BEWERB013.PUNKTE14013, 0) AS MitbewerberHTSonstFortbildungen, 
                      dbo.BEWERB013.FK012013 AS StellenID, ISNULL(dbo.BEWERB013.PUNKTE21013, '') AS MitbewerberMonateAnwalt, 
                      dbo.EmpfängerDaten.RAAmtsbez AS MitbewerberAmtsbez, dbo.EmpfängerDaten.RAAmtsbezKurz AS MitbewerberAmtsbezKurz, 
                      dbo.EmpfängerDaten.RANotar AS MitbewerberNotar, dbo.EmpfängerDaten.RANoteStaatsexamen AS MitbewerberNoteStaatsexamen, 
                      dbo.EmpfängerDaten.RAPunkteStaasexamen AS MitbewerberPunkteStaatsexamen
FROM         dbo.BEWERB013 RIGHT OUTER JOIN
                      dbo.EmpfängerDaten ON dbo.BEWERB013.FK010013 = dbo.EmpfängerDaten.PersID
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

-- Store Procs Aktualisieren
if exists (select * from dbo.sysobjects where id = object_id(N'[NOTARE_DeleteBewerberInfos]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [NOTARE_DeleteBewerberInfos]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[NOTARE_ListeAusgeschriebeneStellen]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [NOTARE_ListeAusgeschriebeneStellen]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[NOTARE_ListeBewerber]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [NOTARE_ListeBewerber]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[NOTARE_ListeBewerberAbgelehnt]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [NOTARE_ListeBewerberAbgelehnt]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[NOTARE_ListeBewerberZugesagt]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [NOTARE_ListeBewerberZugesagt]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[NOTARE_Punkteauflistung]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [NOTARE_Punkteauflistung]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE PROCEDURE NOTARE_DeleteBewerberInfos  AS

-- Note Staatsexamen lÃ¶schen
UPDATE RA010 SET RA010.EXANOTE010 = Null WHERE STATUS010 ='Notar' 
OR STATUS010 = 'ausgeschieden' OR STATUS010 ='a. D.'

-- BewerbungsdatensÃ¤tze lÃ¶schen
DELETE BEWERB013
FROM BEWERB013 INNER JOIN RA010 ON BEWERB013.FK010013 = RA010.ID010
WHERE (RA010.STATUS010 ='Notar' 
OR RA010.STATUS010 = 'ausgeschieden' OR RA010.STATUS010 ='a. D.')
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO

CREATE PROCEDURE NOTARE_ListeAusgeschriebeneStellen @ID020 varchar(50) AS

DECLARE @Bezirk varchar(255)							-- Bezirk der Ausgeschriebenen Stelle(n)
DECLARE @Anz varchar(10)							-- Anzahl der zu besetzenden Stellen
DECLARE @FormatString varchar(4000)                				-- Aktuell zusammengesetzter String
DECLARE @ResultString varchar(4000)						-- Gesamter Ergebnis String
DECLARE @Counter Varchar(3)							-- ZÃ¤hler
DECLARE @StrOffset int 							-- Zeichen Offset zwischen Bezirk und Anz Stellen
DECLARE @Buffer Varchar(15)

SET @ResultString =''
SET @Counter ='1'
SET @Buffer = '               '

DECLARE Stellen_cursor CURSOR FOR
SELECT AGName004, SUM(ANZ012) 
FROM AG004 Left JOIN (STELLEN012 
	INNER JOIN AUSSCHREIBUNG020 ON ID020 = FK020012 
	) ON AGName004 = Bezirk012
WHERE ID020  = @ID020
GROUP BY AGName004
ORDER BY AGName004

OPEN Stellen_cursor								-- Cursor Ã¶ffnen
FETCH NEXT FROM Stellen_cursor						-- 1. DS
INTO @Bezirk, @Anz
	
WHILE @@FETCH_STATUS = 0						-- Alle DatensÃ¤tze durchlaufen
BEGIN
	SET @StrOffset = 15 - Len(@Bezirk)
	--SET @FormatString = @Counter + '. ' + @Bezirk + Left(@Buffer,@StrOffset) + @Anz
	SET @FormatString = @Counter + '. ' + @Bezirk + CHAR(9) + CHAR(9) + CHAR(9) + @Anz
	SET @ResultString = @ResultString  + CHAR(13) + @FormatString
	SET @Counter = @Counter +1 
	-- nÃ¤chsten DS
	FETCH NEXT FROM Stellen_cursor					-- NÃ¤chsten DS
	INTO @Bezirk, @Anz	
END
	
SELECT @ResultString as ListeAusgeschreibeneStellen				-- Ergebniss zurÃ¼ck

-- print @ResultString
CLOSE Stellen_cursor								-- Cursor Schliessen
DEALLOCATE Stellen_cursor
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS OFF 
GO

CREATE PROCEDURE NOTARE_ListeBewerber  @ID012 uniqueidentifier   AS


DECLARE @RangNr  Varchar(3)
DECLARE @Punkte varchar(10)
DECLARE @Anrede varchar(50)
DECLARE @Vorname  Varchar(255)
DECLARE @Nachname Varchar(255) 
DECLARE @FormatString varchar(2000)
DECLARE @ResultRangNamePunkte varchar(4000)
DECLARE @ResultRANachname varchar(4000)

SET @ResultRangNamePunkte =''
SET @ResultRANachname =''

-- Cursor fÃ¼r Stellendaten
DECLARE Bewerber_cursor CURSOR FOR
SELECT Vorname010, Nachname010, Anrede010,Rang013, PUNKTE03013  
FROM Bewerb013 INNER JOIN RA010 on FK010013 = ID010
WHERE FK012013 = Cast(@ID012 as uniqueidentifier)
ORDER BY Rang013

-- Cursor Ã¶ffnen
OPEN Bewerber_cursor
-- 1. DS
FETCH NEXT FROM Bewerber_cursor
INTO @Vorname, @Nachname, @Anrede, @RangNr , @Punkte
	
-- Alle DatensÃ¤tze durchlaufen
WHILE @@FETCH_STATUS = 0
BEGIN

	SET @FormatString = @RangNr + ')' + CHAR(9) +  @Anrede + ' ' + @Nachname + ', ' + @Vorname + CHAR(9) + CHAR(9) + CHAR(9) + @Punkte + ' Punkte ' 
	SET @ResultRangNamePunkte = @ResultRangNamePunkte + CHAR(13) + @FormatString



	SET @FormatString =   CASE WHEN @Anrede = 'Frau' THEN 'RechtsanwÃ¤ltin ' ELSE 'Rechtsanwalt ' END  +  '  ' + @Nachname
	SET @ResultRANachname = @ResultRANachname + @FormatString
   	--SELECT @FormatString
	--SELECT @RangNr + ') ' +  @Anrede + ' ' + @Nachname + ', ' + @Vorname + ' 		'  + @Punkte + ' Punkte ' 
	--SELECT @ResultString
	-- nÃ¤chsten DS
	FETCH NEXT FROM Bewerber_cursor
	INTO @Vorname, @Nachname, @Anrede, @RangNr , @Punkte	
END

SELECT @ResultRangNamePunkte as Bewerber_RangNamePunkte, @ResultRANachname as Bewerber_RANachname

-- print @ResultString
-- Cursor Schliessen
CLOSE Bewerber_cursor
DEALLOCATE Bewerber_cursor
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE PROCEDURE NOTARE_ListeBewerberAbgelehnt  @ID012 uniqueidentifier   AS


DECLARE @RangNr  Varchar(3)
DECLARE @Punkte varchar(10)
DECLARE @Anrede varchar(50)
DECLARE @Vorname  Varchar(255)
DECLARE @Nachname Varchar(255) 
DECLARE @FormatString varchar(2000)
DECLARE @ResultRangNamePunkte varchar(4000)
DECLARE @ResultRANachname varchar(4000)

SET @ResultRangNamePunkte =''
SET @ResultRANachname =''

-- Cursor fÃ¼r Stellendaten
DECLARE Bewerber_cursor CURSOR FOR
SELECT Vorname010, Nachname010, Anrede010,Rang013, PUNKTE03013  
FROM Bewerb013 INNER JOIN RA010 on FK010013 = ID010
WHERE FK012013 = Cast(@ID012 as uniqueidentifier) AND ISNULL(Zusage013,0) = 0
ORDER BY Rang013

-- Cursor Ã¶ffnen
OPEN Bewerber_cursor
-- 1. DS
FETCH NEXT FROM Bewerber_cursor
INTO @Vorname, @Nachname, @Anrede, @RangNr , @Punkte
	
-- Alle DatensÃ¤tze durchlaufen
WHILE @@FETCH_STATUS = 0
BEGIN

	SET @FormatString = @RangNr + ')' + CHAR(9) +  @Anrede + ' ' + @Nachname + ', ' + @Vorname + CHAR(9) + CHAR(9) + CHAR(9) + @Punkte + ' Punkte ' 
	SET @ResultRangNamePunkte = @ResultRangNamePunkte + CHAR(13) + @FormatString



	SET @FormatString =   CASE WHEN @Anrede = 'Frau' THEN 'RechtsanwÃ¤ltin ' ELSE 'Rechtsanwalt ' END  +  '  ' + @Nachname
	SET @ResultRANachname = @ResultRANachname + @FormatString
   	--SELECT @FormatString
	--SELECT @RangNr + ') ' +  @Anrede + ' ' + @Nachname + ', ' + @Vorname + ' 		'  + @Punkte + ' Punkte ' 
	--SELECT @ResultString
	-- nÃ¤chsten DS
	FETCH NEXT FROM Bewerber_cursor
	INTO @Vorname, @Nachname, @Anrede, @RangNr , @Punkte	
END

SELECT @ResultRangNamePunkte as Bewerber_RangNamePunkte, @ResultRANachname as Bewerber_RANachname

-- print @ResultString
-- Cursor Schliessen
CLOSE Bewerber_cursor
DEALLOCATE Bewerber_cursor
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE PROCEDURE NOTARE_ListeBewerberZugesagt  @ID012 uniqueidentifier   AS


DECLARE @RangNr  Varchar(3)
DECLARE @Punkte varchar(10)
DECLARE @Anrede varchar(50)
DECLARE @Vorname  Varchar(255)
DECLARE @Nachname Varchar(255) 
DECLARE @FormatString varchar(2000)
DECLARE @ResultRangNamePunkte varchar(4000)
DECLARE @ResultRANachname varchar(4000)

SET @ResultRangNamePunkte =''
SET @ResultRANachname =''

-- Cursor fÃ¼r Stellendaten
DECLARE Bewerber_cursor CURSOR FOR
SELECT Vorname010, Nachname010, Anrede010,Rang013, PUNKTE03013  
FROM Bewerb013 INNER JOIN RA010 on FK010013 = ID010
WHERE FK012013 = Cast(@ID012 as uniqueidentifier) AND ISNULL(Zusage013,0) = -1
ORDER BY Rang013

-- Cursor Ã¶ffnen
OPEN Bewerber_cursor
-- 1. DS
FETCH NEXT FROM Bewerber_cursor
INTO @Vorname, @Nachname, @Anrede, @RangNr , @Punkte
	
-- Alle DatensÃ¤tze durchlaufen
WHILE @@FETCH_STATUS = 0
BEGIN

	SET @FormatString = @RangNr + ')' + CHAR(9) +  @Anrede + ' ' + @Nachname + ', ' + @Vorname + CHAR(9) + CHAR(9) + CHAR(9) + @Punkte + ' Punkte ' 
	SET @ResultRangNamePunkte = @ResultRangNamePunkte + CHAR(13) + @FormatString



	SET @FormatString =   CASE WHEN @Anrede = 'Frau' THEN 'RechtsanwÃ¤ltin ' ELSE 'Rechtsanwalt ' END  +  '  ' + @Nachname
	SET @ResultRANachname = @ResultRANachname + @FormatString
   	--SELECT @FormatString
	--SELECT @RangNr + ') ' +  @Anrede + ' ' + @Nachname + ', ' + @Vorname + ' 		'  + @Punkte + ' Punkte ' 
	--SELECT @ResultString
	-- nÃ¤chsten DS
	FETCH NEXT FROM Bewerber_cursor
	INTO @Vorname, @Nachname, @Anrede, @RangNr , @Punkte	
END

SELECT @ResultRangNamePunkte as Bewerber_RangNamePunkte, @ResultRANachname as Bewerber_RANachname

-- print @ResultString
-- Cursor Schliessen
CLOSE Bewerber_cursor
DEALLOCATE Bewerber_cursor
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO

CREATE PROCEDURE NOTARE_Punkteauflistung @ID010 uniqueidentifier  AS


DECLARE @a Varchar(4000)
SET @a = (
SELECT '1.' + CHAR(9) +	'StaatsprÃ¼fung (Â§ 6 Abs. 2 Nr. 1 AVNot):' + CHAR(9) + CHAR(9) + CHAR(9) + CHAR(9) + CAST(ISNULL(dbo.RA010.EXANOTE010, 0) * 5 AS varchar(10)) + ' Punkte' + CHAR(13) 
+ CHAR(9) + '(' + ISNULL(CAST(dbo.RA010.EXANOTE010 AS Varchar(8)), '-') + ' Punkte x 5)' + CHAR(13) 
+ '2.' + CHAR(9) + 'Dauer der AnwaltstÃ¤tigkeit (Â§ 6 Abs. 2 Nr. 2 AVNot):' + CHAR(9) + + CHAR(9) + CAST(ISNULL(dbo.BEWERB013.PUNKTE05013, 0) as varchar(5)) + ' Punkte' + CHAR(13) + CHAR(9) + '(' + CAST(ISNULL(DATEDIFF([month], dbo.RA010.ANWALTSEIT010, dbo.STELLEN012.FRIST012), 0) as Varchar(4)) + 'Monate x 0,25 einschlieÃŸlich 15 Monate Grundwehrdienst;' + CHAR(13)
+ CHAR(9) + 'jedoch maximal 45 Punkte)' + CHAR(13) 
+ '3.' + CHAR(9) + 'Fortbildungskurse (Â§ 6 Abs. 2 Nr. 3 AVNot):' + CHAR(13)
+ CHAR(9) + 'a) innerhalb der letzen drei Jahre: ' + CAST(ISNULL(dbo.BEWERB013.PUNKTE13013, 0) as Varchar(4)) + ' x 0,6' + CHAR(9) + CHAR(9) + CAST(ISNULL(dbo.BEWERB013.PUNKTE06013, 0) as varchar(4)) + ' Punkte' + CHAR(13)
+ CHAR(9) + 'b) sonstige Kurse: ' + Cast(ISNULL(dbo.BEWERB013.PUNKTE14013, 0) as varchar(4)) + ' x 0,3' + CHAR(9) + CHAR(9) + CHAR(9) + CHAR(9) + CAST(ISNULL(dbo.BEWERB013.PUNKTE07013, 0) as varchar(4)) + ' Punkte' + CHAR(13)
+ CHAR(9) + 'Insgesamt:' + CHAR(9) + CHAR(9) + CHAR(9) + CHAR(9) + CHAR(9) + CHAR(9) + CHAR(9) + CAST(ISNULL(dbo.BEWERB013.PUNKTE07013, 0) as varchar(4)) + ' Punkte' + CHAR(13)
+ '4.' + CHAR(9) + 'Niederschriften (Â§ 6 Abs. 2 Nr. 4 AVNot):' + CHAR(13)
+ CHAR(9) + 'a) innerhalb der letzten drei Jahre bei mindestens ' + CHAR(9) + CHAR(9) + CAST(ISNULL(dbo.BEWERB013.PUNKTE08013, 0) as varchar(4)) + ' Punkte' + CHAR(13)
+ CHAR(9) + 'zweiwÃ¶chiger Vertretung: ' + CAST(ISNULL(dbo.BEWERB013.PUNKTE02013, 0) as varchar(6)) + ' x 0,2' + CHAR(13)
+ CHAR(9) + 'b) sonstige Vertretungen: ' + Cast(ISNULL(dbo.BEWERB013.PUNKTE18013, 0) as varchar(6)) + ' x 0,1' + CHAR(9) + CHAR(9) + CHAR(9) + CHAR(9) + Cast(ISNULL(dbo.BEWERB013.PUNKTE09013, 0) as varchar(6)) + ' Punkte'

FROM dbo.RA010 LEFT OUTER JOIN
                      dbo.BEWERB013 ON dbo.RA010.ID010 = dbo.BEWERB013.FK010013 LEFT OUTER JOIN
                      dbo.STELLEN012 ON dbo.STELLEN012.ID012 = dbo.BEWERB013.FK012013
WHERE ID010 = '{E682E548-A8C3-44ED-860C-2F798B8C3864}'
)
Print @a
GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO