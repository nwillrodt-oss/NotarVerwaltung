
SET NOCOUNT ON

--***************************
-- UpdateHistory025 Neu in DB v 1.1
--***************************

-- Evtl. vorhandene Tmp Tablele löschen
IF exists (SELECT * FROM sysobjects WHERE id = object_id(N'UpdateHistory025_tmp') and OBJECTPROPERTY(id, N'IsUserTable') = 1) 
	DROP TABLE UpdateHistory025_tmp 

-- Tabelle mit Suffix _tmp anlegen
CREATE TABLE [UpdateHistory025_tmp] (
	[UpdateDate025] [datetime] Not NULL ,
	[DBVersion_Vor025] [varchar] (5) NULL ,
	[DBVersion_Nach025] [varchar] (5) NULL ,
	[Beschreibung025] [varchar] (255) NULL ,
	[Benutzer025] [varchar] (50) NULL,
	[SQL_SRV_Version025] Varchar(255) NULL	
) ON [PRIMARY]
GO

-- Wenn Tabelle existiert
IF EXISTS (SELECT * FROM sysobjects WHERE id = object_id(N'UpdateHistory025') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
BEGIN
	-- Daten in tmp Tabelle kopieren
	INSERT INTO UpdateHistory025_tmp (UpdateDate025, DBVersion_Vor025, DBVersion_Nach025, Beschreibung025, Benutzer025, SQL_SRV_Version025)
	SELECT UpdateDate025, DBVersion_Vor025, DBVersion_Nach025, Beschreibung025, Benutzer025, SQL_SRV_Version025
	FROM UpdateHistory025
	-- Datensätze durchzählen
	DECLARE @OrgCount int
	DECLARE @tmpCount int
	SET @OrgCount = (SELECT count(*) FROM UpdateHistory025)
	SET @tmpCount = (SELECT count(*) FROM UpdateHistory025_tmp)
	IF @OrgCount <> @tmpCount
	BEGIN
		-- Meckermeldung wenn ungleich
		SET ANSI_WARNINGS ON
		PRINT 'Fehler beim Import der Daten aus der Tabelle UpdateHistory025'
		SET ANSI_WARNINGS OFF
	END
	ELSE
	BEGIN
		-- org tabelle löschen
		DROP TABLE UpdateHistory025
	END	
END
GO
-- tmp tabelle umbenennen
sp_rename UpdateHistory025_tmp , UpdateHistory025
GO

--***************************
-- Neue Spalte in Bewerb013
--***************************
ALTER TABLE BEWERB013 ADD [PUNKTESUM013] [float] NULL;
GO

--***************************
-- Löschen Sicht HTFortbildungen3Jahre
--***************************
IF  EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[HTFortbildungen3Jahre]'))
DROP VIEW [HTFortbildungen3Jahre]
GO
--***************************
-- Löschen Sicht HTSonstFortbildungen
--***************************
IF  EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[HTSonstFortbildungen]'))
DROP VIEW [HTSonstFortbildungen]
GO

--***************************
-- Anpassung Sicht BewerberDaten
--***************************

-- Löschen
IF  EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[BewerberDaten]'))
DROP VIEW [BewerberDaten]
GO

-- Neuanlegen
CREATE VIEW [BewerberDaten]
AS
SELECT     RA010.ID010 AS PersID, BEWERB013.EINGANG013 AS BewerberBewDatum, ISNULL(BEWERB013.RANG013, '') AS BewerberRang, 
                      ISNULL(CAST(BEWERB013.PUNKTE01013 AS Varchar(8)), '-') AS BewerberNoteStaatsEx, 
					  ISNULL(BEWERB013.PUNKTE03013, 0) AS BewerberPunkteStaatsexamen,  
						ISNULL(CAST(BEWERB013.PUNKTE02013 AS Varchar(8)), '-') AS BewerberNoteFachprüfung, 
					  ISNULL(BEWERB013.PUNKTE04013, 0) AS BewerberPunkteFachprüfung,  
                      ISNULL(BEWERB013.PUNKTESUM013, 0) AS BewerberPunkte,  
                      RA010.ID010 AS BewerberPersID, 
                      BEWERB013.FK012013 AS StellenID 
FROM         RA010 INNER JOIN
                      BEWERB013 ON RA010.ID010 = BEWERB013.FK010013
GO

--***************************
-- Anpassung Sicht MitbewerberAbgesagtDaten
--***************************

-- Löschen
IF  EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[MitbewerberAbgesagtDaten]'))
DROP VIEW [MitbewerberAbgesagtDaten]
GO

-- Anlegen
CREATE VIEW [MitbewerberAbgesagtDaten]
AS
SELECT     TOP 100 PERCENT RA010.ID010 AS PersID, CASE WHEN ISNULL(RA010.TITEL010, '') = '' THEN ISNULL(RA010.NACHNAME010, '') 
                      + ', ' + ISNULL(RA010.VORNAME010, '') ELSE ISNULL(RA010.TITEL010, '') + ' ' + ISNULL(RA010.NACHNAME010, '') 
                      + ', ' + ISNULL(RA010.VORNAME010, '') END AS AMitbewerberVoll, CASE WHEN ISNULL(RA010.TITEL010, '') 
                      = '' THEN RA010.NACHNAME010 ELSE ISNULL(RA010.TITEL010, '') + ' ' + RA010.NACHNAME010 END AS AMitbewerberNachname, 
                      ISNULL(RA010.ANREDE010, '') AS AMitbewerberAnrede, ISNULL(RA010.TITEL010, '') AS AMitbewerberTitel, ISNULL(RA010.NAMEZUS010, 
                      '') AS AMitbewerberNamenszusatz, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Sehr geehrte Frau' ELSE 'Sehr geehrter Herr' END AS AMitbewerberAnredeVoll, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Frau Rechtsanwältin' ELSE 'Herrn Rechtsanwalt' END AS AMitbewerberAnredeFormal, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Sehr geerhrte Frau ' ELSE 'Sehr geehrter Herr ' END + CASE WHEN ISNULL(TITEL010, '') 
                      = '' THEN RA010.NACHNAME010 ELSE ISNULL(TITEL010, '') + ' ' + RA010.NACHNAME010 END AS AMitbewerberAnredeVollName, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Frau Rechtsanwältin ' ELSE 'Herr Rechtsanwalt ' END + CASE WHEN ISNULL(TITEL010, '') 
                      = '' THEN ISNULL(RA010.VORNAME010, '') + ' ' + RA010.NACHNAME010 ELSE ISNULL(TITEL010, '') + ' ' + ISNULL(RA010.VORNAME010, '') 
                      + ' ' + RA010.NACHNAME010 END AS AMitbewerberAnredeFormalName, ISNULL(RA010.AMTORT010, '-') AS AMitbewerberAmtssitz, 
                      BEWERB013.EINGANG013 AS AMitbewerberBewDatum, RA010.AZ010 AS AMitbewerberAZVI, ISNULL(BEWERB013.RANG013, '') 
                      AS AMitbewerberRang, RA010.AG010 AS AMitbewerberAGzugelassen, RA010.LG010 AS AMitbewerberLGZugelassen, 
                      ISNULL(RA010.VORNAME010, '') AS AMitbewerberVorname, ISNULL(CAST(BEWERB013.PUNKTE01013 AS Varchar(8)), '-') 
                      AS AMitbewerberNoteStaatsEx, ISNULL(BEWERB013.PUNKTE03013, 0) AS AMitbewerberPunkteStaatsexamen, 
                       CAST(CONVERT(money,ISNULL(BEWERB013.PUNKTESUM013, 0)) AS varchar(30)) AS AMitbewerberPunkte,  
                      RA010.ID010 AS AMitbewerberPersID, BEWERB013.FK012013 AS StellenID, 
						CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'RechtsanwÃ¤ltin ' ELSE 'Rechtsanwalt ' END AS AMitbewerberAmtsbez,
                       CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'RA in ' ELSE 'RA ' END AS AMitbewerberAmtsbezKurz, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Notarin' ELSE 'Notar ' END AS AMitbewerberNotar
FROM         RA010 INNER JOIN
                      BEWERB013 ON RA010.ID010 = BEWERB013.FK010013
 WHERE  (BEWERB013.ZUSAGE013 = 0)
ORDER BY ISNULL(BEWERB013.RANG013, '')
GO

--***************************
-- Anpassung Sicht MitbewerberDaten
--***************************

-- Löschen
IF  EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[MitbewerberDaten]'))
DROP VIEW [MitbewerberDaten]
GO

-- Anlegen
CREATE VIEW [MitbewerberDaten]
AS
SELECT     TOP 100 PERCENT EmpfängerDaten.PersID, EmpfängerDaten.PersID AS MitbewerberPersID, BEWERB013.ID013, 
                      EmpfängerDaten.RAVornameNachname AS MitbewerberVoll, EmpfängerDaten.RANachname AS MitbewerberNachname, 
                      EmpfängerDaten.RAVorname AS MitbewerberVorname, EmpfängerDaten.RAAnrede AS MitbewerberAnrede, 
                      EmpfängerDaten.RAAnredeVoll AS MitbewerberAnredeVoll, EmpfängerDaten.RAAnredeFormal AS MitbewerberAnredeFormal, 
                      EmpfängerDaten.RAAnredeVollName AS MitbewerberAnredeVollName, 
                      EmpfängerDaten.RAAnredeFormalName AS MitbewerberAnredeFormalName, EmpfängerDaten.RAAmtssitz AS MitbewerberAmtssitz, 
                      BEWERB013.EINGANG013 AS MitbewerberBewDatum, EmpfängerDaten.RAAZVI AS MitbewerberAZVI, ISNULL(BEWERB013.RANG013, '') 
                      AS MitbewerberRang, EmpfängerDaten.RAAGzugelassen AS MitbewerberAGzugelassen, 
                      EmpfängerDaten.RALGZugelassen AS MitbewerberLGZugelassen,  
                       ISNULL(BEWERB013.PUNKTESUM013, 0) AS MitbewerberPunkte, 
                      BEWERB013.FK012013 AS StellenID,  IsNull(BEWERB013.PUNKTE02013,0) AS MitbewerberNoteFachprüfung,
						IsNull(BEWERB013.PUNKTE04013,0) AS MitbewerberPunkteFachprüfung,
                      EmpfängerDaten.RAAmtsbez AS MitbewerberAmtsbez, EmpfängerDaten.RAAmtsbezKurz AS MitbewerberAmtsbezKurz, 
                      EmpfängerDaten.RANotar AS MitbewerberNotar, BEWERB013.PUNKTE01013 AS MitbewerberNoteStaatsexamen, 
                      BEWERB013.PUNKTE03013 AS MitbewerberPunkteStaatsexamen
FROM         BEWERB013 RIGHT OUTER JOIN
                      EmpfängerDaten ON BEWERB013.FK010013 = EmpfängerDaten.PersID
ORDER BY ISNULL(BEWERB013.RANG013, '')
GO

--***************************
-- Anpassung Sicht MitbewerberZugesagtDaten
--***************************

-- Löschen
IF  EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[MitbewerberZugesagtDaten]'))
DROP VIEW [MitbewerberZugesagtDaten]
GO


-- Anlegen
CREATE VIEW [MitbewerberZugesagtDaten]
AS
SELECT     TOP 100 PERCENT RA010.ID010 AS PersID, CASE WHEN ISNULL(RA010.TITEL010, '') = '' THEN ISNULL(RA010.NACHNAME010, '') 
                      + ', ' + ISNULL(RA010.VORNAME010, '') ELSE ISNULL(RA010.TITEL010, '') + ' ' + ISNULL(RA010.NACHNAME010, '') 
                      + ', ' + ISNULL(RA010.VORNAME010, '') END AS ZMitbewerberVoll, CASE WHEN ISNULL(RA010.TITEL010, '') 
                      = '' THEN RA010.NACHNAME010 ELSE ISNULL(RA010.TITEL010, '') + ' ' + RA010.NACHNAME010 END AS ZMitbewerberNachname, 
                      ISNULL(RA010.ANREDE010, '') AS ZMitbewerberAnrede, ISNULL(RA010.TITEL010, '') AS ZMitbewerberTitel, ISNULL(RA010.NAMEZUS010, 
                      '') AS ZMitbewerberNamenszusatz, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Sehr geehrte Frau' ELSE 'Sehr geehrter Herr' END AS ZMitbewerberAnredeVoll, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Frau Rechtsanwältin' ELSE 'Herrn Rechtsanwalt' END AS ZMitbewerberAnredeFormal, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Sehr geerhrte Frau ' ELSE 'Sehr geehrter Herr ' END + CASE WHEN ISNULL(RA010.TITEL010, '') 
                      = '' THEN RA010.NACHNAME010 ELSE ISNULL(RA010.TITEL010, '') + ' ' + RA010.NACHNAME010 END AS ZMitbewerberAnredeVollName, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Frau Rechtsanwältin ' ELSE 'Herr Rechtsanwalt ' END + CASE WHEN ISNULL(RA010.TITEL010, '') 
                      = '' THEN ISNULL(RA010.VORNAME010, '') + ' ' + RA010.NACHNAME010 ELSE ISNULL(RA010.TITEL010, '') 
                      + ' ' + ISNULL(RA010.VORNAME010, '') + ' ' + RA010.NACHNAME010 END AS ZMitbewerberAnredeFormalName, 
                      ISNULL(RA010.AMTORT010, '-') AS ZMitbewerberAmtssitz, BEWERB013.EINGANG013 AS ZMitbewerberBewDatum, 
                      RA010.AZ010 AS ZMitbewerberAZVI, ISNULL(BEWERB013.RANG013, '') AS ZMitbewerberRang, 
                      RA010.AG010 AS ZMitbewerberAGzugelassen, RA010.LG010 AS ZMitbewerberLGZugelassen, ISNULL(RA010.VORNAME010, '') 
                      AS ZMitbewerberVorname, ISNULL(CAST(BEWERB013.PUNKTE01013 AS Varchar(8)), '-') AS ZMitbewerberNoteStaatsEx, 
						ISNULL(BEWERB013.PUNKTE03013, 0)  AS ZMitbewerberPunkteStaatsexamen, 
					 CAST(CONVERT(money, ISNULL(BEWERB013.PUNKTESUM013, 0)) AS varchar(30)) AS ZMitbewerberPunkte, 
                      RA010.ID010 AS ZMitbewerberPersID,  BEWERB013.FK012013 AS StellenID, 
                      ISNULL(CAST(BEWERB013.PUNKTE02013 AS Varchar(8)), '-') AS ZMitbewerberNoteFachprüfung, 
						ISNULL(BEWERB013.PUNKTE04013, 0)  AS ZMitbewerberPunkteFachprüfung,
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Rechtsanwältin ' ELSE 'Rechtsanwalt ' END AS ZMitbewerberAmtsbez, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'RA in ' ELSE 'RA ' END AS ZMitbewerberAmtsbezKurz, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Notarin' ELSE 'Notar ' END AS ZMitbewerberKurz
FROM         RA010 INNER JOIN
                      BEWERB013 ON RA010.ID010 = BEWERB013.FK010013
					  WHERE  (BEWERB013.ZUSAGE013 = - 1)
ORDER BY ISNULL(BEWERB013.RANG013, '')
GO			  

-- Löschen
IF  EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[PunkteBewerber]'))
DROP VIEW [PunkteBewerber]
GO

-- Anlegen
CREATE VIEW [PunkteBewerber]
AS
SELECT     ISNULL(CAST(BEWERB013.PUNKTE01013 AS Varchar(8)), '-') AS NoteStaatsEx, ISNULL(BEWERB013.PUNKTE03013, 0) AS PunkteStaatsexamen, 
                      ISNULL(CAST(BEWERB013.PUNKTE01013 AS Varchar(8)), '-') AS NoteFachprüfung, 
					  ISNULL(BEWERB013.PUNKTE03013, 0) AS PunkteFachprüfung,
                      ISNULL(BEWERB013.PUNKTESUM013, 0) AS Punkte, STELLEN012.FRIST012 AS Bewebungsfrist, 
					  RA010.ID010 AS PersID
FROM         RA010 LEFT OUTER JOIN
                      BEWERB013 ON RA010.ID010 = BEWERB013.FK010013 LEFT OUTER JOIN
                      STELLEN012 ON STELLEN012.ID012 = BEWERB013.FK012013
GO

--***************************
-- Anpassung StoreProc NOTARE_ListeBewerber
--***************************

-- Löschen
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[NOTARE_ListeBewerber]') AND type in (N'P', N'PC'))
DROP PROCEDURE [NOTARE_ListeBewerber]
GO

-- Anlegen
CREATE PROCEDURE [NOTARE_ListeBewerber]  @ID012 uniqueidentifier   AS

DECLARE @RangNr  Varchar(3)
DECLARE @Punkte varchar(30)
DECLARE @Titel varchar(20)
DECLARE @Anrede varchar(50)
DECLARE @Vorname  Varchar(255)
DECLARE @Nachname Varchar(255) 
DECLARE @FormatString varchar(2000)
DECLARE @ResultRangNamePunkte varchar(4000)
DECLARE @ResultRANachname varchar(4000)

SET @ResultRangNamePunkte =''
SET @ResultRANachname =''

-- Cursor für Stellendaten
DECLARE Bewerber_cursor CURSOR FOR
SELECT IsNull(Vorname010,''), Nachname010, ISNULL(Titel010,''), ISNULL(Anrede010,''),ISNULL(Rang013,''), CAST(CONVERT(money,ISNULL(PUNKTESUM013,0)) as varchar(30)) 
FROM Bewerb013 INNER JOIN RA010 on FK010013 = ID010
WHERE FK012013 = Cast(@ID012 as uniqueidentifier)
ORDER BY Rang013

-- Cursor öffnen
OPEN Bewerber_cursor
-- 1. DS
FETCH NEXT FROM Bewerber_cursor
INTO @Vorname, @Nachname, @Titel, @Anrede, @RangNr , @Punkte
	
-- Alle Datensätze durchlaufen
WHILE @@FETCH_STATUS = 0
BEGIN

	SET @FormatString = @RangNr + ')' + CHAR(9) +  @Anrede + ' ' + 
		CASE WHEN @Titel = '' THEN @Nachname ELSE @Titel + ' ' + @Nachname END + ', ' + @Vorname + CHAR(9) + CHAR(9) + CHAR(9) + @Punkte + ' Punkte ' 
	SET @ResultRangNamePunkte = @ResultRangNamePunkte + CHAR(13) + @FormatString

	SET @FormatString =   CASE WHEN @Anrede = 'Frau' THEN 'Rechtsanwältin ' ELSE 'Rechtsanwalt ' END  +  ' ' + 
		CASE WHEN @Titel = '' THEN @Nachname ELSE @Titel + ' ' + @Nachname END
	SET @ResultRANachname = @ResultRANachname  + CHAR(13) + @FormatString
	-- nächsten DS
	FETCH NEXT FROM Bewerber_cursor
	INTO @Vorname, @Nachname, @Titel, @Anrede, @RangNr , @Punkte
END

SELECT @ResultRangNamePunkte as ListeBewerber_RangNamePunkte, @ResultRANachname as ListeBewerber_RANachname

-- Cursor Schliessen
CLOSE Bewerber_cursor
DEALLOCATE Bewerber_cursor
GO

--***************************
-- Anpassung StoreProc NOTARE_ListeBewerberZugesagt
--***************************

-- Löschen
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[NOTARE_ListeBewerberZugesagt]') AND type in (N'P', N'PC'))
DROP PROCEDURE [NOTARE_ListeBewerberZugesagt]
GO


-- Anlegen
CREATE PROCEDURE [NOTARE_ListeBewerberZugesagt]  @ID012 uniqueidentifier   AS


DECLARE @RangNr  Varchar(3)
DECLARE @Punkte varchar(10)
DECLARE @Titel varchar(20)
DECLARE @Anrede varchar(50)
DECLARE @Vorname  Varchar(255)
DECLARE @Nachname Varchar(255) 
DECLARE @FormatString varchar(2000)
DECLARE @ResultRangNamePunkte varchar(4000)
DECLARE @ResultRANachname varchar(4000)

SET @ResultRangNamePunkte =''
SET @ResultRANachname =''

-- Cursor für Stellendaten
DECLARE Bewerber_cursor CURSOR FOR
SELECT IsNull(Vorname010,''), Nachname010, ISNULL(Titel010,''), ISNULL(Anrede010,''),ISNULL(Rang013,''), CAST(CONVERT(money,ISNULL(PUNKTESUM013,0)) as varchar(30)) 
FROM Bewerb013 INNER JOIN RA010 on FK010013 = ID010
WHERE FK012013 = Cast(@ID012 as uniqueidentifier) AND ISNULL(Zusage013,0) = -1
ORDER BY Rang013

-- Cursor öffnen
OPEN Bewerber_cursor
-- 1. DS
FETCH NEXT FROM Bewerber_cursor
INTO @Vorname, @Nachname, @Titel, @Anrede, @RangNr , @Punkte
	
-- Alle Datensätze durchlaufen
WHILE @@FETCH_STATUS = 0
BEGIN

	SET @FormatString = @RangNr + ')' + CHAR(9) +  @Anrede + ' ' + 
		CASE WHEN @Titel = '' THEN @Nachname ELSE @Titel + ' ' + @Nachname END + ', ' + @Vorname + CHAR(9) + CHAR(9) + CHAR(9) + @Punkte + ' Punkte ' 
	SET @ResultRangNamePunkte = @ResultRangNamePunkte + CHAR(13) + @FormatString

	SET @FormatString =   CASE WHEN @Anrede = 'Frau' THEN 'Rechtsanwältin ' ELSE 'Rechtsanwalt ' END  +  ' ' + 
		CASE WHEN @Titel = '' THEN @Nachname ELSE @Titel + ' ' + @Nachname END
	SET @ResultRANachname = @ResultRANachname + CHAR(13) + @FormatString

	-- nächsten DS
	FETCH NEXT FROM Bewerber_cursor
	INTO @Vorname, @Nachname, @Titel, @Anrede, @RangNr , @Punkte
END

SELECT @ResultRangNamePunkte as ListeBewerberZugesagt_RangNamePunkte, @ResultRANachname as ListeBewerberZugesagt_RANachname

-- Cursor Schliessen
CLOSE Bewerber_cursor
DEALLOCATE Bewerber_cursor
GO

--***************************
-- Anpassung StoreProc NOTARE_ListeBewerberAbgelehnt
--***************************

-- Löschen
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[NOTARE_ListeBewerberAbgelehnt]') AND type in (N'P', N'PC'))
DROP PROCEDURE [NOTARE_ListeBewerberAbgelehnt]
GO

-- anlegen
CREATE PROCEDURE [NOTARE_ListeBewerberAbgelehnt]  @ID012 uniqueidentifier   AS


DECLARE @RangNr  Varchar(3)
DECLARE @Punkte varchar(10)
DECLARE @Titel varchar(20)
DECLARE @Anrede varchar(50)
DECLARE @Vorname  Varchar(255)
DECLARE @Nachname Varchar(255) 
DECLARE @FormatString varchar(2000)
DECLARE @ResultRangNamePunkte varchar(4000)
DECLARE @ResultRANachname varchar(4000)

SET @ResultRangNamePunkte =''
SET @ResultRANachname =''

-- Cursor für Stellendaten
DECLARE Bewerber_cursor CURSOR FOR
SELECT IsNull(Vorname010,''), Nachname010, ISNULL(Titel010,''), ISNULL(Anrede010,''),ISNULL(Rang013,''), CAST(CONVERT(money,ISNULL(PUNKTESUM013,0)) as varchar(30)) 
FROM Bewerb013 INNER JOIN RA010 on FK010013 = ID010
WHERE FK012013 = Cast(@ID012 as uniqueidentifier) AND ISNULL(Zusage013,0) = 0
ORDER BY Rang013

-- Cursor öffnen
OPEN Bewerber_cursor
-- 1. DS
FETCH NEXT FROM Bewerber_cursor
INTO @Vorname, @Nachname, @Titel, @Anrede, @RangNr , @Punkte
	
-- Alle Datensätze durchlaufen
WHILE @@FETCH_STATUS = 0
BEGIN

	SET @FormatString = @RangNr + ')' + CHAR(9) +  @Anrede + ' ' + 
		CASE WHEN @Titel = '' THEN @Nachname ELSE @Titel + ' ' + @Nachname END + ', ' + @Vorname + CHAR(9) + CHAR(9) + CHAR(9) + @Punkte + ' Punkte ' 
	SET @ResultRangNamePunkte = @ResultRangNamePunkte + CHAR(13) + @FormatString

	SET @FormatString =   CASE WHEN @Anrede = 'Frau' THEN 'Rechtsanwältin ' ELSE 'Rechtsanwalt ' END  +  ' ' + 
		CASE WHEN @Titel = '' THEN @Nachname ELSE @Titel + ' ' + @Nachname END
	SET @ResultRANachname = @ResultRANachname + CHAR(13) + @FormatString
	
	-- nächsten DS
	FETCH NEXT FROM Bewerber_cursor
	INTO @Vorname, @Nachname, @Titel, @Anrede, @RangNr , @Punkte
END

SELECT @ResultRangNamePunkte as ListeBewerberAbgelehnt_RangNamePunkte, @ResultRANachname as ListeBewerberAbgelehnt_RANachname

-- print @ResultString
-- Cursor Schliessen
CLOSE Bewerber_cursor
DEALLOCATE Bewerber_cursor
GO


--***************************
-- Neue Punkteberechnung in DB v 1.1
--***************************

-- Alte Berechnungen Löschen
DELETE FROM [BERECHNUNGEN016]
GO

-- Neue Berechnungen eintragen
INSERT INTO [BERECHNUNGEN016] ([ID016], [FAKTOR016], [ORDER016], [VALUETYPE016], [MAXWERT016], [CAPTION016], 
[CAPTIONSQL016], [VALUESQL016], [MAXVALUE016], [LOCKED016], [SAVEFIELD016], [PUNKTESAVEFIELD016])
VALUES ('5ef32adf-5af4-46fe-ab85-550a1479b9de', 0.4, 1, 'float',NULL,'1. Note Große Jur.Staatsprüfung', 
NULL, NULL, 18, 0, 'PUNKTE01013', 'PUNKTE03013')

INSERT INTO [BERECHNUNGEN016] ([ID016], [FAKTOR016], [ORDER016], [VALUETYPE016], [MAXWERT016], [CAPTION016], 
[CAPTIONSQL016], [VALUESQL016], [MAXVALUE016], [LOCKED016], [SAVEFIELD016], [PUNKTESAVEFIELD016])
VALUES ('96ebfc16-ca87-48e5-9cd9-16e821433eb2', 0.6, 2, 'float',NULL,'2. Note der notariellen Fachprüfung', 
NULL, NULL, 18, 0, 'PUNKTE02013', 'PUNKTE04013')

INSERT INTO [BERECHNUNGEN016] ([ID016], [FAKTOR016], [ORDER016], [VALUETYPE016], [MAXWERT016], [CAPTION016], 
[CAPTIONSQL016], [VALUESQL016], [MAXVALUE016], [LOCKED016], [SAVEFIELD016], [PUNKTESAVEFIELD016])
VALUES ('44c096ea-a73f-4c1e-958e-c097c6b4e5dd', 0, 3, 'float',NULL,'Summe', 
NULL, NULL, NULL, -1, NULL, 'PUNKTESUM013')
GO

--***************************
-- Examensnote aus RA010 nach Berwerb013 übertragen
--***************************
UPDATE BEWERB013 SET PUNKTE01013 = EXANOTE010
FROM BEWERB013 INNER JOIN RA010 ON ID010 = FK010013
WHERE EXANOTE010 IS NOT NULL

-- dann noten an alter stelle löschen
UPDATE RA010 SET  EXANOTE010 = NULL
WHERE EXANOTE010 IS NOT NULL

-- Restliche felder Löschen
UPDATE BEWERB013 SET PUNKTE02013 = 0 , PUNKTE04013 =0

-- Schon mal ein bisschen rechnen
UPDATE BEWERB013 SET PUNKTE03013 = (PUNKTE01013 * (0.4)), PUNKTESUM013 = PUNKTE03013 + PUNKTE04013

--***************************
-- DB Version Eintragen
--***************************
DECLARE @OldVers Varchar(5)
DECLARE @NewVers Varchar(5)
DECLARE @SkriptBesch Varchar(255)

SET @NewVers = '1.1'
SET @SkriptBesch ='angepasste Punkteberechnung bei Bewerbern'

IF exists (SELECT * FROM OPTIONS023 WHERE  [Option023] = 'DBVERSION')
	BEGIN
		-- Alt Version merken
		SET @OldVers = (SELECT WERTTEXT023 FROM OPTIONS023 WHERE [Option023] = 'DBVERSION')
		-- Version updaten
		UPDATE OPTIONS023 SET WERTTEXT023 = @NewVers WHERE [Option023] = 'DBVERSION'

		-- in  Update History Eintragen
		IF exists (SELECT * FROM sysobjects WHERE id = object_id(N'UpdateHistory025') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
		BEGIN
			INSERT INTO UpdateHistory025 (UpdateDate025,DBVersion_Vor025,DBVersion_Nach025,Beschreibung025,Benutzer025, SQL_SRV_Version025)
			SELECT GetDate(),@OldVers,@NewVers,@SkriptBesch,CURRENT_USER,@@Version
		END
	END
	ELSE
	BEGIN
		INSERT INTO OPTIONS023 ([Option023], WERTTEXT023)
		VALUES ('DBVERSION',@NewVers)
	END
GO
SET NOCOUNT OFF