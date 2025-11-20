
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
NULL, NULL, NULL, -1, 'PUNKTESUM013', NULL)
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

--***************************
-- DB Version Eintragen
--***************************
DECLARE @OldVers Varchar(5)
DECLARE @NewVers Varchar(5)
DECLARE @SkriptBesch Varchar(255)

SET @NewVers = '1.1'
SET @SkriptBesch =''

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