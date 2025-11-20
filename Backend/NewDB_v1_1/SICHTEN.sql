if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MitbewerberDaten]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[MitbewerberDaten]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[EmpfÑngerDaten]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[EmpfÑngerDaten]
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
						ISNULL(CAST(dbo.BEWERB013.PUNKTE02013 AS Varchar(8)), '-') AS BewerberNoteFachprÅfung, 
					  ISNULL(dbo.BEWERB013.PUNKTE04013, 0) AS BewerberPunkteFachprÅfung,  
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
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Frau RechtsanwÑltin' ELSE 'Herrn Rechtsanwalt' END AS AMitbewerberAnredeFormal, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Sehr geerhrte Frau ' ELSE 'Sehr geehrter Herr ' END + CASE WHEN ISNULL(TITEL010, '') 
                      = '' THEN dbo.RA010.NACHNAME010 ELSE ISNULL(TITEL010, '') + ' ' + dbo.RA010.NACHNAME010 END AS AMitbewerberAnredeVollName, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Frau RechtsanwÑltin ' ELSE 'Herr Rechtsanwalt ' END + CASE WHEN ISNULL(TITEL010, '') 
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
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Frau RechtsanwÑltin' ELSE 'Herrn Rechtsanwalt' END AS ZMitbewerberAnredeFormal, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Sehr geerhrte Frau ' ELSE 'Sehr geehrter Herr ' END + CASE WHEN ISNULL(dbo.RA010.TITEL010, '') 
                      = '' THEN dbo.RA010.NACHNAME010 ELSE ISNULL(dbo.RA010.TITEL010, '') + ' ' + dbo.RA010.NACHNAME010 END AS ZMitbewerberAnredeVollName, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Frau RechtsanwÑltin ' ELSE 'Herr Rechtsanwalt ' END + CASE WHEN ISNULL(dbo.RA010.TITEL010, '') 
                      = '' THEN ISNULL(dbo.RA010.VORNAME010, '') + ' ' + dbo.RA010.NACHNAME010 ELSE ISNULL(dbo.RA010.TITEL010, '') 
                      + ' ' + ISNULL(dbo.RA010.VORNAME010, '') + ' ' + dbo.RA010.NACHNAME010 END AS ZMitbewerberAnredeFormalName, 
                      ISNULL(dbo.RA010.AMTORT010, '-') AS ZMitbewerberAmtssitz, dbo.BEWERB013.EINGANG013 AS ZMitbewerberBewDatum, 
                      dbo.RA010.AZ010 AS ZMitbewerberAZVI, ISNULL(dbo.BEWERB013.RANG013, '') AS ZMitbewerberRang, 
                      dbo.RA010.AG010 AS ZMitbewerberAGzugelassen, dbo.RA010.LG010 AS ZMitbewerberLGZugelassen, ISNULL(dbo.RA010.VORNAME010, '') 
                      AS ZMitbewerberVorname, ISNULL(CAST(dbo.BEWERB013.PUNKTE01013 AS Varchar(8)), '-') AS ZMitbewerberNoteStaatsEx, 
						ISNULL(dbo.BEWERB013.PUNKTE03013, 0)  AS ZMitbewerberPunkteStaatsexamen, 
					 CAST(CONVERT(money, ISNULL(dbo.BEWERB013.PUNKTESUM013, 0)) AS varchar(30)) AS ZMitbewerberPunkte, 
                      dbo.RA010.ID010 AS ZMitbewerberPersID,  dbo.BEWERB013.FK012013 AS StellenID, 
                      ISNULL(CAST(dbo.BEWERB013.PUNKTE02013 AS Varchar(8)), '-') AS ZMitbewerberNoteFachprÅfung, 
						ISNULL(dbo.BEWERB013.PUNKTE04013, 0)  AS ZMitbewerberPunkteFachprÅfung,
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
                      ISNULL(CAST(dbo.BEWERB013.PUNKTE01013 AS Varchar(8)), '-') AS NoteFachprÅfung, 
					  ISNULL(dbo.BEWERB013.PUNKTE03013, 0) AS PunkteFachprÅfung,
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

CREATE VIEW dbo.EmpfÑngerDaten
AS
SELECT     dbo.RA010.ID010 AS PersID, CASE WHEN ISNULL(TITEL010, '') = '' THEN '' + NACHNAME010 ELSE ISNULL(TITEL010, '') 
                      + ' ' + NACHNAME010 END AS RANachname, ISNULL(dbo.RA010.VORNAME010, '') AS RAVorname, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Sehr geehrte Frau' ELSE 'Sehr geehrter Herr' END AS RAAnredeVoll, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Frau RechtsanwÑltin' ELSE 'Herrn Rechtsanwalt' END AS RAAnredeFormal, 
                      dbo.RA010.KSTR010 + ' ' + dbo.RA010.KPLZ010 + ' ' + dbo.RA010.KORT010 AS RAAnschriftKanzlei, ISNULL(dbo.RA010.KSTR010, '') AS RAStrKanzlei, 
                      dbo.RA010.KPLZ010 + ' ' + dbo.RA010.KORT010 AS RAPLZOrtKanzlei, ISNULL(dbo.RA010.AZ010, '-') AS RAAZVI, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Sehr geehrte Frau ' ELSE 'Sehr geehrter Herr ' END + CASE WHEN ISNULL(TITEL010, '') 
                      = '' THEN '' + NACHNAME010 ELSE ISNULL(TITEL010, '') + ' ' + NACHNAME010 END AS RAAnredeVollName, 
                      CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'Frau RechtsanwÑltin ' ELSE 'Herr Rechtsanwalt ' END + CASE WHEN ISNULL(TITEL010, '') 
                      = '' THEN ISNULL(VORNAME010, '') + ' ' + NACHNAME010 ELSE ISNULL(TITEL010, '') + ' ' + ISNULL(VORNAME010, '') 
                      + ' ' + NACHNAME010 END AS RAAnredeFormalName, CASE WHEN ISNULL(TITEL010, '') = '' THEN ISNULL(VORNAME010, '') 
                      + ' ' + NACHNAME010 ELSE ISNULL(TITEL010, '') + ' ' + ISNULL(VORNAME010, '') + ' ' + NACHNAME010 END AS RAVornameNachname, 
                      CASE WHEN ISNULL(TITEL010, '') = '' THEN '' + NACHNAME010 ELSE ISNULL(TITEL010, '') 
                      + ' ' + NACHNAME010 END + ', ' + ISNULL(dbo.RA010.VORNAME010, '') AS RANachnameVorname, 
                      CASE WHEN RA010.ANREDE010 = 'Fau' THEN 'Frau' ELSE 'Herr' END AS RAAnrede, dbo.RA010.GEB010 AS RAGebDat, 
                      ISNULL(dbo.RA010.AMTORT010, '-') AS RAAmtssitz, dbo.RA010.ANWALTSEIT010 AS RAAnwaltSeit, dbo.RA010.AG010 AS RAAGzugelassen, 
                      dbo.RA010.LG010 AS RALGZugelassen, CASE WHEN RA010.ANREDE010 = 'Frau' THEN 'RechtsanwÑltin ' ELSE 'Rechtsanwalt ' END AS RAAmtsbez, 
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
SELECT     TOP 100 PERCENT dbo.EmpfÑngerDaten.PersID, dbo.EmpfÑngerDaten.PersID AS MitbewerberPersID, dbo.BEWERB013.ID013, 
                      dbo.EmpfÑngerDaten.RAVornameNachname AS MitbewerberVoll, dbo.EmpfÑngerDaten.RANachname AS MitbewerberNachname, 
                      dbo.EmpfÑngerDaten.RAVorname AS MitbewerberVorname, dbo.EmpfÑngerDaten.RAAnrede AS MitbewerberAnrede, 
                      dbo.EmpfÑngerDaten.RAAnredeVoll AS MitbewerberAnredeVoll, dbo.EmpfÑngerDaten.RAAnredeFormal AS MitbewerberAnredeFormal, 
                      dbo.EmpfÑngerDaten.RAAnredeVollName AS MitbewerberAnredeVollName, 
                      dbo.EmpfÑngerDaten.RAAnredeFormalName AS MitbewerberAnredeFormalName, dbo.EmpfÑngerDaten.RAAmtssitz AS MitbewerberAmtssitz, 
                      dbo.BEWERB013.EINGANG013 AS MitbewerberBewDatum, dbo.EmpfÑngerDaten.RAAZVI AS MitbewerberAZVI, ISNULL(dbo.BEWERB013.RANG013, '') 
                      AS MitbewerberRang, dbo.EmpfÑngerDaten.RAAGzugelassen AS MitbewerberAGzugelassen, 
                      dbo.EmpfÑngerDaten.RALGZugelassen AS MitbewerberLGZugelassen,  
                       ISNULL(dbo.BEWERB013.PUNKTESUM013, 0) AS MitbewerberPunkte, 
                      dbo.BEWERB013.FK012013 AS StellenID,  IsNull(BEWERB013.PUNKTE02013,0) AS MitbewerberNoteFachprÅfung,
						IsNull(BEWERB013.PUNKTE04013,0) AS MitbewerberPunkteFachprÅfung,
                      dbo.EmpfÑngerDaten.RAAmtsbez AS MitbewerberAmtsbez, dbo.EmpfÑngerDaten.RAAmtsbezKurz AS MitbewerberAmtsbezKurz, 
                      dbo.EmpfÑngerDaten.RANotar AS MitbewerberNotar, dbo.BEWERB013.PUNKTE01013 AS MitbewerberNoteStaatsexamen, 
                      dbo.BEWERB013.PUNKTE03013 AS MitbewerberPunkteStaatsexamen
FROM         dbo.BEWERB013 RIGHT OUTER JOIN
                      dbo.EmpfÑngerDaten ON dbo.BEWERB013.FK010013 = dbo.EmpfÑngerDaten.PersID
ORDER BY ISNULL(dbo.BEWERB013.RANG013, '')



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

