--Select count(*) FROM Amtierende
--Select count(*) FROM Mitglied
INSERT INTO RA010 ( Vorname010, Nachname010, TITEL010, NAmezus010, AZ010, geb010, kstr010, kPLZ010, kOrt010,
	ktel010, kfax010, kemail010, status010, ag010, lg010, AnwaltSeit010, Create010, Modify010, cfrom010,mfrom010)
Select --* , 
Amtierende.Vorname AS Vorname010, Amtierende.name as Nachname010,
CASE When ISNULL(Mitglied.Titel,'')='' Then Amtierende.titel ELSE Mitglied.Titel END AS TITEL010 
,Amtierende.zusatz as NAmezus010, Cast(AZVI as Varchar(5)) as AZ010, 
Amtierende.gebdatum as gebdat010, Mitglied.Straﬂe as kstr010, 
Cast(ZustellPLZ as Varchar(5)) AS kPLZ010 , 
CASE WHEN ISNull(Mitglied.Ort,'')='' THEN amtssitz ELSE Mitglied.Ort END  as kOrt010
, Mitglied.Telefon as ktel010, Mitglied.fax as kfax010,Mitglied.email as kemail010, 'Notar' as status010,
[ZulassungAG] as ag010, [ZulassungLG] as lg010, 
CASE WHEN [Datum Zulassung LG] < [Datum Zulassung AG] AND [Datum Zulassung LG] < [Datum Zulassung OLG] THEN
	[Datum Zulassung AG]
     WHEN [Datum Zulassung AG] < [Datum Zulassung LG] AND [Datum Zulassung AG] < [Datum Zulassung OLG] THEN
	[Datum Zulassung AG]
     WHEN  [Datum Zulassung OLG] < [Datum Zulassung AG] AND [Datum Zulassung OLG] < [Datum Zulassung LG] THEN
	[Datum Zulassung OLG]
END AS AnwaltSeit010,
Getdate() as Create010, Getdate() as Modify010, 'Import' as cfrom010, 'Import' as mfrom010
From Amtierende left Join  Mitglied
	on Amtierende.name = Mitglieder_Name And Amtierende.vorname = Mitglied.Vorname