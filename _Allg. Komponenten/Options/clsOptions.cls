VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit                                                     ' Variaben Deklaration erzwingen
Option Compare Text                                                 ' Sortierreihenfolge festlegen
Private Const MODULNAME = "clsOptions"                              ' Modulname für Fehlerbehandlung

' Nina: Fehlende Variablendeklarationen hinzugefügt
Private objObjectBag As Object
Private objError As Object
Private objTools As Object
Private objRegTools As Object

Private szIniFile As String
Private szAppTitel As String
Private szAppFolder As String
Private szAppRegRoot As String
Private szOptionIni As String


Public Sub Class_Terminate()
On Error Resume Next                                                ' Hier keine Fehler mehr
    Set objError = Nothing
    Set objTools = Nothing
    Set objRegTools = Nothing
    Set objObjectBag = Nothing
End Sub

Public Sub InitObjectBag(objOb As Object)
    Set objObjectBag = objOb                                        ' ObjectBag Obj. merken
    If Not objObjectBag Is Nothing Then                             ' nur wenn ObjectBag vorhanden
        Set objError = objObjectBag.GetErrorObj                     ' Error Obj. aus ObjectBag holen
        Set objTools = objObjectBag.GetToolsObj                     ' Tools Obj. aus ObjectBag holen
        Set objRegTools = objObjectBag.GetRegToolsObj               ' Registry Tools Obj. aus ObjectBag holen
        szIniFile = objObjectBag.GetIniFile                         ' XMLpafd aus ObjectBag holen
        szAppTitel = objObjectBag.GetAppTitle                       ' Anwendungstitel aus ObjectBag holen
        szAppFolder = objObjectBag.getappdir                        ' AnwendungsVerz aus ObjectBag holen
        szAppRegRoot = objObjectBag.getappregRoot
    End If
End Sub
                                                                    ' *****************************************
                                                                    ' Properties (Eigenschaften)
Public Function GetOptionCount() As Long
     GetOptionCount = OptionGetCount                                ' Options Anzahl
End Function

Public Sub SetOptionByName(szOptionName As String, Value As Variant)
    Call OptionSetByName(szOptionName, Value)                       ' Option mit Namen setzen
End Sub

Public Function GetOptionByName(szOptionName As String) As Variant
    GetOptionByName = OptionGetByName(szOptionName)                 ' Option mit Namen auslesen
End Function

Public Property Let SetOptioniniPath(szPath As String)
    If InStr(szPath, "\") = 0 Then
        szOptionIni = objObjectBag.getappdir & szPath
    Else
        szOptionIni = szPath
    End If
End Property

Public Sub SetOptionByCaption(szOptionCaption As String, Value As Variant)
    Call OptionSetByCaption(szOptionCaption, Value)
End Sub
                                                                    ' *****************************************
                                                                    ' Methoden
Public Sub InitOptions(Optional bForceIni As Boolean)
' Liest options aus ini (XML) in Array
' Schreibt diese in Reg (HKCU) wenn sie nicht existieren
' Übernimmt werte aus Reg wenn diese existieren
' Erst HKLM dann HKCU
    Dim i As Integer                                                ' counter
    Dim szDetail As String                                          ' Details für Fehlerbehandlung
    Dim szDefaultValue As String                                    ' Defaultwert
    Dim szOptionsList As String                                     ' ; getrente liste aller optionen
    Dim szOptListArray() As String                                  ' Optionslist als Array
    Dim NewVal As OptionValue                                       ' Neuen Oprions Wert holen
    
    ' Nina: Temporäre Variablen mit EXAKTEN Datentypen für .NET Interop (Wichtig!)
    Dim sName As String
    Dim sCaption As String
    Dim sValue As String    ' String statt Variant, da .NET hier exakten Typ erwartet
    Dim bCrypt As Boolean   ' Boolean statt Variant
    Dim bEdit As Boolean
    Dim bDisabled As Boolean
    Dim bPath As Boolean
    Dim bFile As Boolean
    Dim bBool As Boolean
    Dim bExpert As Boolean
    Dim sList As String     ' String statt Variant
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    
    szOptionsList = objTools.GetOptionsNodeList(szOptionIni)        ' Liste aller optionen holen
    If szOptionsList = "" Then GoTo exithandler                     ' Keine optionen -> fertig
    szOptListArray = Split(szOptionsList, ";")                      ' Optionsliste in Array
    
    For i = 0 To UBound(szOptListArray)
        Call PrepareOptionValue(NewVal)                             ' Option Value initialisieren
        With NewVal
            .Name = szOptListArray(i)
            
            ' Nina: Vorbelegung der Temp-Variablen
            sName = .Name
            sCaption = ""
            sValue = ""
            bCrypt = False
            bEdit = False
            bDisabled = False
            bPath = False
            bFile = False
            bBool = False
            bExpert = False
            sList = ""
            
            ' Nina: Aufruf mit strikten Typen für .NET
            Call objTools.GetOptionInfoFromXML(szOptionIni, sName, _
                    sCaption, sValue, bCrypt, bEdit, bDisabled, _
                    bPath, bFile, bBool, bExpert, sList)
            
            ' Nina: Werte zurückschreiben
            .Name = sName
            .Caption = sCaption
            .Value = sValue
            .bCrypt = bCrypt
            .bEdit = bEdit
            .bDisabled = bDisabled
            .bPath = bPath
            .bFile = bFile
            .bBool = bBool
            .bExpert = bExpert
            .szList = sList

            ' Wenn user admin dann in Lokal machine schreiben
            If .bPath Or .bFile Then                                ' Wenn Pfad angaben überprüfen ob Dynamische pfadangaben enthalten sind
                .Value = CheckSystemFolder(.Value, .bPath, .bFile)
            End If
            
            szDefaultValue = .Value                                 ' Dann mit Reg einträgen überschreiben
            If Not bForceIni Then                                   ' Wenn bForceIni = true dann keine Werte aus Reg. holen
                .Value = ReadOptionReg(.Name, szDefaultValue)
            Else
                Call WriteOptionReg(.Name, szDefaultValue)
            End If
            If .bCrypt Then                                         ' Falls Verschlüsselt
                .Value = objTools.Crypt(CStr(.Value), False)        ' Entschlüsseln
            End If
            If .bBool Then                                          ' Boolwerte behandeln
                If .Value = "" Then .Value = False
                .Value = CBool(.Value)
            End If
            
            Call InitOptionValue(.Name, .Caption, CStr(.Value), .Kategorie, .bCrypt, .bBool, .bEdit, _
                    .bPath, .bFile, .bDisabled, .bExpert, .szList)  ' In Options Array Schreiben
        End With
    Next i                                                          ' Nächster Eintrag
    
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
Exit Sub                                                            ' Sub Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "InitOptions", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Sub

Public Sub SaveOptions(Optional bINI As Boolean)
' Speichert die Optionen aus Array in der Registry
    Dim i As Integer                                                ' Counter
    Dim szRegKey As String                                          ' Regschlüssel
    Dim szOptName As String                                         ' Options Name
    Dim szOptValue As String                                        ' Options Value
    Dim bCrypt As Boolean                                           ' Options Verschlüsselt
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    szRegKey = "SOFTWARE\" & szAppRegRoot & "\"                       ' Reg Schlüssel festlegen
    For i = 0 To OptionGetCount                                     ' Alle Optionen durchlaufen
        szOptName = OptionGetNameByIndex(i)                         ' Options name holen
        szOptValue = OptionGetByName(szOptName)                     ' Optionswert holen
        bCrypt = OptionGetCyptByIndex(i)                            ' Verschüsselt ?
        If bCrypt Then szOptValue = objTools.Crypt(szOptValue, True) ' wenn ja -> verschlüsseln
        Call objRegTools.WriteRegValue("HKCU", szRegKey, szOptName, CStr(szOptValue))    ' In Reg Speichern
        If bINI Then                                                ' Auch in INI speichern
            ' Noch nicht realisiert
        End If
        szOptName = ""                                              ' Optionsname initialisieren
        szOptValue = ""                                             ' Optionswert initialisieren
        bCrypt = False                                              ' Verschlüsselt initialisieren
    Next i                                                          ' nächste Option
exithandler:
On Error Resume Next
Exit Sub
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On