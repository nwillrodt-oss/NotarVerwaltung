Attribute VB_Name = "modOptions"
Option Explicit                                                     ' Variaben Deklaration erzwingen
Option Compare Text                                                 ' Sortierreihenfolge festlegen
Private Const MODULNAME = "modOptions"                              ' Modulname für Fehlerbehandlung

Public objError As Object                                           ' Error Object
Public objObjectBag As Object                                       ' ObjectBag object
Public objTools As Object
Public objRegTools As Object
Public szOptionIni As String                                        ' Pfad zur  options ini (jetzt XML)
Public szIniFile As String                                          ' Name der IniDatei
Public szAppTitel As String                                         ' Anwendungsname
Public szAppFolder As String                                        ' Anwendungsverzeichnis
Public szAppRegRoot As String                                       ' Basis Verz. in registry

Private VarOptions() As OptionValue                                 ' Optionen Array

Public Type OptionValue
    Name As String                                                  ' Name der Option
    Caption As String                                               ' Angez. Bezeichnung der option
    Value As Variant                                                ' Wert
    bCrypt As Boolean                                               ' Verschlüsselt
    bEdit As Boolean                                                ' von Anw. editierbar
    bDisabled As Boolean                                            ' Wir angzeigt aber nicht änderbar
    bBool As Boolean                                                ' Option ist ein Boolwert
    bPath As Boolean                                                ' Option ist eine Pfadangabe
    bFile As Boolean                                                ' Option ist eine Dateiangabe
    Kategorie As String                                             ' Kategoriebez. zum zusammenfassen
    bExpert As Boolean
    szList As String
End Type

Public Function InitOptionsForm(f As Form, Optional bExpert As Boolean)
    Dim i As Integer                                                ' Options Counter
    Dim c As Integer                                                ' CTL Counter
    Dim szCaption As String
    Dim szDetails As String                                         ' Zusätzliche Fehlerinformationen
    Dim x As Integer                                                ' Kategorie Counter
    Dim szKategorieList As String                                   ' ; getrente Liste aller Options Kategorien
    Dim szOptionsList As String                                     ' ; Getrente Liste aller Optionen
    Dim szKatArray() As String                                      ' Kategorien liste als Array
    Dim szOptArray() As String                                      ' Optionsliste als array
    Dim Opt As OptionValue
    Dim optionsindex As Integer
    Dim TopPos As Integer
    Dim lastVisFrameindex As Integer
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    f.bInit = True                                                  ' Zur zeit wird initialisiert
    c = 0                                                           ' Control Counter initialisieren
    f.MaxDisplayPathLen = 180
    szKategorieList = objTools.GetOptKategorieListFromXML(szOptionIni) ' Liste aller kategorien ermitteln
    szKatArray = Split(szKategorieList, ";")                        ' Kategorie liste in array
    f.FrameKategorie(0).Top = 10                                    ' Kategorie (frame) TopPos festlegen
    For x = 0 To UBound(szKatArray)                                 ' Alle kategorien durchgehen
        If x > 0 Then Load f.FrameKategorie(x)                      ' Pro Kategorie einen Frame laden
        f.FrameKategorie(x).Caption = szKatArray(x)                 ' Frame Caption setzen
        szOptionsList = objTools.GetOptionsNodeList(szOptionIni, szKatArray(x))       ' Liste aller optionen holen
        szOptArray = Split(szOptionsList, ";")                      ' optionen liste in array
        TopPos = 300
        For i = 0 To UBound(szOptArray)                             ' Alle optionen durchgehen
            optionsindex = OptionGetIndexByName(szOptArray(i))      ' index der Option ermitteln
            Opt = OptionGet(optionsindex)                           ' Options Type holen
            With Opt
                If .bEdit Or (.bExpert And bExpert) Then             ' Nur bEdit = true anzeigen
                    f.FrameKategorie(x).Visible = True              ' Frame anzeigen
                    lastVisFrameindex = x
                    If x > 0 Then f.FrameKategorie(x).Top = f.FrameKategorie(x - 1).Top _
                        + f.FrameKategorie(x - 1).Height + 100      ' und positionieren
                    If c > 0 Then                                   ' Neue Controls laden
                        Load f.CheckOption(c)                       ' Neue Chekbox laden
                        Load f.lblOption(c)                         ' Neues Lable laden
                        Load f.txtOption(c)                         ' Neue Textbox laden
                        Load f.cmdPath(c)                           ' Neuen Pfadauswahl Button laden
                        Load f.cmdFile(c)                           ' Neuen Fileauswahl Button laden
                        Load f.cmbOption(c)                         ' Neue ComboBox laden
                    End If
                    Set f.CheckOption(c).Container = f.FrameKategorie(x) ' als Container entsprechenden Frame festlegen
                    Set f.lblOption(c).Container = f.FrameKategorie(x)
                    Set f.txtOption(c).Container = f.FrameKategorie(x)
                    Set f.cmdPath(c).Container = f.FrameKategorie(x)
                    Set f.cmdFile(c).Container = f.FrameKategorie(x)
                    Set f.cmbOption(c).Container = f.FrameKategorie(x)
                    f.CheckOption(c).Top = TopPos                   ' Top Pos festlegen
                    f.lblOption(c).Top = TopPos
                    f.txtOption(c).Top = TopPos
                    f.cmdPath(c).Top = TopPos
                    f.cmdFile(c).Top = TopPos
                    f.cmbOption(c).Top = TopPos
                    f.CheckOption(c).Tag = ""                       ' Tags initialisieren
                    f.lblOption(c).Tag = ""
                    f.txtOption(c).Tag = ""
                    f.cmdPath(c).Tag = ""
                    f.cmdFile(c).Tag = ""
                    f.cmbOption(c).Tag = ""
                    f.FrameKategorie(x).Height = f.lblOption(c).Top + f.lblOption(c).Height + 100 ' Frame höhe festlegen
                    f.lblOption(c).Caption = .Caption               ' Caption in lable setzen
                    f.CheckOption(c).Caption = .Caption             ' Caption in Checkbox setzen
                    If .bBool Then                                  ' Wenn Boolwert
                        If CBool(.Value) Then                       ' Wert setzen
                            f.CheckOption(c).Value = 1
                        Else
                            f.CheckOption(c).Value = 0
                        End If
                        f.CheckOption(c).Tag = .Name                ' Name als Tag übergeben (fürs Speichern)
                        f.CheckOption(c).Visible = True             ' Optionsfeld sichtbar
                        f.CheckOption(c).Enabled = Not (.bDisabled) ' Evtl. diablen
                        f.lblOption(c).Visible = False              ' Label ausblenden
                        f.txtOption(c).Visible = False              ' Textbox ausblenden
                        f.cmdPath(c).Visible = False                ' Button Pfadauswahl ausblenden
                        f.cmdFile(c).Visible = False                ' Button Fileauswahl ausblenden
                        f.cmbOption(c).Visible = False              ' Combo box ausblenden
                    ElseIf .szList <> "" Then                       ' Werteliste
                        Call FillOptionCMBlist(f.cmbOption(c), .szList) ' Comboliste füllen
                        f.cmbOption(c).Text = CStr(.Value)          ' Wert setzen
                        f.cmbOption(c).Tag = .Name                  ' Name als Tag übergeben (fürs Speichern)
                        f.CheckOption(c).Visible = False            ' Optionsfeld ausblenden
                        f.cmbOption(c).Visible = True               ' Combobox einblenden
                        f.lblOption(c).Visible = True               ' Label sichtbar
                        f.txtOption(c).Visible = False              ' Textbox ausblenden
                    Else                                            ' Sonst Textbox
                        f.txtOption(c).Text = CStr(.Value)          ' Wert setzen
                        f.txtOption(c).Tag = .Name                  ' Name als Tag übergeben (fürs Speichern)
                        f.CheckOption(c).Visible = False            ' Optionsfeld ausblenden
                        f.cmbOption(c).Visible = False              ' Combobox ausblenden
                        f.lblOption(c).Visible = True               ' Label sichtbar
                        f.txtOption(c).Visible = True               ' Textbox sichtbar
                        f.txtOption(c).Enabled = Not (.bDisabled)   ' Evtl. diablen
                        If .bPath Then                              ' Option ist Verzeichnis
                            f.cmdPath(c).Visible = .bPath           ' Button Pfadauswahl sichtbar
                            f.cmdPath(c).Enabled = Not (.bDisabled) ' Evtl. diablen
                            f.cmdPath(c).Tag = f.txtOption(c).Text  ' text im Button Tag speichern
                            f.txtOption(c).Text = objTools.GetShortPath(f, CStr(f.cmdPath(c).Tag), _
                                    f.MaxDisplayPathLen)            ' Pfad anzeige kürzen
                        ElseIf .bFile Then                          ' Option ist Datei
                            f.cmdFile(c).Visible = .bFile           ' Button Fileauswahl sichtbar
                            f.cmdFile(c).Enabled = Not (.bDisabled) ' Evtl. diablen
                            f.cmdFile(c).Tag = f.txtOption(c).Text  ' text im Button Tag speichern
                            f.txtOption(c).Text = objTools.GetShortPath(f, CStr(f.cmdFile(c).Tag), _
                                    f.MaxDisplayPathLen)            ' Pfad anzeige kürzen
                        End If
                    End If
                    TopPos = f.lblOption(c).Top + f.lblOption(c).Height + 50 ' Neue Toppos für nächstes control bestimmen
                    c = c + 1                                       ' Control Counter hochzählen
                End If
            End With
        Next i                                                      ' Nächste optionen
    Next x                                                          ' Nächste Kategorie
    f.OKButton.Top = f.FrameKategorie(lastVisFrameindex).Top + _
            f.FrameKategorie(lastVisFrameindex).Height + 100        ' Buttons Positionieren
    f.cmdUpdate.Top = f.OKButton.Top
    f.CancelButton.Top = f.OKButton.Top
    f.Height = f.OKButton.Top + f.OKButton.Height + 100 + 405        ' Form vergößern

exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
    f.bInit = False                                                 ' Init flag setzen
    Err.Clear                                                       ' Evtl. Error clearen
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "InitOptionsForm", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Private Function FillOptionCMBlist(cmbCTL As ComboBox, szValueList As String) As Boolean
    Dim szListArray() As String                                     ' Array mit werten
    Dim i As Integer                                                ' Counter (Array)
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If cmbCTL Is Nothing Then GoTo exithandler                      ' Keine ComboBox -> Fertig
    cmbCTL.Clear                                                    ' Alte Werte raus
    If szValueList = "" Then GoTo exithandler                       ' Keine Werteliste -> Fertig
    szListArray = Split(szValueList, ";")                           ' Liste in Array aufsplaten
    If objTools.CheckArray(szListArray) Then                        ' Array nicht leer
        For i = 0 To UBound(szListArray)                            ' Alle Array Items durchlaufen
            If szListArray(i) <> "" Then cmbCTL.AddItem szListArray(i) ' Wert hinzufügen
        Next i                                                      ' Nächstes Array item
    End If
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr                                                     ' Evtl. Error clearen
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "FillOptionCMBlist", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function PrepareOptionValue(NewVal As OptionValue)
' Initialisiert Optios Wert
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    With NewVal
        .Name = ""
        .Value = ""
        .Caption = ""
        .bBool = False
        .bCrypt = False
        .bEdit = False
        .bDisabled = False
        .bPath = False
        .bFile = False
        .Kategorie = ""
    End With
    Err.Clear                                                       ' Evtl. error clearen
End Function

Public Function InitOptionValue(szValueName As String, _
        szValueCaption As String, _
        szDefaultValue As String, _
        szKategorie As String, _
        Optional bCrypt As Boolean, _
        Optional bBool As Boolean, _
        Optional bEdit As Boolean, _
        Optional bPath As Boolean, _
        Optional bFile As Boolean, _
        Optional bDisabled As Boolean, _
        Optional bExpert As Boolean, _
        Optional szList As String) As String

    Dim szValue As String
On Error GoTo Errorhandler
    If OptionGetIndexByName(szValueName) < 0 Then
        If CheckOptionArray(VarOptions) Then
            ReDim Preserve VarOptions(UBound(VarOptions) + 1)
        Else
            ReDim VarOptions(0)
        End If
        VarOptions(UBound(VarOptions)).Name = szValueName
        VarOptions(UBound(VarOptions)).Kategorie = szKategorie
        VarOptions(UBound(VarOptions)).Caption = szValueCaption
        VarOptions(UBound(VarOptions)).bCrypt = bCrypt
        VarOptions(UBound(VarOptions)).bBool = bBool
        VarOptions(UBound(VarOptions)).bEdit = bEdit
        VarOptions(UBound(VarOptions)).bDisabled = bDisabled
        VarOptions(UBound(VarOptions)).bPath = bPath
        VarOptions(UBound(VarOptions)).bFile = bFile
        VarOptions(UBound(VarOptions)).bExpert = bExpert
        VarOptions(UBound(VarOptions)).szList = szList
        VarOptions(UBound(VarOptions)).Value = szDefaultValue
    Else
        Call OptionSetByName(szValueName, szDefaultValue)
    End If
'    If bCrypt Then
'        szValue = objTools.Crypt(szValue, False)    ' Entschlüsseln
'    End If
    'Call VarOptions(UBound(VarOptions)).SetValue(szValue)
exithandler:
On Error Resume Next
    InitOptionValue = szValue
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "InitOptionValue", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function OptionGet(Index As Integer) As OptionValue
' Liefert Optiontyp nach index zurück
    OptionGet = VarOptions(Index)
End Function

Public Function OptionGetCount()
' Liefert anzsah der Optionen zurück
On Error Resume Next                                                ' Hier keine Fehlerbehandlung
    If CheckOptionArray(VarOptions) Then                            ' Array Prüfen
        OptionGetCount = UBound(VarOptions)                         ' Array Count zurück
    Else                                                            ' Sonst (Array Leer)
        OptionGetCount = -1                                         ' Neg. wert zurück
    End If
    Err.Clear                                                       ' Evtl. errror Clearen
End Function

Public Function OptionGetCryptByName(szOptionName As String) As Boolean
' Gibt zurück ob option Verschüsselt ist. Option wird nach name ausgewählt.
    Dim i As Integer                                                ' Counter
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If Not CheckOptionArray(VarOptions) Then GoTo exithandler       ' Array Leer -> Raus
    For i = 0 To UBound(VarOptions)                                 ' Alle optionen durchlaufen
        If UCase(VarOptions(i).Name) = UCase(szOptionName) Then     ' Wenn gesuchter Optionname gefunden
            OptionGetCryptByName = VarOptions(i).bCrypt             ' Wert auslesen
            Exit For                                                ' fertig
        End If
    Next i                                                          ' Nächste Option
exithandler:
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "OptionGetCryptByName", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function OptionGetByName(szOptionName As String) As Variant
' Liest optionswert aus. Option wird nach name ausgewählt.
    Dim i As Integer                                                ' Counter
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If Not CheckOptionArray(VarOptions) Then GoTo exithandler       ' Array Leer -> Raus
    For i = 0 To UBound(VarOptions)                                 ' Alle optionen durchlaufen
        If UCase(VarOptions(i).Name) = UCase(szOptionName) Then     ' Wenn gesuchter Optionname gefunden
            OptionGetByName = VarOptions(i).Value                   ' Wert auslesen
            Exit For                                                ' fertig
        End If
    Next i                                                          ' Nächste Option
exithandler:
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "OptionGetByName", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function OptionGetIndexByName(szOptionName As String) As Variant
' Liest optionswert aus. Option wird nach name ausgewählt.
    Dim i As Integer                                                ' Counter
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    OptionGetIndexByName = -1
    If Not CheckOptionArray(VarOptions) Then GoTo exithandler       ' Array Leer -> Raus
    For i = 0 To UBound(VarOptions)                                 ' Alle optionen durchlaufen
        If UCase(VarOptions(i).Name) = UCase(szOptionName) Then     ' Wenn gesuchter Optionname gefunden
            OptionGetIndexByName = i                                ' Index auslesen
            Exit For                                                ' Fertig
        End If
    Next i                                                          ' Nächste Option
exithandler:
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "OptionGetNameByIndex", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function OptionGetNameByIndex(Index As Integer) As String
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If Not CheckOptionArray(VarOptions) Then GoTo exithandler       ' Array Leer -> Raus
    OptionGetNameByIndex = VarOptions(Index).Name
exithandler:
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "OptionGetNameByIndex", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function OptionGetCyptByIndex(Index As Integer) As Boolean
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If Not CheckOptionArray(VarOptions) Then GoTo exithandler       ' Array Leer -> Raus
    OptionGetCyptByIndex = VarOptions(Index).bCrypt
exithandler:
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "OptionGetCyptByIndex", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function OptionSetByCaption(szOptionCaption As String, Value As Variant)
' Setzt Option Wert. Option wird nach Caption ausgewählt
    Dim i As Integer                                                ' Counter
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If Not CheckOptionArray(VarOptions) Then GoTo exithandler       ' Array Leer -> Raus
    For i = 0 To UBound(VarOptions)                                 ' Alle optionen durchlaufen
        If UCase(VarOptions(i).Caption) = UCase(szOptionCaption) Then  ' Wenn gesuchter Optioncaption gefunden
            VarOptions(i).Value = Value                             ' Wert setzen
            Exit For                                                ' fertig
        End If
    Next i                                                          ' Nächste option
exithandler:
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "OptionSetByCaption", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function OptionSetByName(szOptionName As String, Value As Variant)
' Setzt Option Wert. Option wird nach Name ausgewählt
    Dim i As Integer                                                ' Counter
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If Not CheckOptionArray(VarOptions) Then GoTo exithandler       ' Array Leer -> Raus
    For i = 0 To UBound(VarOptions)                                 ' Alle optionen durchlaufen
        If UCase(VarOptions(i).Name) = UCase(szOptionName) Then     ' Wenn gesuchter Optionname gefunden
            VarOptions(i).Value = Value                             ' Wert setzen
            Exit For                                                ' fertig
        End If
    Next i                                                          ' Nächste option
exithandler:
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "OptionSetByName", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function CheckOptionArray(VarArray() As OptionValue) As Boolean
' prüft ob das Option Array schon einen Eintrag hat
' zur Fehlervermeidung
    Dim i As Integer                                                ' Array größe
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    i = UBound(VarArray)                                            ' Array Größe ermitteln
    If Err.Number <> 0 Then                                         ' Ist ein Fehler aufgetreten
        Err.Clear                                                   ' Err Clearen
        CheckOptionArray = False                                    ' Array ist leer
        Exit Function                                               ' Fertig
    End If
    CheckOptionArray = True                                         ' Sonst Array Definiert
    Err.Clear                                                       ' Evtl. error clearen
End Function




