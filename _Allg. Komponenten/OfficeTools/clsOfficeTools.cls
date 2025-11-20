VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOfficeTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const MODULNAME = "clsOfficeTools"                          ' Modulname für Fehlerbehandlung

Private objObjectBag As Object                                      ' ObjectBag object
Private objError As Object                                          ' Error Object
Private objWord As Word.Application                                 ' Word Application Object
Private objExcel As excel.Application                               ' Ecxel Application Object
Private OfficeVersionen As OfficeInfo

Private bXMLTemplatesPosible As Boolean                             ' XML Vorlagen sind zulässig

Private Sub Class_Initialize()
'
    
End Sub

Private Sub Class_Terminate()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Set objError = Nothing
    Set objObjectBag = Nothing
    Set objWord = Nothing
    Set objExcel = Nothing
    Err.Clear                                                       ' Evtl. error clearen
End Sub

Public Function InitObjectBag(objOb As Object)
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    Set objObjectBag = objOb
    If Not objObjectBag Is Nothing Then
        Set objError = objObjectBag.GetErrorObj
        Call InitOfficeTools
    End If
    Err.Clear                                                       ' Evtl. error clearen
End Function

Public Function InitOfficeTools()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    With OfficeVersionen
        .WordVersion = objObjectBag.getWordVersion()                ' Word version aus Obag ermitteln
        If .WordVersion >= "11" Then bXMLTemplatesPosible = True    ' Word >=11 (2003) XML Vorlagen zulassen
        .WordPath = objObjectBag.GetWordPath()
        .ExcelVersion = objObjectBag.GetExcelVersion()              ' Excel version aus Obag ermitteln
        .ExcelPath = objObjectBag.GetExcelVersion()
    End With
    Err.Clear                                                       ' Evtl. error clearen
End Function
                                                                    ' *****************************************
                                                                    ' Word Methoden
Public Function InitWordObj() As Boolean
On Error Resume Next                                                ' Hier erstmal keine Fehler
    Set objWord = GetObject(, "Word.Application")                   ' Versuchen laufende Word Version zu ermitteln
    If Err.Number <> 0 Then                                         ' Sind wir auf einen Fehler gelaufen ?
        Err.Clear                                                   ' Dann Err Clearen
        Set objWord = Nothing                                       ' Obj Killen
    End If
On Error GoTo Errorhandler                                          ' Fehlerbehandlung wieder aktivieren
    If objWord Is Nothing Then                                      ' Wenn Word nicht ausgeführt wird, Word starten:
        Set objWord = CreateObject("Word.Application")              ' Neues Word Application object initialisieren
    End If
    InitWordObj = True                                              ' Erfolg zurück
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
    Err.Clear                                                       ' Evtl. error clearen
Exit Function                                                            ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "InitWordObj", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function CloseWordObj()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    objWord.Application.Quit                                        ' Word Schliessen
    Set objWord = Nothing                                           ' Obj Killen
    Err.Clear                                                       ' Fehler Clearen
End Function

Public Function BringWordOnTop()
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    If objWord Is Nothing Then GoTo exithandler                     ' Kein Word Obj. -> Fertig
    objWord.Application.Activate                                    ' Word in Vordergrung
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
    Err.Clear                                                       ' Evtl. error clearen
End Function

Public Function ShowWord(objDoc As Object, bVisible As Boolean)
On Error Resume Next                                                ' Fehlerbehandlung deaktivieren
    If objWord Is Nothing Then GoTo exithandler                     ' Kein Word Obj. -> Fertig
    objWord.Visible = bVisible                                      ' Word Application anzeigen
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
    Err.Clear                                                       ' Evtl. error clearen
End Function

Public Function OpenNewWordDoc(szVorlagePath As String, Optional bHide As Boolean) As Object
    Dim objDoc As Word.Document                                     ' Document Object
    Dim szDetails As String                                         ' Details für Fehlerbehandlung
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If Not InitWordObj Then GoTo exithandler                        ' Word Application object initialisieren
    szDetails = "Vorlage: " & szVorlagePath                         ' Details für Fehlermeldung            '
    objWord.Visible = Not bHide                                     ' Anzeigen ?
    Set objDoc = objWord.Documents.Open(szVorlagePath)              ' Vorlage öffnen
    Set OpenNewWordDoc = objDoc                                     ' Doc Object zurückgeben
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
    Err.Clear                                                       ' Evtl. error clearen
Exit Function                                                       ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "OpenNewWordDoc", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function SaveWordDoc(objDoc As Object, szSavePath As String) As Boolean
    Dim szDetails As String                                         ' Details für Fehlerbehandlung
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If szSavePath = "" Then                                         ' Wenn SavePath Leer
        szDetails = "Kein Speicherpfad angegeben."                  ' Details für Fehlermeldung
        GoTo exithandler                                            ' Fertig
    End If
    szDetails = szSavePath                                          ' Details für Fehlermeldung
    objDoc.SaveAs szSavePath                                        ' Word Dokument speichern
    SaveWordDoc = True                                              ' Erfolg zurück
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
    Err.Clear                                                       ' Evtl. error clearen
Exit Function                                                       ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "SaveWordDoc", errNr, errDesc, szDetails) ' Fehler behandlung aufrufen
    SaveWordDoc = False                                             ' Misserfolg zurück
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function SetDataInDoc(objWDoc As Object, _
        rsDoc As Object, _
        Optional ProgBarTitle As String, _
        Optional ProgbarAktion As String) As Boolean
    Dim i As Integer                                                ' Counter DocVariablen
    Dim x As Integer                                                ' Counter RS felder
    Dim tmpValue As String                                          ' Variablen Wert
    Dim szDocVarName As String                                      ' Name der DocVariable
    Dim szFieldName As String                                       ' Feldname des RS
    Dim lngDSNum As Integer                                         ' DS nummer
    Dim objDoc As Word.Document
    Dim MaxSteps As Integer
    Dim Step As Integer
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    Set objDoc = objWDoc
    If rsDoc Is Nothing Then                                        ' Wenn Kein RS
        GoTo Cancelhandler                                          ' -> fertig
    End If
    If rsDoc.RecordCount = 0 Then                                   ' Wenn Kein DS
        GoTo Cancelhandler                                          ' -> Ferig
    End If
    If ProgBarTitle <> "" Then                                      ' Titel für ablauf anzeige
        MaxSteps = objObjectBag.GetMaxProgresSteps(rsDoc.RecordCount)
        Step = objObjectBag.GetProgresStepsWidth(rsDoc.RecordCount, MaxSteps)
        Call objObjectBag.ShowMSGForm(True, ProgBarTitle, MaxSteps, ProgbarAktion)
    End If
    If rsDoc.RecordCount >= 1 Then                                  ' Wenn mehr als ein DS
        For i = 1 To objDoc.Fields.Count                            ' Alle Dokumenten Felder durchlaufen
            szDocVarName = GetDocVarFieldname(objDoc, i)            ' DocVariablen Namen holen
            If szDocVarName <> "" Then                              ' Wenn DocVargefunden
                If IsNumeric(Right(szDocVarName, 1)) Then           ' Einstellige Ziffer am VarNamen
                    szFieldName = Left(szDocVarName, Len(szDocVarName) - 1) ' Feldnamen holen
                    lngDSNum = CLng(Right(szDocVarName, 1))         ' DS Nummer aus dieser Zahl
                    lngDSNum = lngDSNum - 1
                ElseIf IsNumeric(Right(szDocVarName, 2)) Then       ' Zweistellige Ziffer am VarNamen
                    szFieldName = Left(szDocVarName, Len(szDocVarName) - 2) ' Feldnamen holen
                    lngDSNum = CLng(Right(szDocVarName, 2))         ' DS Nummer aus dieser Zahl
                    lngDSNum = lngDSNum - 1
                End If
                If szFieldName <> "" And lngDSNum > -1 Then         ' Wenn Feldname vorhanden
                    If lngDSNum < rsDoc.RecordCount Then                ' DS Nummer  innerhalb des RS
                        rsDoc.movefirst
                        rsDoc.Move lngDSNum                         ' Zum DS springen
On Error Resume Next                                                ' Fehlerbeh. deakt. da feld evtl. nicht vorhanden
                        tmpValue = rsDoc.Fields(szFieldName).Name
                        If tmpValue <> "" Then
                            tmpValue = CStr(rsDoc.Fields(szFieldName).Value) ' Wert holen
                            If tmpValue = "" Then tmpValue = " "
                            objDoc.Variables.Add szDocVarName, tmpValue ' Doc Variable hinzufügen
                            objDoc.Variables(szDocVarName).Value = tmpValue ' Wert setzen
                        End If
                        Err.Clear
On Error GoTo Errorhandler                                          ' fehlerbeh. wieder aktiv.
                    End If
                End If
            End If
            tmpValue = ""
            szDocVarName = ""
            szFieldName = ""
            lngDSNum = 0
        Next i                                                      ' Nächste Doc Variable
    End If
    rsDoc.movefirst                                                 ' Zum 1. DS Springen
    For i = 1 To objDoc.Fields.Count                                ' Alle Dokumenten Felder durchlaufen
        szDocVarName = GetDocVarFieldname(objDoc, i)                ' DocVariablen Namen holen
On Error Resume Next                                                ' Fehlerbeh. deakt. da feld evtl. nicht vorhanden
        tmpValue = rsDoc.Fields(szDocVarName).Name
        If tmpValue <> "" Then
            tmpValue = CStr(rsDoc.Fields(szDocVarName).Value)       ' Wert holen
            If tmpValue = "" Then tmpValue = " "
                objDoc.Variables.Add szDocVarName, tmpValue         ' Doc Variable hinzufügen
                objDoc.Variables(szDocVarName).Value = tmpValue     ' Wert setzen
                Call objObjectBag.ShowMsgNextStep(ProgbarAktion, Step)  ' Progbar
                Err.Clear                                           ' Evtl. Error Clearen
On Error GoTo Errorhandler                                          ' Fehlerbehandlung wieder Aktivieren
        End If
        tmpValue = ""
        szDocVarName = ""
    Next i                                                          ' Nächste Doc Variable
    For i = 1 To objDoc.Fields.Count                                ' Feld funktionen im Dokument aktualisieren
        objDoc.Fields(i).Update
    Next i
    objDoc.ActiveWindow.View.ShowFieldCodes = False                 ' Doc variablen ausbelnden
    SetDataInDoc = True                                             ' Erfolg zurück
    
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
    Err.Clear                                                       ' Evtl. error clearen
Exit Function                                                       ' Function Beenden
Cancelhandler:
    SetDataInDoc = False                                            ' Abbruch zurück
    GoTo exithandler                                                ' Weiter mit Exithandler
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "SetDataInDoc", errNr, errDesc) ' Fehler behandlung aufrufen
    SetDataInDoc = False                                            ' Misserfolg zurück
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function SetDataInDocNeu(objWDoc As Object, _
        rsDoc As Object, _
        Optional bListe As Boolean, _
        Optional ProgBarTitle As String, _
        Optional ProgbarAktion As String) As Boolean
    Dim i As Integer                                                ' Counter DocVariablen
    Dim x As Integer                                                ' Counter RS felder
    Dim tmpValue As String                                          ' Variablen Wert
    Dim szDocVarName As String                                      ' Name der DocVariable
    Dim szFieldName As String                                       ' Feldname des RS
    Dim lngDSNum As Integer                                         ' DS nummer
    Dim objDoc As Word.Document                                     ' Akt Word Dokument
    Dim MaxSteps As Integer
    Dim Step As Integer
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    Set objDoc = objWDoc                                            ' Übergebenes Word Dokument
    If rsDoc Is Nothing Then                                        ' Wenn Kein RS
        GoTo Cancelhandler                                          ' -> fertig
    End If
    If rsDoc.RecordCount = 0 Then                                   ' Wenn Kein DS
        GoTo Cancelhandler                                          ' -> Ferig
    End If
    If bListe Then                                                  ' Wenn mehr als ein DS
        If ProgBarTitle <> "" Then                                  ' gibts einene Meldungstitel
            MaxSteps = objObjectBag.GetMaxProgresSteps(rsDoc.RecordCount * rsDoc.Fields.Count)
            Step = objObjectBag.GetProgresStepsWidth(rsDoc.RecordCount, MaxSteps)
            Call objObjectBag.ShowMSGForm(True, ProgBarTitle, MaxSteps, ProgbarAktion) ' Meldung anzeigen /Aktualisieren
        End If
        rsDoc.movefirst                                             ' Zum 1. DS Springen
        x = 1
        While Not rsDoc.EOF                                         ' Solange RS nicht Leer
            Call objObjectBag.ShowMsgNextStep(ProgbarAktion, Step)  ' Progbar
            For i = 0 To rsDoc.Fields.Count - 1                     ' Alle Felder duchlaufen
                szDocVarName = rsDoc.Fields(i).Name
                szDocVarName = szDocVarName & CStr(x)
            On Error Resume Next                                    ' Fehlerbehandlung deaktivieren
                tmpValue = rsDoc.Fields(i).Value                    ' Feldwert holen
                Err.Clear                                           ' Evtl. error clearen
            On Error GoTo Errorhandler                              ' Fehlerbehandlung aktivieren
                If tmpValue = "" Then tmpValue = " "                ' Leerstring fürt zu fehler
            On Error Resume Next                                    ' Fehlerbehandlung deaktivieren
                objDoc.Variables(szDocVarName).Value = tmpValue
                If Err.Number <> 0 Then                             ' Wenn kein Fehler
                    objDoc.Variables.Add szDocVarName, tmpValue     ' Doc Variable hinzufügen
                End If
                Err.Clear                                           ' Evtl. error clearen
            On Error GoTo Errorhandler                              ' Fehlerbehandlung aktivieren
                szDocVarName = ""
                tmpValue = ""
            Next i                                                  ' Nächstes Feld
            x = x + 1
            rsDoc.movenext
        Wend
    Else
        If ProgBarTitle <> "" Then                                  ' Gibts einen Meldungstitel
            MaxSteps = objObjectBag.GetMaxProgresSteps(rsDoc.Fields.Count)
            Step = objObjectBag.GetProgresStepsWidth(rsDoc.Fields.Count, MaxSteps)
            Call objObjectBag.ShowMSGForm(True, ProgBarTitle, MaxSteps, ProgbarAktion) ' Meldung anzeigen /Aktualisieren
        End If
        rsDoc.movefirst                                             ' Zum 1. DS Springen
        For i = 0 To rsDoc.Fields.Count - 1                         ' Alle Felder Duchlaufen
            Call objObjectBag.ShowMsgNextStep(ProgbarAktion, Step)  ' Progbar
            szDocVarName = rsDoc.Fields(i).Name
        On Error Resume Next                                        ' Fehlerbehandlung deaktivieren
            tmpValue = rsDoc.Fields(i).Value                        ' Feldwert holen
            Err.Clear                                               ' Evtl. error clearen
        On Error GoTo Errorhandler                                  ' Fehlerbehandlung aktivieren
            If tmpValue = "" Then tmpValue = " "
        On Error Resume Next                                        ' Fehlerbehandlung deaktivieren
            If objDoc.Variables(szDocVarName).Name = szDocVarName Then
                objDoc.Variables(szDocVarName).Value = tmpValue
            End If
            If Err.Number <> 0 Then                                 ' Wenn Fehler
                objDoc.Variables.Add szDocVarName, tmpValue         ' Doc Variable hinzufügen
            End If
            Err.Clear                                               ' Evtl. error clearen
        On Error GoTo Errorhandler                                  ' Fehlerbehandlung aktivieren
            szDocVarName = ""
            tmpValue = ""
        Next i
    End If

    For i = 1 To objDoc.Fields.Count                                ' Feld funktionen im Dokument aktualisieren
        objDoc.Fields(i).Update
    Next i                                                          ' Nächstes Feld
    objDoc.ActiveWindow.View.ShowFieldCodes = False                 ' Doc variablen ausbelnden
    SetDataInDocNeu = True                                          ' Erfolg zurück
    
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
    Call objObjectBag.ShowMSGForm(False, "")                        ' Meldungs Fenster ausblenden
    Err.Clear                                                       ' Evtl. error clearen
Exit Function                                                       ' Function Beenden
Cancelhandler:
    SetDataInDocNeu = False                                         ' Abbruch zurück
    GoTo exithandler                                                ' Weiter mit Exithandler
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "SetDataInDocNeu", errNr, errDesc) ' Fehler behandlung aufrufen
    SetDataInDocNeu = False                                         ' Misserfolg zurück
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function DocHasVarField(objWDoc As Object, _
        DocVarName As String, _
        Optional bBeginWith As Boolean) As Boolean
' Prüft das vorkommen von Dokumentvariablen
    Dim objDoc As Word.Document                                     ' Word Doc obj
    Dim i As Integer                                                ' Counter
    Dim szDocVarName As String                                      ' Akt DocVariablen Name
    Dim lngFoundPos As Integer                                      ' Pos vorn vorkommen DocVaName in szDocVarname
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    Set objDoc = objWDoc
    For i = 1 To objDoc.Fields.Count                                ' Alle Dokumenten Felder durchlaufen
        szDocVarName = GetDocVarFieldname(objDoc, i)                ' Name der DocVariable aus feld ermitteln
        If szDocVarName <> "" Then                                  ' Wenn DocVariable vorhanden
            If bBeginWith Then                                      ' Nur anfang oder
                If InStr(UCase(szDocVarName), UCase(DocVarName)) = 1 _
                        And Len(szDocVarName) > Len(DocVarName) Then ' vergleichen
                    DocHasVarField = True                           ' Gefunden
                    Exit For
                End If
            Else                                                    ' Komplett gleich
                If UCase(szDocVarName) = UCase(DocVarName) Then     ' vergleichen
                    DocHasVarField = True                           ' gefunden
                    Exit For
                End If
            End If
        End If
        szDocVarName = ""
    Next i                                                          ' Nächstes feld
    
exithandler:
On Error Resume Next                                                ' Hier keine Fehler mehr
    Err.Clear                                                       ' Evtl. error clearen
Exit Function                                                       ' Function Beenden
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    Call objError.Errorhandler(MODULNAME, "DocHasVarField", errNr, errDesc) ' Fehler behandlung aufrufen
    DocHasVarField = False                                          ' Misserfolg zurück
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function DocHasVarFields(objWDoc As Object, _
        szDocVarNames As String, _
        Optional bBeginWith As Boolean) As String
' Prüft das vorkommen von Dokumentvariablen
    Dim objDoc As Word.Document                                     ' Word Doc obj
    Dim VarNamesArray() As String                                   ' szDocVarNames in Array
    Dim i As Integer                                                ' Counter (Docvariablen)
    Dim n As Integer                                                ' Counter (VarNamesArray)
    Dim szDocVarName As String                                      ' Akt DocVariablen Name
    Dim lngFoundPos As Integer                                      ' Pos vorn vorkommen DocVaName in szDocVarname
    Dim FoundResult As String
    Dim bComplete As String
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If szDocVarNames = "" Then GoTo exithandler
    Set objDoc = objWDoc
    VarNamesArray = Split(szDocVarNames, ";")
    
    For i = 1 To objDoc.Fields.Count                                ' Alle Dokumenten Felder durchlaufen
        lngFoundPos = -1
        szDocVarName = GetDocVarFieldname(objDoc, i)                ' Name der DocVariable aus feld ermitteln
        If szDocVarName <> "" Then                                  ' Wenn DocVariable vorhanden
            If bBeginWith Then                                      ' Nur anfang oder
                For n = 0 To UBound(VarNamesArray)
                    If VarNamesArray(n) <> "" Then                  ' Nur wenn gesuchter DocVarname nicht Leer
                        If InStr(UCase(szDocVarName), UCase(VarNamesArray(n))) = 1 _
                                And Len(szDocVarName) >= Len(VarNamesArray(n)) Then ' vergleichen
                            FoundResult = FoundResult & VarNamesArray(n) & ";"
                            VarNamesArray(n) = ""
                        End If
                    End If
                Next n
            Else                                                    ' Komplett gleich
                For n = 0 To UBound(VarNamesArray) - 1
                    If VarNamesArray(n) <> "" Then                  ' Nur wenn gesuchter DocVarname nicht Leer
                         If UCase(szDocVarName) = UCase(VarNamesArray(n)) Then     ' vergleichen
                            FoundResult = FoundResult & VarNamesArray(n) & ";"
                            VarNamesArray(n) = ""
                         End If
                    End If
                Next n
            End If
        End If
        szDocVarName = ""
        For n = 0 To UBound(VarNamesArray) - 1
            If VarNamesArray(n) <> "" Then
                lngFoundPos = n
            End If
            
        Next n
        If lngFoundPos = -1 Then Exit For
    Next i                                                          ' Nächstes feld
    
    DocHasVarFields = FoundResult
exithandler:
        
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    DocHasVarFields = ""
    Call objError.Errorhandler(MODULNAME, "DocHasVarFields", errNr, errDesc)
    Resume exithandler
End Function

Private Function GetDocVarFieldname(objWDoc As Object, Fieldindex As Integer) As String

    Dim szFieldCode As String                                       ' Code zeile des feldes
    Dim szDocVarName As String                                      ' Name der Dokument Variablen
    Dim objDoc As Word.Document                                     ' Word Doc obj
    Dim EndPos As Integer                                           ' End Position der DocVariablennamens
    
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    
    EndPos = -1                                                     ' End Pos auf abbruch setzen
    Set objDoc = objWDoc                                            ' Word Doc obj. übergeben

On Error Resume Next                                                ' Fehlerbeh. deakt.

    szFieldCode = objDoc.Fields(Fieldindex).Code                    ' Feldcode holen
    Err.Clear                                                       ' Error reseten
    
On Error GoTo Errorhandler                                          ' Fehlerbeh. wieder aktiv.

    If szFieldCode <> "" Then                                       ' Field Code vorhanden
        szFieldCode = Trim(szFieldCode)
        If InStr(szFieldCode, "DOCVARIABLE") > 0 Then               ' Field ist eine Docvariable
            EndPos = InStr(szFieldCode, "\*")
            If EndPos > 0 Then
                szDocVarName = Mid(szFieldCode, 12, EndPos - 13)
                szDocVarName = Trim(szDocVarName)
            End If
        End If
    End If
    
    GetDocVarFieldname = szDocVarName
exithandler:

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetDocVarFieldname", errNr, errDesc)
    Resume exithandler
End Function

                                                                    ' *****************************************
                                                                    ' Excel Methoden
Public Function InitExcelObj()

On Error GoTo Errorhandler

    If objExcel Is Nothing Then
        Set objExcel = New excel.Application                        ' Excel Application object initialisieren
    End If
    
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "InitExcelObj", errNr, errDesc)
    Resume exithandler
End Function

Public Function ShowExcel(bVisible As Boolean)

    objExcel.Visible = bVisible                                     ' Excel Application anzeigen
    
End Function

Public Function OpenNewExcelWorkbook(szVorlagePath As String, _
                Optional bHide As Boolean) As Object

    Dim objWorkbook As excel.Workbook

On Error GoTo Errorhandler

    Call InitExcelObj                                               ' Excel Application object initialisieren
        
    objExcel.Visible = Not bHide                                    ' Anzeigen ?
    If szVorlagePath = "" Then
        Set objWorkbook = objExcel.Workbooks.Add
    Else
        Set objWorkbook = objExcel.Workbooks.Open(szVorlagePath)    ' Vorlage öffnen
    End If
    
    Set OpenNewExcelWorkbook = objWorkbook
        
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "OpenNewExcelWorkbook", errNr, errDesc)
    Resume exithandler
End Function

Public Function SaveExcelWorkbook(objWorkbook As Object, szSavePath As String) As Boolean

On Error GoTo Errorhandler

    objWorkbook.SaveAs szSavePath                                   ' Excel Workbook speichern
    SaveExcelWorkbook = True
     
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "SaveExcelWorkbook", errNr, errDesc)
    SaveExcelWorkbook = False
    Resume exithandler
End Function

Public Function GetExcelWorkSheet(objWorkbook As Object, Optional SheetIndex As Integer) As Object

On Error GoTo Errorhandler

    If SheetIndex = 0 Then SheetIndex = 1                           ' Index mind 1
    Set GetExcelWorkSheet = objWorkbook.Worksheets(SheetIndex)      ' Worksheet zurückgeben
    
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "GetExcelWorkSheet", errNr, errDesc)
    Resume exithandler
End Function

Public Function WriteExcelCell(objWorkSheet As Object, szRange As String, szValue As String) As Boolean

On Error GoTo Errorhandler

    If objWorkSheet Is Nothing Then GoTo exithandler                ' Kein Worksheet -> fertig
    
    objWorkSheet.Range(szRange).Value = szValue                     ' Wert in Range (Zelle) Schreiben
    'objWorkSheet.Range("A1").Value = "blablabla"
    WriteExcelCell = True                                           ' Erfolg zurück
    
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "WriteExcelCell", errNr, errDesc)
    Resume exithandler
End Function

Public Function SetExcelPrintSetup(objWorkSheet As Object, _
    szPrintArea As String, _
    szPrintTitleRows As String, _
    Optional bTitleBold As Boolean, _
    Optional bPagenubers As Boolean, _
    Optional szHeaderText As String, _
    Optional bPrintNow As Boolean, _
    Optional bPrintPreview As Boolean, _
    Optional bQuerformat As Boolean) As Boolean

On Error GoTo Errorhandler

    If objWorkSheet Is Nothing Then GoTo exithandler                ' Kein Worksheet -> fertig
    If bPagenubers Then
        Call SetExcelPageNumbers(objWorkSheet)                      ' Seitenzahlem mitausdrucken
    End If
        
    If szPrintArea <> "" Then                                       ' Wenn Druckbereich angegeben
        If SetExcelPrintArea(objWorkSheet, szPrintArea) = "" Then GoTo exithandler ' Druck bereich setzen
    End If
    
    If szHeaderText <> "" Then
        Call SetExcelHeaderText(objWorkSheet, szHeaderText, bTitleBold) ' Uberschrift (Kopzeile) setzen
    End If
    
    If bQuerformat Then
        Call SetExcelPageFormat(objWorkSheet, bQuerformat)          ' Format Setzen
    End If
    
    If szPrintTitleRows <> "" Then                                  ' Wenn Wiederholter Titel angegeben
        If SetExcelPrintTitel(objWorkSheet, szPrintTitleRows) = "" Then GoTo exithandler ' Setzen
        If bTitleBold Then
            Call SetExcelCellBold(objWorkSheet, szPrintTitleRows)   ' Titel Fett Setzen
        End If
    End If

    If bPrintPreview Then                                           ' In Seitenansicht anzeigen
        Call ExcelPrintPreview(objWorkSheet)                        ' Druckvorschau anzeigen
        'objWorkSheet.Saved = True
    Else
        If bPrintNow Then                                           ' Sofort Drucken ?
            If ExcelPrintOut(objWorkSheet) Then                     ' Wenn Ausdruck erfolgreich
                Call CloseExcelObj(True)                            ' Excel Schliessen
            End If
        End If
    End If
    SetExcelPrintSetup = True                                       ' Erfolg zurück
        
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "SetExcelPrintSetup", errNr, errDesc)
    Resume exithandler
End Function

Public Function ExcelPrintOut(objWorkSheet As Object) As Boolean

On Error GoTo Errorhandler

    If objWorkSheet Is Nothing Then GoTo exithandler                ' Kein Worksheet -> fertig
    
    objWorkSheet.PrintOut Copies:=1, Collate:=True                  ' Ausdrucken
    'ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
    ExcelPrintOut = True
    
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "ExcelPrintOut", errNr, errDesc)
    Resume exithandler
End Function

Public Function ExcelPrintPreview(objWorkSheet As Object) As Boolean

On Error GoTo Errorhandler

    If objWorkSheet Is Nothing Then GoTo exithandler                ' Kein Worksheet -> fertig
    Call ShowExcel(True)
    objWorkSheet.PrintPreview
    'ActiveWindow.SelectedSheets.PrintPreview
    ExcelPrintPreview = True
    
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "ExcelPrintPreview", errNr, errDesc)
    Resume exithandler
End Function

Public Function SetExcelPageFormat(objWorkSheet As Object, Optional bQuerformat As Boolean) As Boolean

On Error GoTo Errorhandler

    If objWorkSheet Is Nothing Then GoTo exithandler                ' Kein Worksheet -> fertig
    
    If bQuerformat Then
        objWorkSheet.PageSetup.Orientation = xlLandscape            ' Querformat
    Else
        objWorkSheet.PageSetup.Orientation = xlPortrait             ' Hochkant
    End If
    
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "SetExcelPageFormat", errNr, errDesc)
    Resume exithandler
End Function

Public Function SetExcelHeaderText(objWorkSheet As Object, szHeaderText As String, Optional bBold As Boolean) As Boolean

On Error GoTo Errorhandler

    If objWorkSheet Is Nothing Then GoTo exithandler                ' Kein Worksheet -> fertig
    
    If bBold Then
        '"&""Arial,Fett"
         objWorkSheet.PageSetup.LeftHeader = szHeaderText
         'objWorkSheet.PageSetup.LeftHeader = Chr(34) & "&" & Chr(34) & "Arial,Fett" & szHeaderText
    Else
        'With ActiveSheet.PageSetup.LeftHeader = "> Liste > Name"
        objWorkSheet.PageSetup.LeftHeader = szHeaderText                ' Kopzeilen Text Rechts setzen
    End If
    
    SetExcelHeaderText = True
    
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "SetExcelHeaderText", errNr, errDesc)
    Resume exithandler
End Function

Public Function SetExcelPageNumbers(objWorkSheet As Object) As Boolean

On Error GoTo Errorhandler

    If objWorkSheet Is Nothing Then GoTo exithandler                ' Kein Worksheet -> fertig

    objWorkSheet.PageSetup.CenterFooter = "Seite &P von &N"         ' Seitenzahlen setzen
     
    SetExcelPageNumbers = True                                      ' Erfolg zurück
     
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "SetExcelPageNumbers", errNr, errDesc)
    Resume exithandler
End Function

Public Function SetExcelCellBold(objWorkSheet As Object, szRange As String) As Boolean
    
On Error GoTo Errorhandler

    If objWorkSheet Is Nothing Then GoTo exithandler                ' Kein Worksheet -> fertig
    
    'Range("A1:B1").Select
    'Selection.Font.Bold = True
    objWorkSheet.Range(szRange).Font.Bold = True                    ' Range Selecten & Front = bold
    'objWorkSheet.Selection.Font.Bold = True                        ' Schrift Fett setzen
    
    SetExcelCellBold = True                                         ' Erfolg zurück
    
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "SetExcelCellBold", errNr, errDesc)
    Resume exithandler
End Function

Public Function SetExcelPrintArea(objWorkSheet As Object, szRange As String) As String

On Error GoTo Errorhandler

    If objWorkSheet Is Nothing Then GoTo exithandler                ' Kein Worksheet -> fertig

    'ActiveSheet.PageSetup.PrintArea = "$A$11:$J$47"
    objWorkSheet.PageSetup.PrintArea = szRange                      ' Druckbereich setzen

    SetExcelPrintArea = objWorkSheet.PageSetup.PrintArea

exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "SetExcelPrintArea", errNr, errDesc)
    Resume exithandler
End Function

Public Function SetExcelPrintTitel(objWorkSheet As Object, szRange As String) As String

On Error GoTo Errorhandler

    If objWorkSheet Is Nothing Then GoTo exithandler                ' Kein Worksheet -> fertig
    
    ' ActiveSheet.PageSetup.PrintTitleRows = "$1:$1"
    objWorkSheet.PageSetup.PrintTitleRows = szRange                 ' Spaltentitel für jede Seite setzen
    
    SetExcelPrintTitel = objWorkSheet.PageSetup.PrintTitleRows
    
exithandler:
On Error Resume Next

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "SetExcelPrintTitel", errNr, errDesc)
    Resume exithandler
End Function

Public Function CloseExcelObj(Optional bNoAskForSave As Boolean)

On Error Resume Next
    
    If bNoAskForSave Then                                           ' Nicht nachfragen ob Gespeichert werden soll
        objExcel.ActiveWorkbook.Saved = True                        ' Auf Gespeichert setzen
    End If
    objExcel.Application.Quit                                       ' Excel Schliessen
    Set objExcel = Nothing
    
exithandler:

Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "CloseExcelObj", errNr, errDesc)
    Resume exithandler
End Function
