Attribute VB_Name = "modTextTools"
Option Explicit                                                     ' Variaben Deklaration erzwingen
Const MODULNAME = "modTextTools"                                    ' Modulname für Fehlerbehandlung
Public objError As Object                                           ' Error Object

Public Function SetHTTPSyntax(str1 As String) As String
' Ersetzt Sonderzeichen im String für Browser Adressen
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    str1 = Replace(str1, "%", "%25")
    str1 = Replace(str1, vbCrLf, "%0A")
    str1 = Replace(str1, vbCr, "%0A")
    str1 = Replace(str1, " ", "%20")
    str1 = Replace(str1, "!", "%21")
    str1 = Replace(str1, "#", "%23")
    str1 = Replace(str1, "*", "%2A")
    str1 = Replace(str1, "/", "%2F")
    str1 = Replace(str1, "?", "%3F")
    str1 = Replace(str1, "Ä", "%C4")
    str1 = Replace(str1, "Ö", "%D6")
    str1 = Replace(str1, "Ü", "%DC")
    str1 = Replace(str1, "ß", "%DF")
    str1 = Replace(str1, "ä", "%E4")
    str1 = Replace(str1, "ö", "%F6")
    str1 = Replace(str1, "ü", "%FC")
    SetHTTPSyntax = str1
exithandler:
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "SetHTTPSyntax", errNr, errDesc)
End Function

Public Function SetXMLSyntax(str1 As String) As String
' Ersetzt Sonderzeichen im String für XML/HTML Text
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    str1 = Replace(str1, "&", "&amp;")                              ' Kaufmanns UND erstezen
    str1 = Replace(str1, "'", "&apos;")                             ' Hochkomma ersetzen
    str1 = Replace(str1, "<", "&lt;")                               ' Kleiner ersetzen
    str1 = Replace(str1, ">", "&gt;")                               ' Größer ersetzen
    str1 = Replace(str1, Chr(34), "&quot;")                         ' " erstezen
    str1 = Replace(str1, "Ä", "&#196;")                             ' Ä erstezen
    str1 = Replace(str1, "Ö", "&#214;")                             ' Ö erstezen
    str1 = Replace(str1, "Ü", "&#220;")                             ' Ü erstezen
    str1 = Replace(str1, "ä", "&#228;")                             ' ä erstezen
    str1 = Replace(str1, "ö", "&#246;")                             ' ö erstezen
    str1 = Replace(str1, "ü", "&#252;")                             ' ü erstezen
    str1 = Replace(str1, "ß", "&#223;")                             ' ß erstezen
    SetXMLSyntax = str1
exithandler:
Exit Function
Errorhandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.Errorhandler(MODULNAME, "SetXMLSyntax", errNr, errDesc)
End Function

Public Function TextTo1337(szText As String) As String
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    szText = Replace(szText, "a", "4")                              ' A = 4
    szText = Replace(szText, "A", "4")
    szText = Replace(szText, "b", "8")                              ' B = 8
    szText = Replace(szText, "B", "8")
    szText = Replace(szText, "C", "c")                              ' C = c
    szText = Replace(szText, "D", "d")                              ' D = d
    szText = Replace(szText, "E", "3")
    szText = Replace(szText, "e", "3")                              ' E = 3
    szText = Replace(szText, "F", "f")                              ' F = f
    szText = Replace(szText, "g", "6")
    szText = Replace(szText, "G", "6")                              ' G = 6
    szText = Replace(szText, "H", "h")                              ' H = h
    szText = Replace(szText, "i", "1")
    szText = Replace(szText, "I", "1")                              ' i = 1
    szText = Replace(szText, "J", "j")                              ' J = j
    szText = Replace(szText, "k", "X")
    szText = Replace(szText, "K", "X")                              ' K = X
    szText = Replace(szText, "l", "1")
    szText = Replace(szText, "L", "1")                              ' L = 1
    szText = Replace(szText, "M", "m")                              ' M = m
    szText = Replace(szText, "N", "n")                              ' N = n
    szText = Replace(szText, "O", "0")
    szText = Replace(szText, "o", "0")                              ' O = 0
    szText = Replace(szText, "p", "9")
    szText = Replace(szText, "P", "9")                              ' P = 9
    szText = Replace(szText, "Q", "q")                              ' Q = q
    szText = Replace(szText, "R", "r")                              ' R = r
    szText = Replace(szText, "s", "5")
    szText = Replace(szText, "S", "5")                              ' S = 5
    szText = Replace(szText, "t", "7")
    szText = Replace(szText, "T", "7")                              ' T = 7
    szText = Replace(szText, "U", "u")                              ' U = u
    szText = Replace(szText, "V", "v")                              ' V = v
    szText = Replace(szText, "W", "w")                              ' W = w
    szText = Replace(szText, "X", "x")                              ' X = x
    szText = Replace(szText, "Y", "y")                              ' Y = y
    szText = Replace(szText, "z", "2")
    szText = Replace(szText, "Z", "2")                              ' Z = 2
    TextTo1337 = szText
exithandler:
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
'    Call objError.Errorhandler(MODULNAME, "TextTo1337", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function TextToHex(szText As String, Optional szSpaceChar As String) As String
    Dim szRest As String
    Dim szResult As String
    Dim szChar As String
    Dim lngASC As Long
    Dim szHex As String
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    szRest = szText
    While szRest <> ""
        szChar = Left(szRest, 1)
        szRest = Right(szRest, Len(szRest) - 1)
        lngASC = Asc(szChar)
        szResult = szResult & szSpaceChar & Hex(lngASC)
    Wend
    TextToHex = szResult
exithandler:
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
'    Call objError.Errorhandler(MODULNAME, "TextToHex", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function TextToMorse(szText As String) As String
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    szText = Replace(szText, "a", ".-")                             ' A = .-
    szText = Replace(szText, "A", ".-")
    szText = Replace(szText, "b", "-...")                           ' B = -...
    szText = Replace(szText, "B", "-...")
    szText = Replace(szText, "C", "-.-.")                           ' C = -.-.
    szText = Replace(szText, "c", "-.-.")
    szText = Replace(szText, "D", "-..")                            ' D = -..
    szText = Replace(szText, "d", "-..")
    szText = Replace(szText, "E", ".")                              ' E = .
    szText = Replace(szText, "e", ".")
    szText = Replace(szText, "F", "..-.")                           ' F = ..-.
    szText = Replace(szText, "f", "..-.")
    szText = Replace(szText, "G", "--.")                            ' G = --.
    szText = Replace(szText, "g", "--.")
    szText = Replace(szText, "H", "....")                           ' H = ....
    szText = Replace(szText, "h", "....")
    szText = Replace(szText, "i", "..")
    szText = Replace(szText, "I", "..")                             ' i = ..
    szText = Replace(szText, "J", ".---")                           ' J = .---
    szText = Replace(szText, "j", ".---")
    szText = Replace(szText, "k", "-.-")
    szText = Replace(szText, "K", "-.-")                            ' K = -.-
    szText = Replace(szText, "l", ".-..")
    szText = Replace(szText, "L", ".-..")                           ' L = .-..
    szText = Replace(szText, "M", "--")                             ' M = --
    szText = Replace(szText, "m", "--")
    szText = Replace(szText, "N", "-.")                             ' N = -.
    szText = Replace(szText, "n", "-.")
    szText = Replace(szText, "O", "---")
    szText = Replace(szText, "o", "---")                            ' O = ---
    szText = Replace(szText, "p", ".--.")
    szText = Replace(szText, "P", ".--.")                           ' P = .--.
    szText = Replace(szText, "Q", "--.-")                           ' Q = --.-
    szText = Replace(szText, "q", "-.-")
    szText = Replace(szText, "R", ".-.")                            ' R = .-.
    szText = Replace(szText, "s", "...")
    szText = Replace(szText, "S", "...")                            ' S = ...
    szText = Replace(szText, "t", "-")
    szText = Replace(szText, "T", "-")                              ' T = -
    szText = Replace(szText, "U", "..-")                            ' U = ..-
    szText = Replace(szText, "u", "..-")
    szText = Replace(szText, "V", "...-")                           ' V = ...-
    szText = Replace(szText, "v", "...-")
    szText = Replace(szText, "W", ".--")                            ' W = .--
    szText = Replace(szText, "w", ".--")
    szText = Replace(szText, "X", "-..-")                           ' X = -..-
    szText = Replace(szText, "x", "-..-")
    szText = Replace(szText, "Y", "-.--")                           ' Y = -.--
    szText = Replace(szText, "y", "-.--")
    szText = Replace(szText, "z", "--..")
    szText = Replace(szText, "Z", "--..")                           ' Z = --..
    
    szText = Replace(szText, "1", "-----")                          ' 1 = -----
    szText = Replace(szText, "2", "..---")                          ' 2 = ..---
    szText = Replace(szText, "3", "...--")                          ' 3 = ...--
    szText = Replace(szText, "4", "....-")                          ' 4 = ....-
    szText = Replace(szText, "5", ".....")                          ' 5 = .....
    szText = Replace(szText, "6", "-....")                          ' 6 = -....
    szText = Replace(szText, "7", "--...")                          ' 7 = --...
    szText = Replace(szText, "8", "---..")                          ' 8 = ---..
    szText = Replace(szText, "9", "----.")                          ' 9 = ----.
    
    TextToMorse = szText
exithandler:
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
    'Call objError.Errorhandler(MODULNAME, "TextToMorse", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function TextToBin(szText As String) As String
    Dim szRest As String
    Dim szResult As String
    Dim szChar As String
    Dim lngASC As Long
    Dim szHex As String
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    szRest = szText
    While szRest <> ""
        szChar = Left(szRest, 1)
        szRest = Right(szRest, Len(szRest) - 1)
        lngASC = Asc(szChar)
        szHex = Hex(lngASC)
        szResult = szResult & " " & DecToBin(lngASC)
    Wend
    TextToBin = szResult
exithandler:
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
'    Call objError.Errorhandler(MODULNAME, "TextToBin", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Private Function DecToBin(lngDecimal As Long) As String
'Die Dezimalzahl 48 wird ins 2er-System umgewandelt
'Gehe nach folgendem Verfahren vor:
' (1) Teile die Zahl mit Rest durch 2.
' (2) Der Divisionsrest ist die nächste Ziffer (von rechts nach links).
' (3) Falls der (ganzzahlige) Quotient = 0 ist, bist du fertig,
'     andernfalls nimm den (ganzzahligen) Quotienten als neue Zahl
'     und wiederhole ab (1).
'     48 : 2 = 24  Rest: 0
'     24 : 2 = 12  Rest: 0
'     12 : 2 =  6  Rest: 0
'      6 : 2 =  3  Rest: 0
'      3 : 2 =  1  Rest: 1
'      1 : 2 =  0  Rest: 1
'     Resultat: 110000
    Dim szResult As String
    Dim lngTmp As Long
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    lngTmp = lngDecimal
    While lngTmp > 0
        
        szResult = CStr(lngTmp Mod 2) & szResult
        lngTmp = lngTmp \ 2
    Wend
    While Len(szResult) < 8
        szResult = "0" & szResult
    Wend
    DecToBin = szResult
exithandler:
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
'    Call objError.Errorhandler(MODULNAME, "DecToBin", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function FormatMem(szMembyte As String, lngUnit As Integer) As String
' Formatiert Speichergrößen
    Dim lngByte As Double
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If szMembyte = "" Then szMembyte = "0"
    lngByte = CDbl(szMembyte)
    Select Case lngUnit
    Case 0                                                          ' Byte
        lngByte = Round(lngByte)
        FormatMem = CStr(lngByte) & " Byte"
    Case 1                                                          ' Kilobyte
        lngByte = lngByte / 1024
        lngByte = Round(lngByte)
        FormatMem = CStr(lngByte) & " KB"
    Case 2                                                          ' Mega Byte
        lngByte = (lngByte / 1024) / 1024
        lngByte = Round(lngByte)
        FormatMem = CStr(lngByte) & " MB"
    Case 3                                                          ' Giga Byte
        lngByte = ((lngByte / 1024) / 1024) / 1024
        lngByte = Round(lngByte)
        FormatMem = CStr(lngByte) & " GB"
    Case 4                                                          ' Terra Byte
    
    Case Else ' Nur Byte
    
    End Select
exithandler:
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
'    Call objError.Errorhandler(MODULNAME, "FormatMem", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function FormatHZ(szMHz As String, lngUnit As Integer) As String
' Formateri Herz angaben
    Dim lngMhz As Double
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If szMHz = "" Then szMHz = "0"
    lngMhz = CDbl(szMHz)
    Select Case lngUnit
    Case 0
        lngMhz = Round(lngMhz)
        lngMhz = (lngMhz * 1000) * 1000
         FormatHZ = CStr(lngMhz) & " Hz"
    Case 1 ' kHz
        lngMhz = lngMhz * 1000
        lngMhz = Round(lngMhz)
        FormatHZ = CStr(lngMhz) & " KHz"
    Case 2 ' MHz
        lngMhz = Round(lngMhz)
        FormatHZ = CStr(lngMhz) & " MHz"
    Case 3  ' GHz
        lngMhz = lngMhz / 1000
        lngMhz = Round(lngMhz)
        FormatHZ = CStr(lngMhz) & " GHz"
    Case Else ' Nur Byte
    
    End Select
exithandler:
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
'    Call objError.Errorhandler(MODULNAME, "FormatHZ", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

Public Function FormatBiosDate(szBDate As String) As String
    Dim szYear As String
    Dim szMonth As String
    Dim szDay As String
On Error GoTo Errorhandler                                          ' Fehlerbehandlung aktivieren
    If Len(szBDate) < 8 Then FormatBiosDate = ""
    szYear = Left(szBDate, 4)
    szMonth = Mid(szBDate, 5, 2)
    szDay = Mid(szBDate, 7, 2)
    FormatBiosDate = szDay & "." & szMonth & "." & szYear
exithandler:
Exit Function
Errorhandler:
    Dim errNr As String                                             ' Fehlernummer
    Dim errDesc As String                                           ' Fehler beschreibung
    errNr = Err.Number                                              ' Fehlernummer auslesen
    errDesc = Err.Description                                       ' Fehler beschreibung auslesen
    Err.Clear                                                       ' Fehler Clearen
On Error Resume Next                                                ' Keinen Fehler in de rFehlerbehandlung zulassen
'    Call objError.Errorhandler(MODULNAME, "FormatBiosDate", errNr, errDesc) ' Fehler behandlung aufrufen
    Resume exithandler                                              ' Weiter mit Exithandler
End Function

