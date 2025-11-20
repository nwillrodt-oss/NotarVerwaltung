Attribute VB_Name = "modOptions"
Option Explicit

Private Const MODULNAME = "modOptions"

Public Type OptionValue
    Name As String
    Caption As String
    Value As Variant
    bCrypt As Boolean
End Type

Public VarOptions() As OptionValue     ' Optionen

Public Function InitOptions()
' Optionen aug INI bez. Reg einlesen
    
On Error GoTo ErrorHandler
    
    ' ValueArray Init
    ReDim VarOptions(0)
        
    Call InitOptionValue(OPTION_SPLIT, "", "0,3", INI_APP)              ' Spliter POS fürs DB Windows
    Call InitOptionValue(OPTION_MAINSIZE, "", "6000/10000", INI_APP)    ' Size Main Window
    Call InitOptionValue(OPTION_DBSTATE, "", vbNormal, INI_APP)         ' WindowState DB Windows
    Call InitOptionValue(OPTION_DBSIZE, "", "6000/1500", INI_APP)       ' Size DB Windows
    Call InitOptionValue(OPTION_LASTCON, "Letze DB Verbindung", "", "") ' Letzte DB Verbindung
    Call InitOptionValue(OPTION_SPLASH, "Splash anzeigen", CStr(bNotShowSplash), INI_APP)   ' Splash anzeigen
    Call InitOptionValue(OPTION_SPLASH_IMG, "Splash Image", "images\Splash.jpg", INI_APP)   ' Splash Picture
    Call InitOptionValue(OPTION_ABOUT_IMG, "Info Image", "images\Wagge.gif", INI_APP)       ' About Image
    Call InitOptionValue(OPTION_ERRLOG, OPTION_ERRLOG, "ErrorLog.txt", INI_APP)             ' Name Error LogDatei
    objError.SetErrFileName = objObjectBag.GetAppDir & GetOptionByName(OPTION_ERRLOG)       ' Gleich an error Obj Weitergeben
    Call InitOptionValue(OPTION_APPLOG, OPTION_APPLOG, "Log.txt", INI_APP)                  ' Name LogDatei
    objError.SetProtFileName = objObjectBag.GetAppDir & GetOptionByName(OPTION_APPLOG)      ' Gleich an error Obj Weitergeben
    
exithandler:
On Error Resume Next

Exit Function
ErrorHandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.ErrorHandler(MODULNAME, "InitOptions", errNr, errDesc)
    Resume exithandler
End Function

Public Function SaveOptions()

    Dim i As Integer ' Counter
    Dim szRegKey As String
    
On Error GoTo ErrorHandler
    
    szRegKey = "SOFTWARE\" & SZ_APPTITLE & "\"
    
    For i = 0 To UBound(VarOptions)
        With VarOptions(i)
            Call objRegTools.WriteRegValue("HKCU", szRegKey, .Name, CStr(.Value))
        End With
    Next i
    
exithandler:
On Error Resume Next

Exit Function
ErrorHandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.ErrorHandler(MODULNAME, "SaveOptions", errNr, errDesc)
    Resume exithandler
End Function

Public Function InitOptionValue(szValueName As String, _
        szValueCaption As String, _
        szDefaultValue As String, _
        szIniSection As String, _
        Optional bCrypt As Boolean)

    Dim szValue As String
    
On Error GoTo ErrorHandler

    ReDim Preserve VarOptions(UBound(VarOptions) + 1)
    VarOptions(UBound(VarOptions)).Name = szValueName
    VarOptions(UBound(VarOptions)).Caption = szValueCaption
    VarOptions(UBound(VarOptions)).bCrypt = bCrypt
    
    If Trim(szIniSection) <> "" Then
        '   Erst Defaultwerte aus Inidateilesen
        szValue = objTools.GetINIValue(App.Path & "\" & INI_FILENAME, szIniSection, szValueName)
    End If
        
    If bCrypt Then
        ' Entschlüsseln
        szValue = objTools.Crypt(szValue, False)
    End If
    
    ' Dann mit Reg einträgen überschreiben
    If bCrypt Then
        ' Verschlüsseln
        szValue = objTools.Crypt(szValue, True)
    End If
    If szDefaultValue = "" Then szDefaultValue = szValue
    szValue = Trim(objRegTools.ReadRegValue("HKCU", "SOFTWARE\" & SZ_APPTITLE, szValueName, szDefaultValue))
    szValue = objTools.cutlastChar(szValue, Chr(32))
    If bCrypt Then
        ' Entschlüsseln
        szValue = objTools.Crypt(szValue, False)
    End If
    VarOptions(UBound(VarOptions)).Value = szValue
    'Call VarOptions(UBound(VarOptions)).SetValue(szValue)
    
exithandler:
On Error Resume Next

Exit Function
ErrorHandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.ErrorHandler(MODULNAME, "InitOptionValue", errNr, errDesc)
    Resume exithandler
End Function

Public Function SetOptionByName(szOptionName As String, Value As Variant)

Dim i As Integer

On Error GoTo ErrorHandler

    For i = 0 To UBound(VarOptions)
        If UCase(VarOptions(i).Name) = UCase(szOptionName) Then
            VarOptions(i).Value = Value
            Exit For
        End If
    Next i
    
exithandler:
On Error Resume Next

Exit Function
ErrorHandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.ErrorHandler(MODULNAME, "SetOptionByName", errNr, errDesc)
    Resume exithandler
End Function

Public Function GetOptionByName(szOptionName As String) As Variant

Dim i As Integer

On Error GoTo ErrorHandler

    For i = 0 To UBound(VarOptions)
        If UCase(VarOptions(i).Name) = UCase(szOptionName) Then
            GetOptionByName = VarOptions(i).Value
            Exit For
        End If
    Next i
    
exithandler:
On Error Resume Next

Exit Function
ErrorHandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.ErrorHandler(MODULNAME, "GetOptionByName", errNr, errDesc)
    Resume exithandler
End Function

Public Function SetOptionByCaption(szOptionCaption As String, Value As Variant)

Dim i As Integer

On Error GoTo ErrorHandler

    For i = 0 To UBound(VarOptions)
        If UCase(VarOptions(i).Caption) = UCase(szOptionCaption) Then
            VarOptions(i).Value = Value
            Exit For
        End If
    Next i

exithandler:
On Error Resume Next

Exit Function
ErrorHandler:
    Dim errNr As String
    Dim errDesc As String
    errNr = Err.Number
    errDesc = Err.Description
    Err.Clear
    Call objError.ErrorHandler(MODULNAME, "SetOptionByCaption", errNr, errDesc)
    Resume exithandler
End Function

