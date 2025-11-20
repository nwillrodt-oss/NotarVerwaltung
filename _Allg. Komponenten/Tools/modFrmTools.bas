Attribute VB_Name = "modFrmTools"
Option Explicit                                                     ' Variaben Deklaration erzwingen
Option Compare Text                                                 ' Sortierreihenfolge festlegen
Const MODULNAME = "modFormTools"                                    ' Modulname für Fehlerbehandlung

Public objError As Object                                           ' Error Object


Public Sub SetFormTransparency(F As Form, sinPercent As Single)
' Setzt Fenster transparenz
' funktioniert nur unter Windows 2000 oder XP!!!
' RateOfT: 254 = normal 0 = ganz transparent (also unsichtbar)
    Dim WinInfo As Long                                             ' Windows Handel
    Dim RateOfT As Byte                                             ' Transparenz Rate 255 Keine, 0 Totale Tranzparenz
On Error Resume Next                                                ' Keine Fehlerbehandlung
    If F Is Nothing Then GoTo exithandler                           ' Kein Form -> Fertig
    WinInfo = GetWindowLong(F.hWnd, GWL_EXSTYLE)
    If sinPercent < 0.1 Then sinPercent = 0.1                       ' Minimal tranzparenz festlegen
    RateOfT = sinPercent * 254
    If RateOfT < 255 Then                                           ' Tranzparenz setzen
        WinInfo = WinInfo Or WS_EX_LAYERED
        SetWindowLong F.hWnd, GWL_EXSTYLE, WinInfo
        SetLayeredWindowAttributes F.hWnd, 0, RateOfT, LWA_ALPHA
    Else                                                            ' Wenn als Rate 255 angegeben wird,
        WinInfo = WinInfo Xor WS_EX_LAYERED                         ' so wird der Ausgangszustand wiederhergestellt
        SetWindowLong F.hWnd, GWL_EXSTYLE, WinInfo
    End If
exithandler:
    Err.Clear                                                       ' Evtl. error clearen
End Sub

Public Function CtlIsOptionButton(Ctl As Control) As Boolean
    Dim optCtl As OptionButton                                      ' OptionButton control
On Error Resume Next                                                ' Keine Fehlerbehandlung
    Set optCtl = Ctl                                                ' Ctl (Control) als OptionButton setzen
    If Err = 0 Then                                                 ' Ist kein Fehler aufgetreten
        CtlIsOptionButton = True                                    ' Erfolg zurück
    End If
    Err.Clear                                                       ' Evtl. error clearen
End Function

Public Function CtlIsTextBox(Ctl As Control) As Boolean
    Dim txtCtl As TextBox                                           ' Textbox control
On Error Resume Next                                                ' Keine Fehlerbehandlung
    Set txtCtl = Ctl                                                ' Ctl (Control) als Textbox setzen
    If Err = 0 Then                                                 ' Ist kein Fehler aufgetreten
        CtlIsTextBox = True                                         ' Erfolg zurück
    End If
    Err.Clear                                                       ' Evtl. error clearen
End Function

Public Function CtlIsLabel(Ctl As Control) As Boolean
    Dim lblCtl As Label                                             ' Label control
On Error Resume Next                                                ' Keine Fehlerbehandlung
    Set lblCtl = Ctl                                                ' Ctl (Control) als Label setzen
    If Err = 0 Then                                                 ' Ist kein Fehler aufgetreten
        CtlIsLabel = True                                           ' Erfolg zurück
    End If
    Err.Clear                                                       ' Evtl. error clearen
End Function

Public Function CtlIsButton(Ctl As Control) As Boolean
    Dim cmdCtl As CommandButton                                     ' Button control
On Error Resume Next                                                ' Keine Fehlerbehandlung
    Set cmdCtl = Ctl                                                ' Ctl (Control) als Button setzen
    If Err = 0 Then                                                 ' Ist kein Fehler aufgetreten
        CtlIsButton = True                                          ' Erfolg zurück
    End If
    Err.Clear                                                       ' Evtl. error clearen
End Function

Public Function CtlIsCheck(Ctl As Control) As Boolean
    Dim chkCtl As CheckBox                                          ' Checkbox control
On Error Resume Next                                                ' Keine Fehlerbehandlung
    Set chkCtl = Ctl                                                ' Ctl (Control) als Button setzen
    If Err = 0 Then                                                 ' Ist kein Fehler aufgetreten
        CtlIsCheck = True                                           ' Erfolg zurück
    End If
    Err.Clear                                                       ' Evtl. error clearen
End Function

Public Function CtlIsFrame(Ctl As Control) As Boolean
    Dim frameCtl As Frame                                           ' Frame control
On Error Resume Next                                                ' Keine Fehlerbehandlung
    Set frameCtl = Ctl                                              ' Ctl (Control) als Button setzen
    If Err = 0 Then                                                 ' Ist kein Fehler aufgetreten
        CtlIsFrame = True                                           ' Erfolg zurück
    End If
    Err.Clear                                                       ' Evtl. error clearen
End Function

Public Function CtlIsStatusbar(Ctl As Control) As Boolean
    Dim lngPanelCount As Integer                                    ' Pannel anzahl
On Error Resume Next                                                ' Keine Fehlerbehandlung
    lngPanelCount = Ctl.Panels.Count                                ' Pannel anz. eritteln
    If Err = 0 Then                                                 ' Ist kein Fehler aufgetreten
        CtlIsStatusbar = True                                       ' Erfolg zurück
    End If
    Err.Clear                                                       ' Evtl. error clearen
End Function

Public Function CtlIsCombo(Ctl As Control) As Boolean
    Dim comboCtl As ComboBox                                        ' Combo control
On Error Resume Next                                                ' Keine Fehlerbehandlung
    Set comboCtl = Ctl                                              ' Ctl (Control) als Button setzen
    If Err = 0 Then                                                 ' Ist kein Fehler aufgetreten
        CtlIsCombo = True                                           ' Erfolg zurück
    End If
    Err.Clear                                                       ' Evtl. error clearen
End Function
