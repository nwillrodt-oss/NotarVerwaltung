Public Class frmConParams
    Private Const MODULNAME = "frmConParams"                                    ' Modulname für Fehlerbehandlung

    Private objBag As clsObjectBag                                              ' Sammelobject
    Private ObjCon As clsDBConnect                                              ' DB Verbindungsklasse
    Public bCancel As Boolean                                                   ' Abbruch variable 

#Region "Constructor"

    Public Sub New(ByVal oBag As clsObjectBag)
        ' Dieser Aufruf ist für den Designer erforderlich.
        InitializeComponent()

        ' Fügen Sie Initialisierungen nach dem InitializeComponent()-Aufruf hinzu.
        objBag = oBag                                                           ' Objectbag festlegen
        ObjCon = objBag.ObjDBConnect                                            ' Verbindungsklasse holen

        If InitConParamForm() Then

        End If
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

#End Region

    Private Function InitConParamForm()
        ' Initialisiert das Parameter form
        Try                                                                     ' Fehlerbehandlung aktivieren
            If ObjCon.AccessPossible And ObjCon.SQLPossible Then                ' Nur wen beide Verbindungstypen möglich
                Me.radAcc.Visible = True                                        ' Radiobuttons einblenden
                Me.radSQL.Visible = True
            Else                                                                ' Sonst
                Me.radAcc.Visible = False                                       ' Radiobuttons ausblenden
                Me.radSQL.Visible = False
            End If
            If ObjCon.SQLPossible Then
                Me.radSQL.Checked = True
                Call SwitchConType(1)
            Else
                Me.radSQL.Checked = False
                Call SwitchConType(2)
            End If
            Return True                                                         ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehlerbehandlung
            Call objBag.ErrorHandler(MODULNAME, "InitConParamForm", ex)         ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function SwitchConType(ByVal ConType As Integer)
        Try                                                                     ' Fehlerbehandlung aktivieren
            Select Case ConType                                                 ' Connectiontyp auswerten
                Case 1                                                          ' 1 = SQL server
                    chkNT.Visible = True
                    lblServer.Visible = True
                    txtServer.Visible = True
                    txtServer.Text = ObjCon.SQLServer
                    lblDBName.Text = "Datenbankname:"                    
                    btnDBPath.Visible = False
                    txtDBName.Text = ObjCon.DBName
                    txtDBUSer.Text = ObjCon.DBUser
                Case 2                                                          ' 2 = Access
                    chkNT.Visible = False
                    lblServer.Visible = False
                    txtServer.Visible = False
                    lblDBName.Text = "Datenbank:"
                    txtDBName.Text = ObjCon.DBName
                    txtDBUSer.Text = ObjCon.DBUser
                    btnDBPath.Visible = True
                Case Else

            End Select
            Return True                                                         ' erfolg zurück
        Catch ex As Exception                                                   ' Fehlerbehandlung
            Call objBag.ErrorHandler(MODULNAME, "SwitchConType", ex)            ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Sub radSQL_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles radSQL.CheckedChanged
        Call SwitchConType(1)
    End Sub

    Private Sub radAcc_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles radAcc.CheckedChanged
        Call SwitchConType(2)
    End Sub

    Private Sub btnEsc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEsc.Click
        bCancel = True                                                          ' Abbruch
        Me.Close()                                                              ' Form schliessen
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click

        Me.Close()                                                              ' Form schliessen
    End Sub
End Class