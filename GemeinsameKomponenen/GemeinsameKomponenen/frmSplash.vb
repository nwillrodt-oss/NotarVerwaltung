Public Class frmSplash
    Private Const MODULNAME = "frmSplash"                                       ' Modulname für Fehlerbehandlung

    Private objBag As clsObjectBag                                              ' Sammelobject

    Public Sub New(ByVal oBag As clsObjectBag, Optional ByVal ImagePath As String = "")
        ' Dieser Aufruf ist für den Designer erforderlich.
        InitializeComponent()

        ' Fügen Sie Initialisierungen nach dem InitializeComponent()-Aufruf hinzu.
        objBag = oBag
        If Not InitSplash(ImagePath) Then
            Me.Close()
        End If
    End Sub

    Private Function InitSplash(Optional ByVal ImagePath As String = "") As Boolean
        Dim oApp As clsApp
        Try                                                                     ' Fehlerbehandlung aktivieren
            oApp = objBag.oClsApp
            If ImagePath <> "" Then
                Dim newImage As Image = Image.FromFile(ImagePath)
                Me.PictureSplash.Image = newImage
            End If
            Me.lblAppTitel.Text = oApp.AppTitle
            Me.lblCopyright.Text = oApp.Copyright
            Me.lblWWW.Text = ""  'oApp.
            Me.lblVersion.Text = oApp.AppVersion
            Me.lblAktion.Text = ""
            Me.Refresh()
            Return True
        Catch ex As Exception                                                   ' Fehlerbehandlung
            Call objBag.ErrorHandler(MODULNAME, "InitSplash", ex)               ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

End Class