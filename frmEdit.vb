Public Class frmEdit
    Private Const MODULNAME = "frmEdit"                                         ' Modulname für Fehlerbehandlung
    Private ObjBag As clsObjectBag                                              ' Sammelobject
    Private ObjCon As clsDBConnect                                              ' Datenbank Verbindungs klasse
    Private bInit As Boolean
    Private ObjValue As clsValueList                                            ' Werte Klasse
    'Private oDataView As DataView

    Const DefLeftPos = 5
    Const DefCtlHeight = 20
    Const CtlDiff = 5
    Const GrpBoxDiff = 5

#Region "Constructor"

    Public Sub New(ByVal oBag As clsObjectBag, _
                   ByVal RootKey As String, _
                   ByVal DetailKey As String)
        ' Dieser Aufruf ist für den Designer erforderlich.
        InitializeComponent()

        ' Fügen Sie Initialisierungen nach dem InitializeComponent()-Aufruf hinzu.
        ObjBag = oBag                                                           ' Objectbag festlegen
        ObjCon = ObjBag.ObjDBConnect                                            ' Verbindungsklasse holen

        If InitEditForm(RootKey, DetailKey) Then

        End If
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

#End Region

#Region "Properties"

    Public ReadOnly Property IsNew() As Boolean                                 ' Neuer Datensatz ?
        Get
            IsNew = ObjValue.IsNew
        End Get
    End Property

    Public ReadOnly Property ID() As String                                     ' ID Value
        Get
            ID = ObjValue.ID
        End Get
    End Property

    Public ReadOnly Property Dirty() As Boolean                                 ' Geändert ?
        Get
            Dirty = ObjValue.Dirty
        End Get
    End Property

#End Region

#Region "Init Funktionen"

    Private Function InitEditForm(ByVal RootKey As String, _
                                  Optional ByVal DetailKey As String = "")
        Dim szDeteils As String = ""                                            ' Details für Fehlerbehandlung
        Try                                                                     ' Fehlerbehandlung aktivieren
            ObjValue = New clsValueList(ObjBag, RootKey, DetailKey)             ' WertKlasse initialisieren
            If ObjValue.InitOK Then
                Me.BindingContext(ObjValue.DataSet).Position = 0
                szDeteils = "Name: " & RootKey & " DS ID: " & DetailKey         ' Details für Fehlerbehandlung
                bInit = True                                                    ' Wir initialisieren das Form-> andere vorgänge nicht ausführen
                'frmParent = ParentForm
                If ObjValue.IsNew Then                                          ' Wenn Neuer DS
                    'Me.cmdDel.Enabled = False                                   ' Löschen Disablen
                Else                                                            ' Wenn bestehender DS
                    'Me.cmdDel.Enabled = True                                    ' Löschen enablen
                End If
                'If ObjValue.IsDeletable Then                                    ' DS ist Löschbar
                '    Me.cmdDel.Visible = True                                    ' Löschen Sichtbar
                'Else                                                            ' Sonst
                '    Me.cmdDel.Visible = False                                   ' Löschen unsichtbar
                'End If

                Me.Refresh()                                                    ' Form aktualisieren
                Me.Text = ObjValue.Caption                                      ' Form Caption abhängig von den daten setzen
                Call InitEditFrame()                                            ' Reiter 'Allgemein' initialisiren
            End If
            bInit = False                                                       ' Initialisierung abgeschlossen
            Return True
        Catch ex As Exception                                                   ' Fehlerbehandlung
            Call ObjBag.ErrorHandler(MODULNAME, "InitEditForm", ex, szDeteils)  ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function InitEditFrame()
        Dim i As Integer                                                        ' counter
        Dim szFieldName As String                                               ' Feldname
        Dim szText As String                                                    ' Textwert eines Datenfeldes
        Dim szToolTip As String                                                 ' Tooltip
        Dim dTable As DataTable
        Dim TopPosGrp As Integer = 5                                            ' Top Position für Groupboxen (wird hochgezählt)
        Dim TopPosCtl As Integer = 5                                            ' Top Position für Controls in Groupbox (wird Hochgezählt)
        Dim GrpHeight As Integer = 0                                            ' Höhe der akt Grop Box
        Dim FormHeight As Integer = 5                                           ' Höhe des forms
        Dim fInfo As clsValueList.FieldInfo
        Try                                                                     ' Fehlerbehandlung aktivieren
            GrpHeight = 15                                                      ' (platz vor dem 1. control)
            TopPosCtl = 15
            dTable = ObjValue.DataSet.Tables(0)
            For i = 0 To dTable.Columns.Count - 1                               ' Alle Felder durchlaufen
                szFieldName = dTable.Columns(i).Caption
                If Not ObjValue.IsNew Then                                      ' Falls Kein Neuer DS
                    szText = dTable.Rows(0).Item(szFieldName).ToString          ' Wert Aus DS holen
                Else                                                            ' Sonst
                    szText = ObjValue.DefaultValue(szFieldName)                 ' Default value holen
                End If

                'szToolTip = szFieldName & " (" & _
                '    ObjCon.ConvertColumnType(dTable.Columns(i).DataType, _
                '    dTable.Columns(i).MaxLength) & ")"              ' ToolTip aus Feldnamen und Size zusamensetzen
                fInfo = ObjValue.GetFieldInfo(szFieldName)                      ' Feld informationen holen
                szToolTip = szFieldName                                         ' tooltip erstmal so
                If fInfo.Valuelist <> "" Or fInfo.ValueListSQL <> "" Then       ' gibt es eine Werteliste
                    Call AddValueKombo(GrpBoxDaten, szFieldName, szText, TopPosCtl, fInfo, szToolTip, szFieldName, szFieldName) ' Kombobox hinzu
                Else                                                            ' Sonst 
                    Call AddValueText(GrpBoxDaten, szFieldName, szText, TopPosCtl, fInfo, szToolTip, szFieldName, szFieldName) ' Textbox hinzu
                End If
                If fInfo.bVisible Then                                          ' Nue wenn control sichtbar
                    GrpHeight = GrpHeight + DefCtlHeight                        ' Höhe GrpBox um Höhe Kindcontrol hochzälen
                    GrpHeight = GrpHeight + CtlDiff                             ' Höhe GrpBox um abstand der Kindcontrols hochzählen
                    TopPosCtl = TopPosCtl + DefCtlHeight + CtlDiff              ' Neue KindControl lHöhe bsetimmen
                End If
                szFieldName = ""
                szToolTip = ""
                szText = ""
            Next                                                                ' Nächstes Feld
            If GrpBoxDaten.Controls.Count = 0 Then

            Else
                GrpBoxDaten.Size = New System.Drawing.Size(Me.Width - (3 * DefLeftPos), GrpHeight) ' Groupbox Größe festlegen
                GrpBoxDaten.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
                TopPosGrp = TopPosGrp + GrpBoxDaten.Size.Height + GrpBoxDiff    ' Neue TopPos für GrpBox bestimmen
                'Me.Controls.Add(GrpBoxDaten)                                   ' Groupbox ans fom hängen
                FormHeight = FormHeight + GrpBoxDaten.Size.Height + GrpBoxDiff  ' Fomular höhe hochzählen
            End If

            FormHeight = FormHeight + Me.btnEsc.Size.Height + 10
            If FormHeight < Me.MinimumSize.Height Then FormHeight = Me.MinimumSize.Height ' Minimum größe beachten
            If FormHeight > Me.Size.Height Then

            Else

            End If
            Dim btnTop As Integer = FormHeight - Me.btnEsc.Size.Height - 30     ' Top Pos. für Buttonleiste
            Me.lblDirty.Location = New System.Drawing.Point(Me.lblDirty.Location.X, btnTop) ' Pos. setzen
            Me.btnEsc.Location = New System.Drawing.Point(Me.btnEsc.Location.X, btnTop) ' Pos. setzen
            Me.btnOK.Location = New System.Drawing.Point(Me.btnOK.Location.X, btnTop) ' Pos. setzen
            Me.btnSave.Location = New System.Drawing.Point(Me.btnSave.Location.X, btnTop) ' Pos. setzen
            Return True                                                         ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehlerbehandlung
            Call ObjBag.ErrorHandler(MODULNAME, "InitEditFrame", ex)            ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function AddValueText(ByRef GrpBox As System.Windows.Forms.GroupBox, _
                                  ByVal Caption As String, _
                                  ByVal TextValue As String, _
                                  ByVal TopPos As Integer, _
                                  ByVal InfoField As clsValueList.FieldInfo, _
                                  Optional ByVal szToolTip As String = "", _
                                  Optional ByVal szTag As String = "", _
                                  Optional ByVal szDatafield As String = "") As Boolean

        Dim TxtBox As System.Windows.Forms.TextBox                              ' TextBox für Boolwerte
        Dim LblBox As System.Windows.Forms.Label                                ' Label für Caption
        Dim TxtBoxWidth As Integer
        Const LblWidth = 150
        Try                                                                     ' Fehlerbehandlung aktivieren
            TxtBox = New System.Windows.Forms.TextBox                           ' Neue Textbox erstellen
            LblBox = New System.Windows.Forms.Label                             ' Neues Lable erstellen
            TxtBoxWidth = GrpBox.Width - (2 * DefLeftPos) - LblWidth - 20
            LblBox.Name = "lbl" & szDatafield                                   ' Namen Setzen
            TxtBox.Name = "txt" & szDatafield                                   ' Namen Setzen
            LblBox.Location = New System.Drawing.Point(DefLeftPos, TopPos)
            LblBox.Size = New System.Drawing.Size(LblWidth, DefCtlHeight)       ' Grüße festlegen
            TxtBox.Location = New System.Drawing.Point(LblWidth + DefLeftPos, TopPos)
            TxtBox.Size = New System.Drawing.Size(TxtBoxWidth, DefCtlHeight)    ' Grüße festlegen
            TxtBox.Tag = szDatafield
            LblBox.Text = Caption                                               ' Bezeichnung
            TxtBox.Text = TextValue                                             ' Datenwert setzen
            AddHandler TxtBox.TextChanged, AddressOf Value_TextChanged          ' Handler binden
            LblBox.Anchor = AnchorStyles.Top Or AnchorStyles.Left               ' Verankern
            TxtBox.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right ' Verankern
            If Not IsNothing(InfoField) Then
                LblBox.Visible = InfoField.bVisible                                           ' Sichtbar setzen
                TxtBox.Visible = InfoField.bVisible
                LblBox.Enabled = InfoField.bEnabled                                           ' Enablen
                TxtBox.Enabled = InfoField.bEnabled
                TxtBox.ReadOnly = InfoField.bLocked                                           ' Eingaben sperren
            End If

            GrpBox.Controls.Add(TxtBox)                                         ' An Groupbox anhängen
            GrpBox.Controls.Add(LblBox)                                         ' An Groupbox anhängen
            Return True
        Catch ex As Exception                                                   ' Fehlerbehandlung
            Call ObjBag.ErrorHandler(MODULNAME, "AddValueText", ex)             ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function AddValueKombo(ByRef GrpBox As System.Windows.Forms.GroupBox, _
                                  ByVal Caption As String, _
                                  ByVal TextValue As String, _
                                  ByVal TopPos As Integer, _
                                  ByVal InfoField As clsValueList.FieldInfo, _
                                  Optional ByVal szToolTip As String = "", _
                                  Optional ByVal szTag As String = "", _
                                  Optional ByVal szDatafield As String = "") As Boolean

        Dim CmbBox As System.Windows.Forms.ComboBox                             ' ComboBox für Boolwerte
        Dim LblBox As System.Windows.Forms.Label                                ' Label für Caption
        Dim CmbBoxWidth As Integer
        Const LblWidth = 150
        Try                                                                     ' Fehlerbehandlung aktivieren
            CmbBox = New System.Windows.Forms.ComboBox                          ' Neue ComboBox erstellen
            LblBox = New System.Windows.Forms.Label                             ' Neues Lable erstellen
            CmbBoxWidth = GrpBox.Width - (2 * DefLeftPos) - LblWidth - 20
            LblBox.Name = "lbl" & szDatafield                                   ' Namen Setzen
            CmbBox.Name = "txt" & szDatafield                                   ' Namen Setzen
            LblBox.Location = New System.Drawing.Point(DefLeftPos, TopPos)
            LblBox.Size = New System.Drawing.Size(LblWidth, DefCtlHeight)       ' Grüße festlegen
            CmbBox.Location = New System.Drawing.Point(LblWidth + DefLeftPos, TopPos)
            CmbBox.Size = New System.Drawing.Size(CmbBoxWidth, DefCtlHeight)    ' Grüße festlegen
            CmbBox.Tag = szDatafield
            LblBox.Text = Caption                                               ' Bezeichnung
            CmbBox.Text = TextValue                                             ' Datenwert setzen
            AddHandler CmbBox.TextChanged, AddressOf Value_TextChanged          ' Handler binden
            LblBox.Anchor = AnchorStyles.Top Or AnchorStyles.Left               ' Verankern
            CmbBox.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right ' Verankern
            If Not IsNothing(InfoField) Then
                LblBox.Visible = InfoField.bVisible                             ' Sichtbar setzen
                CmbBox.Visible = InfoField.bVisible
                LblBox.Enabled = InfoField.bEnabled                             ' Enablen
                CmbBox.Enabled = InfoField.bEnabled
                'CmbBox.ReadOnly = InfoField.bLocked                             ' Eingaben sperren
                If InfoField.Valuelist <> "" Then
                    Call FillCMBlist(CmbBox, InfoField.Valuelist, ObjBag)
                ElseIf InfoField.ValueListSQL <> "" Then
                    Call FillCMBlist(CmbBox, InfoField.ValueListSQL, ObjBag)
                End If
            End If

            GrpBox.Controls.Add(CmbBox)                                         ' An Groupbox anhängen
            GrpBox.Controls.Add(LblBox)                                         ' An Groupbox anhängen
            Return True
        Catch ex As Exception                                                   ' Fehlerbehandlung
            Call ObjBag.ErrorHandler(MODULNAME, "AddValueText", ex)             ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    '     Public Function InitLabel(lblCTL As Label, szCaption As String, bNotEmpty As Boolean, _
    '            Optional lngTop As Integer)
    '        Try                                                                     ' Fehlerbehandlung aktivieren
    '            If bNotEmpty Then
    '                lblCTL.Caption = szCaption & " *"
    '            Else
    '                lblCTL.Caption = szCaption
    '            End If
    '            lblCTL.Height = lngCtlHeight
    '            lblCTL.Visible = True
    '            If lngTop > 0 Then
    '                lblCTL.Top = lngTop
    '            End If

    '        Catch ex As Exception                                                   ' Fehlerbehandlung
    '            Call ObjBag.ErrorHandler(MODULNAME, "InitLabel", ex)   ' Fehlermeldung ausgeben
    '            Return False                                                        ' Misserfolg zurück
    '        End Try
    '    End Function

    'Public Function InitTextBox(txtCTL As TextBox, szText As String, _
    '        Optional lngTop As Integer, _
    '        Optional szToolTip As String, _
    '        Optional szTag As String, _
    '        Optional Datassource As Object, _
    '        Optional szDatafield As String)
    ' Try                                                                     ' Fehlerbehandlung aktivieren
    '        If szDatafield <> "" And Not Datassource Is Nothing Then        ' Wenn Datafield und Datasource angegeben
    '            txtCTL.DataSource = Datassource                         ' Datasource setzen
    '            txtCTL.DataField = szDatafield                              ' Datafield setzen
    '        End If
    '        txtCTL.Text = szText                                            ' Text Setzen
    '        txtCTL.Height = lngCtlHeight                                    ' Höhe setzen
    '        txtCTL.ToolTipText = szToolTip                                  ' Tooltip setzen
    '        txtCTL.Tag = szTag                                              ' Tag Setzen
    '        If lngTop > 0 Then                                              ' Evtl. Top setzen
    '            txtCTL.Top = lngTop
    '        End If
    '        txtCTL.Visible = True                                           ' Anzeigen
    'Catch ex As Exception                                                   ' Fehlerbehandlung
    '            Call ObjBag.ErrorHandler(MODULNAME, "InitTextBox", ex)   ' Fehlermeldung ausgeben
    '            Return False                                                        ' Misserfolg zurück
    '        End Try
    '    End Function

    Private Function SetToolTips()
        Try                                                                     ' Fehlerbehandlung aktivieren
            'Call SetFormIcon(Me, ObjBag.MainForm.iltree, ObjValue.Imageindex, ObjBag)                         ' Form Icon setzen

            Return True
        Catch ex As Exception                                                   ' Fehlerbehandlung
            Call ObjBag.ErrorHandler(MODULNAME, "SetToolTips", ex)         ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function SetIcon()
        Try                                                                     ' Fehlerbehandlung aktivieren
            'Call SetFormIcon(Me, ObjBag.MainForm.iltree, ObjValue.Imageindex, ObjBag)                         ' Form Icon setzen

            Return True
        Catch ex As Exception                                                   ' Fehlerbehandlung
            Call ObjBag.ErrorHandler(MODULNAME, "SetIcon", ex)         ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function GetLockedControls()
        Try                                                                     ' Fehlerbehandlung aktivieren
            Return True
        Catch ex As Exception                                                   ' Fehlerbehandlung
            Call ObjBag.ErrorHandler(MODULNAME, "GetLockedControls", ex)         ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function GetDefaultValues()
        Try                                                                     ' Fehlerbehandlung aktivieren
            Return True
        Catch ex As Exception                                                   ' Fehlerbehandlung
            Call ObjBag.ErrorHandler(MODULNAME, "GetDefaultValues", ex)         ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

#End Region

    Private Function CheckDirty() As Boolean
        Try                                                                     ' Fehlerbehandlung aktivieren
            If Not IsNothing(ObjValue) Then
                If ObjValue.Dirty Then
                    lblDirty.Text = "ungespeichert"
                    btnSave.Enabled = True
                    Return True
                Else
                    lblDirty.Text = "gespeichert"
                    btnSave.Enabled = False
                    Return True
                End If
            Else
                Return False
            End If
        Catch ex As Exception                                                   ' Fehlerbehandlung
            Call ObjBag.ErrorHandler(MODULNAME, "GetDefaultValues", ex)         ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function UpdateEditForm()
        Dim dPos As Double
        Try                                                                     ' Fehlerbehandlung aktivieren
            With ObjValue.DataSet.Tables(0).Rows(0)
                For Each grp In Me.Controls
                    For Each ctl In grp.Controls
                        If ctl.tag <> "" Then
                            .Item(ctl.tag) = ctl.text
                        End If
                    Next
                Next
            End With
            Return True
        Catch ex As Exception                                                   ' Fehlerbehandlung
            Call ObjBag.ErrorHandler(MODULNAME, "UpdateEditForm", ex)           ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

#Region "Control Events"

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        If UpdateEditForm() Then
            If ObjValue.Save Then
                Me.Close()
            End If
        End If


    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If UpdateEditForm() Then
            If ObjValue.Save Then

            End If
        End If

        Call CheckDirty()
    End Sub

    Private Sub btnEsc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEsc.Click
        Me.Close()
    End Sub

#End Region

    Private Sub Value_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        ObjValue.Dirty = True
        Call CheckDirty()
    End Sub

End Class