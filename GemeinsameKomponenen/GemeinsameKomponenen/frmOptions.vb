Imports Notarverwaltung.clsOptionList

Public Class frmOptions
    Private Const MODULNAME = "frmOptions"                                      ' Modulname für Fehlerbehandlung

    Private objBag As clsObjectBag                                              ' Sammelobject
    Private oClsOptions As clsOptionList
    Private oClsOptionsXML As clsXmlFile                                        ' Akt XML Doc mit Optionen
    Private bExpert As Boolean                                                  ' Experten einstellungen anzeigen

    Const DefLeftPos = 5
    Const DefCtlHeight = 20
    Const CtlDiff = 5
    Const GrpBoxDiff = 5

    Public Sub New(ByVal oBag As clsObjectBag, ByVal oOptions As clsOptionList, ByRef VarArray() As OptionValue)
        ' Dieser Aufruf ist für den Designer erforderlich.
        InitializeComponent()

        ' Fügen Sie Initialisierungen nach dem InitializeComponent()-Aufruf hinzu.
        objBag = oBag
        oClsOptions = oOptions
        oClsOptionsXML = oClsOptions.OptionsXML
        bExpert = True                                                          ' Zum Entwickeln erstmal true
        If InitOptionsForm() Then
            'If test() Then
        End If
    End Sub

    Private Function test() As Boolean
        Dim GrpBox As System.Windows.Forms.GroupBox                             ' GroupBox für Kategorie
        GrpBox = New System.Windows.Forms.GroupBox()
        GrpBox.Text = "Test1"
        GrpBox.Size = New System.Drawing.Size(400, 100)
        GrpBox.Location = New System.Drawing.Point(10, 10)

        Dim TxtBox As System.Windows.Forms.TextBox                              ' TextBox für Boolwerte
        TxtBox = New System.Windows.Forms.TextBox
        TxtBox.Text = "test 2"
        TxtBox.Size = New System.Drawing.Size(300, 20)
        TxtBox.Location = New System.Drawing.Point(10, 10)

        TxtBox.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        GrpBox.Controls.Add(TxtBox)                                         ' An Groupbox anhängen

        GrpBox.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

        Me.Controls.Add(GrpBox)                                 ' Broupbox ans fom hängen
        Return True
    End Function

    Private Function InitOptionsForm() As Boolean
        Dim OptionsRootNode As Xml.XmlElement
        Dim CategorieNode As Xml.XmlElement
        Dim OptionsNode As Xml.XmlElement
        Dim TopPosGrp As Integer = 5                                            ' Top Position für Groupboxen (wird hochgezählt)
        Dim TopPosCtl As Integer = 5                                            ' Top Position für Controls in Groupbox (wird Hochgezählt)
        Dim GrpHeight As Integer = 0                                            ' Höhe der akt Grop Box
        Dim FormHeight As Integer = 5                                           ' Höhe des forms
        Dim GrpBox As System.Windows.Forms.GroupBox                             ' GroupBox für Kategorie
        Try                                                                     ' Fehlerbehandlung aktivieren
            OptionsRootNode = oClsOptionsXML.RootElement                        ' Wurzelknoten ermittelm
            For i = 0 To OptionsRootNode.ChildNodes.Count - 1                   ' Alle kategorien durchlaufen
                CategorieNode = OptionsRootNode.ChildNodes(i)                   ' Akt. Kategorieknoten ermitteln
                If CategorieNode.HasChildNodes Then                             ' Nur wenn KAtegorie auch optionen hat
                    GrpBox = AddCategory(CategorieNode, TopPosGrp)              ' Pro Kategorie einen Frame laden
                    GrpHeight = 15                                              ' (platz vor dem 1. control)
                    TopPosCtl = 15
                    For n = 0 To CategorieNode.ChildNodes.Count - 1             ' Alle Optionen dieser Kategorie durchlaufen                        
                        OptionsNode = CategorieNode.ChildNodes(n)
                        ' hier noch prüfen ob überhaupt angezeigt                    
                        'If OptionsNode.GetAttribute("bEdit") Or bExpert Then    ' Nur Editierbare optionen anzeigen
                        If OptionsNode.GetAttribute("bEdit") Then    ' Nur Editierbare optionen anzeigen
                            If OptionsNode.GetAttribute("bBool") = "True" Then  ' Wenn Option ein Boolwert
                                Call AddOptionBool(GrpBox, OptionsNode, TopPosCtl) ' Dann Checkbox generieren
                                GrpHeight = GrpHeight + DefCtlHeight            ' Höhe GrpBox um Höhe Kindcontrol hochzälen
                                GrpHeight = GrpHeight + CtlDiff                 ' Höhe GrpBox um abstand der Kindcontrols hochzählen
                                TopPosCtl = TopPosCtl + DefCtlHeight + CtlDiff  ' Neue KindContro lHöhe bsetimmen
                            Else                                                ' Sonst Textbox
                                Call AddOptionText(GrpBox, OptionsNode, TopPosCtl) ' Text Box generieren
                                GrpHeight = GrpHeight + DefCtlHeight            ' Höhe GrpBox um Höhe Kindcontrol hochzälen
                                GrpHeight = GrpHeight + CtlDiff                 ' Höhe GrpBox um abstand der Kindcontrols hochzählen
                                TopPosCtl = TopPosCtl + DefCtlHeight + CtlDiff  ' Neue KindContro lHöhe bsetimmen
                            End If
                        End If
                    Next                                                        ' Nächste Option dieser Kategorie
                    If GrpBox.Controls.Count = 0 Then
                        GrpBox = Nothing
                    Else
                        GrpBox.Size = New System.Drawing.Size(Me.Width - (3 * DefLeftPos), GrpHeight) ' Groupbox Größe festlegen
                        GrpBox.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
                        TopPosGrp = TopPosGrp + GrpBox.Size.Height + GrpBoxDiff ' Neue TopPos für GrpBox bestimmen
                        Me.Controls.Add(GrpBox)                                 ' Groupbox ans fom hängen
                        FormHeight = FormHeight + GrpBox.Size.Height + GrpBoxDiff ' Fomular höhe hochzählen
                    End If
                End If
            Next                                                                ' Nächste Kategorie

            FormHeight = FormHeight + Me.btnEsc.Size.Height + 10
            If FormHeight < Me.MinimumSize.Height Then FormHeight = Me.MinimumSize.Height ' Minimum größe beachten
            If FormHeight > Me.Size.Height Then

            Else

            End If
            Dim btnTop As Integer = FormHeight - Me.btnEsc.Size.Height - 30     ' Top Pos. für Buttonleiste
            Me.btnEsc.Location = New System.Drawing.Point(Me.btnEsc.Location.X, btnTop) ' Pos. setzen
            Me.btnOK.Location = New System.Drawing.Point(Me.btnOK.Location.X, btnTop) ' Pos. setzen
            Me.btnSave.Location = New System.Drawing.Point(Me.btnSave.Location.X, btnTop) ' Pos. setzen
            Return True                                                         ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehlerbehandlung
            Call objBag.ErrorHandler(MODULNAME, "InitOptionsForm", ex)          ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function AddCategory(ByVal CategorieNode As Xml.XmlElement, _
                                 ByVal TopPos As Integer) As System.Windows.Forms.GroupBox
        Dim GrpBox As System.Windows.Forms.GroupBox
        Try                                                                     ' Fehlerbehandlung aktivieren
            GrpBox = New System.Windows.Forms.GroupBox()                        ' Pro Kategorie einen Frame laden
            'GrpBox.SuspendLayout()
            GrpBox.Name = "grp" & CategorieNode.GetAttribute("Name")            ' Namen setzen
            'GrpBox.TabIndex = i
            GrpBox.Text = CategorieNode.GetAttribute("Name")                    ' Text Setzen
            GrpBox.Location = New System.Drawing.Point(DefLeftPos, TopPos)      ' Pos. setzen
            'GrpBox.Visible = True

            Return GrpBox                                                       ' Control zurück
        Catch ex As Exception                                                   ' Fehlerbehandlung
            Call objBag.ErrorHandler(MODULNAME, "AddCategory", ex)              ' Fehlermeldung ausgeben
            Return Nothing                                                      ' Misserfolg zurück
        End Try
    End Function

    Private Function AddOptionBool(ByRef GrpBox As System.Windows.Forms.GroupBox, _
                                   ByVal OptionsNode As Xml.XmlElement, _
                                   ByVal TopPos As Integer) As Boolean
        Dim ChkBox As System.Windows.Forms.CheckBox                             ' Checkbox für Boolwerte
        Dim ChkBoxWidth As Integer                                              ' Checkbox breite
        Try                                                                     ' Fehlerbehandlung aktivieren
            ChkBoxWidth = GrpBox.Width - (2 * DefLeftPos)
            ChkBox = New System.Windows.Forms.CheckBox                          ' Neue Checkbox erstellen         

            ChkBox.Location = New System.Drawing.Point(DefLeftPos, TopPos)      ' Pos setzen (links 5 rechts TopPos)
            ChkBox.Name = "chk" & OptionsNode.GetAttribute("Name")              ' Namen setzen
            ChkBox.Tag = OptionsNode.GetAttribute("Name")                       ' Tag Setzen
            ChkBox.Size = New System.Drawing.Size(ChkBoxWidth, DefCtlHeight)    ' Grüße festlegen
            'ChkBox.TabIndex = n
            ChkBox.Text = OptionsNode.GetAttribute("Caption")                   ' Text = Options Caption
            ChkBox.Visible = True                                               ' Sichtbar
            ChkBox.Checked = CBool(oClsOptions.OptionByName(ChkBox.Tag).Value)  ' Wert setzen
            ChkBox.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right ' Verankern
            GrpBox.Controls.Add(ChkBox)                                         ' An Groupbox anhängen
            'ChkBox.PerformLayout()

            Return True
        Catch ex As Exception                                                   ' Fehlerbehandlung
            Call objBag.ErrorHandler(MODULNAME, "AddOptionBool", ex)            ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function AddOptionText(ByRef GrpBox As System.Windows.Forms.GroupBox, _
                                   ByVal OptionsNode As Xml.XmlElement, _
                                   ByVal TopPos As Integer) As Boolean
        Dim TxtBox As System.Windows.Forms.TextBox                              ' TextBox für Boolwerte
        Dim LblBox As System.Windows.Forms.Label                                ' Label für Caption
        Dim TxtBoxWidth As Integer
        Const LblWidth = 150
        Try                                                                     ' Fehlerbehandlung aktivieren
            TxtBox = New System.Windows.Forms.TextBox                           ' Neue Textbox erstellen
            LblBox = New System.Windows.Forms.Label                             ' Neues Lable erstellen
            TxtBoxWidth = GrpBox.Width - (2 * DefLeftPos) - LblWidth - 20
            LblBox.Name = "lbl" & OptionsNode.GetAttribute("Name")              ' Namen Setzen
            TxtBox.Name = "txt" & OptionsNode.GetAttribute("Name")              ' Namen Setzen
            LblBox.Location = New System.Drawing.Point(DefLeftPos, TopPos)
            LblBox.Size = New System.Drawing.Size(LblWidth, DefCtlHeight)       ' Grüße festlegen
            TxtBox.Location = New System.Drawing.Point(LblWidth + DefLeftPos, TopPos)
            TxtBox.Size = New System.Drawing.Size(TxtBoxWidth, DefCtlHeight)    ' Grüße festlegen
            TxtBox.Tag = OptionsNode.GetAttribute("Name")
            'TxtBox.TabIndex = n
            LblBox.Text = OptionsNode.GetAttribute("Caption")                   ' Text = Options Caption
            TxtBox.Text = oClsOptions.OptionByName(TxtBox.Tag).Value            ' Wert setzen
            LblBox.Anchor = AnchorStyles.Top Or AnchorStyles.Left               ' Verankern
            TxtBox.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right ' Verankern
            GrpBox.Controls.Add(TxtBox)                                         ' An Groupbox anhängen
            GrpBox.Controls.Add(LblBox)                                         ' An Groupbox anhängen
            'LblBox.PerformLayout()
            'TxtBox.PerformLayout()
            'If .bPath Then                              ' Option ist Verzeichnis
            '    f.cmdPath(c).Visible = .bPath           ' Button Pfadauswahl sichtbar
            '    f.cmdPath(c).Enabled = Not (.bDisabled) ' Evtl. diablen
            '    f.cmdPath(c).Tag = f.txtOption(c).Text  ' text im Button Tag speichern
            '    f.txtOption(c).Text = objTools.GetShortPath(f, CStr(f.cmdPath(c).Tag), _
            '            f.MaxDisplayPathLen)            ' Pfad anzeige kürzen
            'ElseIf .bFile Then                          ' Option ist Datei
            '    f.cmdFile(c).Visible = .bFile           ' Button Fileauswahl sichtbar
            '    f.cmdFile(c).Enabled = Not (.bDisabled) ' Evtl. diablen
            '    f.cmdFile(c).Tag = f.txtOption(c).Text  ' text im Button Tag speichern
            '    f.txtOption(c).Text = objTools.GetShortPath(f, CStr(f.cmdFile(c).Tag), _
            '            f.MaxDisplayPathLen)            ' Pfad anzeige kürzen
            'End If
            Return True
        Catch ex As Exception                                                   ' Fehlerbehandlung
            Call objBag.ErrorHandler(MODULNAME, "AddOptionText", ex)            ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function SaveOptions(Optional ByVal bInReg As Boolean = False) As Boolean
        ' Schreibt die Optionen zurück ins Array
        'Dim ctl As Control
        Try                                                                     ' Fehlerbehandlung aktivieren
            For Each grp As Control In Me.Controls                              ' Alle Groupboxen druchlaufen
                If Mid(grp.Name, 1, 3).ToUpper = "GRP" Then
                    For Each ctl In grp.Controls                                ' Alle Kindcontrols der Groupbox durchlaufen
                        If ctl.Tag <> "" Then                                   ' Wenn Ctl einen Tag hat
                            Select Case Mid(ctl.Name, 1, 3).ToUpper
                                Case "TXT"                                      ' Ist von uns erstellte Text box
                                    oClsOptions.OptionValueByName(ctl.Tag) = ctl.Text
                                Case "CHK"                                      ' Ist von uns erstellte Check box
                                    oClsOptions.OptionValueByName(ctl.Tag) = CBool(ctl.checked)
                                Case "GRP"

                                Case Else

                            End Select
                        End If
                    Next                                                        ' Nächstes Control
                End If
            Next                                                                ' Nächste groupbox
            If bInReg Then oClsOptions.SaveOptions() ' Optionen in registry speichern
            Return True
        Catch ex As Exception                                                   ' Fehlerbehandlung
            Call objBag.ErrorHandler(MODULNAME, "SaveOptions", ex)              ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Sub btnEsc_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEsc.Click
        Me.Close()                                                              ' Formular schliessem
    End Sub

    Private Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.Click
        If SaveOptions() Then                                                   ' Wenn Optionen zurückgeschrieben
            Me.Close()                                                          ' Form Schliessen
        End If
    End Sub

    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Call SaveOptions()                                                      ' Optionen zurückschreiben ins array
    End Sub
End Class