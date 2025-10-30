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

    Public Sub New(ByVal oBag As clsObjectBag, ByVal oOptions As clsOptionList, ByRef VarArray() As OptionValue)
        ' Dieser Aufruf ist für den Designer erforderlich.
        InitializeComponent()

        ' Fügen Sie Initialisierungen nach dem InitializeComponent()-Aufruf hinzu.
        objBag = oBag
        oClsOptions = oOptions
        oClsOptionsXML = oClsOptions.OptionsXML
        bExpert = True                                                          ' Zum Entwickeln erstmal true
        If InitOptionsForm() Then

        End If
    End Sub

    Private Function InitOptionsForm() As Boolean
        Dim OptionsRootNode As Xml.XmlElement
        Dim CategorieNode As Xml.XmlElement
        Dim OptionsNode As Xml.XmlElement
        Dim TopPosGrp As Integer = 5                                            ' Top Position für Groupboxen (wird hochgezählt)
        Dim TopPosCtl As Integer = 5                                            ' Top Position für Controls in Groupbox (wird Hochgezählt)
        Dim GrpHeight As Integer = 0                                            ' Höhe der akt Grop Box
        Dim FormHeight As Integer                                               ' Höhe des forms
        Dim GrpBox As System.Windows.Forms.GroupBox                             ' GroupBox für Kategorie
        Try                                                                     ' Fehlerbehandlung aktivieren
            OptionsRootNode = oClsOptionsXML.RootElement                        ' Wurzelknoten ermittelm
            For i = 0 To OptionsRootNode.ChildNodes.Count - 1                   ' Alle kategorien durchlaufen
                CategorieNode = OptionsRootNode.ChildNodes(i)                   ' Akt. Kategorieknoten ermitteln
                Me.SuspendLayout()
                If CategorieNode.HasChildNodes Then                             ' Nur wenn KAtegorie auch optionen hat
                    GrpBox = AddCategory(CategorieNode, TopPosGrp)              ' Pro Kategorie einen Frame laden
                    GrpHeight = 0
                    TopPosCtl = 15
                    For n = 0 To CategorieNode.ChildNodes.Count - 1             ' Alle Optionen dieser Kategorie durchlaufen
                        OptionsNode = CategorieNode.ChildNodes(n)
                        ' hier noch prüfen ob überhaupt angezeigt
                        'If .bEdit Or (.bExpert And bExpert) Then                ' Nur bEdit = true anzeigen
                        If OptionsNode.GetAttribute("bEdit") Or bExpert Then    ' Nur Editierbare optionen anzeigen
                            If OptionsNode.GetAttribute("bBool") = "True" Then  ' Wenn Option ein Boolwert
                                Call AddOptionBool(GrpBox, OptionsNode, TopPosCtl) ' Dann Checkbox generieren
                                GrpHeight = GrpHeight + 20
                            Else                                                ' Sonst Textbox
                                Call AddOptionText(GrpBox, OptionsNode, TopPosCtl)
                                GrpHeight = GrpHeight + 20

                            End If
                            GrpHeight = GrpHeight + CtlDiff
                            TopPosCtl = TopPosCtl + 20 + CtlDiff
                        End If
                    Next                                                        ' Nächste Option dieser Kategorie
                    If GrpBox.Controls.Count = 0 Then
                        GrpBox = Nothing
                    Else
                        GrpBox.Size = New System.Drawing.Size(Me.Width - 15, GrpHeight + 15) ' Groupbox Größe festlegen
                        TopPosGrp = TopPosGrp + GrpBox.Size.Height + 10
                        Me.Controls.Add(GrpBox)                                 ' Broupbox ans fom hängen
                        'GrpBox.Anchor = CType((System.Windows.Forms.AnchorStyles.Left Or _
                        '           System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
                    End If
                End If
            Next                                                                ' Nächste Kategorie

            Me.ResumeLayout(True)
            Return True
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
            GrpBox.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
            GrpBox.Location = New System.Drawing.Point(DefLeftPos, TopPos)
            GrpBox.Name = "grp" & CategorieNode.GetAttribute("Name")
            'GrpBox.TabIndex = i
            GrpBox.Text = CategorieNode.GetAttribute("Name")
            
            GrpBox.Visible = True
            Return GrpBox
        Catch ex As Exception                                                   ' Fehlerbehandlung
            Call objBag.ErrorHandler(MODULNAME, "AddCategory", ex)              ' Fehlermeldung ausgeben
            Return Nothing                                                      ' Misserfolg zurück
        End Try
    End Function

    Private Function AddOptionBool(ByVal GrpBox As System.Windows.Forms.GroupBox, _
                                   ByVal OptionsNode As Xml.XmlElement, _
                                   ByVal TopPos As Integer) As Boolean
        Dim ChkBox As System.Windows.Forms.CheckBox                             ' Checkbox für Boolwerte
        Try                                                                     ' Fehlerbehandlung aktivieren
            ChkBox = New System.Windows.Forms.CheckBox                          ' Neue Checkbox erstellen
            GrpBox.Controls.Add(ChkBox)                                         ' An Groupbox anhängen
            ChkBox.Location = New System.Drawing.Point(DefLeftPos, TopPos)               ' Pos setzen (links 5 rechts TopPos)
            ChkBox.Name = "chk" & OptionsNode.GetAttribute("Name")              ' Namen setzen
            ChkBox.Tag = OptionsNode.GetAttribute("Name")
            ChkBox.Size = New System.Drawing.Size(420, DefCtlHeight)            ' Grüße festlegen
            'ChkBox.TabIndex = n
            ChkBox.Text = OptionsNode.GetAttribute("Caption")                   ' Text = Options Caption
            ChkBox.Visible = True                                               ' Sichtbar
            ' Wert setzen
            ChkBox.Checked = CBool(oClsOptions.OptionByName(ChkBox.Tag).Value)
            Return True
        Catch ex As Exception                                                   ' Fehlerbehandlung
            Call objBag.ErrorHandler(MODULNAME, "AddOptionBool", ex)            ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function AddOptionText(ByVal GrpBox As System.Windows.Forms.GroupBox, _
                                   ByVal OptionsNode As Xml.XmlElement, _
                                   ByVal TopPos As Integer) As Boolean
        Dim TxtBox As System.Windows.Forms.TextBox                              ' TextBox für Boolwerte
        Dim LblBox As System.Windows.Forms.Label                                ' Label für Caption
        Try                                                                     ' Fehlerbehandlung aktivieren
            TxtBox = New System.Windows.Forms.TextBox                           ' Neue Textbox erstellen
            LblBox = New System.Windows.Forms.Label                             ' Neues Lable erstellen
            LblBox.Name = "lbl" & OptionsNode.GetAttribute("Name")              ' Namen Setzen
            TxtBox.Name = "txt" & OptionsNode.GetAttribute("Name")              ' Namen Setzen
            GrpBox.Controls.Add(TxtBox)                                         ' An Groupbox anhängen
            GrpBox.Controls.Add(LblBox)                                         ' An Groupbox anhängen
            LblBox.Location = New System.Drawing.Point(DefLeftPos, TopPos)
            LblBox.Size = New System.Drawing.Size(150, DefCtlHeight)            ' Grüße festlegen
            TxtBox.Location = New System.Drawing.Point(LblBox.Size.Width + DefLeftPos, TopPos)
            TxtBox.Size = New System.Drawing.Size(250, DefCtlHeight)            ' Grüße festlegen
            TxtBox.Tag = OptionsNode.GetAttribute("Name")
            'TxtBox.TabIndex = n
            LblBox.Visible = True                                               ' Sichtbar
            TxtBox.Visible = True                                               ' Sichtbar
            LblBox.Text = OptionsNode.GetAttribute("Caption")                   ' Text = Options Caption
            ' Wert setzen
            TxtBox.Text = oClsOptions.OptionByName(TxtBox.Tag).Value
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

End Class