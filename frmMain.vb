Public Class frmMain

    Private Const MODULNAME = "frmMain"                                         ' Modulname für Fehlerbehandlung
    Private ObjBag As clsObjectBag                                              ' Sammelobject
    Private ObjCon As clsDBConnect                                              ' Datenbank Verbindungs klasse

    Private NavArrayForward(5) As String                                        ' Array für die Nav Schaltflächen Vorwärts
    Private NavArrayBack(5) As String                                           ' Array für die Nav Schaltflächen Zurück

    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If InitApplication() Then                                               ' Anwendungs Objecte initialisieren
            If ObjBag.InitDBClass() Then                                        ' Datenbank Klasse initiallisieren
                ObjCon = ObjBag.ObjDBConnect                                    ' Verbindungsobject abholen
                'ObjCon.BulidDBConnection(1, "OLG-SL-SRV-VM1", "", "", "", False) ' Test SQL Verbindung
                If Not ObjCon.BulidDBConnection(1, "OLG-SL-SRV-VM1", "Notare", "sa", "P!5nelke", False) Then ' Test SQL Verbindung
                    Call CriticalAppExit()
                End If
            End If
            If InitMainForm() Then                                              ' Hauptform Initialisieren
                Me.TVMain.SelectedNode = Me.TVMain.Nodes(0)                     ' Bis auf weiteres 1. knoten selecten
                Call RefreshNavList(Me.TVMain.SelectedNode.FullPath)
            End If
            Call ObjBag.ShowSplashForm(False)                                   ' Hier spätestens Splash ausbenden
        End If
    End Sub

    Private Sub frmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        'Stop
        Dim szPlitterPos As String = Me.SpliterMain.SplitterDistance.ToString
    End Sub

    Private Sub CriticalAppExit()
        Call MsgBox("Bei der Initialisierung der Anwendung ist ein Fehler aufgetreten. Die Anwendung wird beendet.", MsgBoxStyle.Critical, "Fehler")
        Me.Close()
    End Sub

#Region "initialisierungen"

    Private Function InitApplication() As Boolean
        Try                                                                     ' Fehlerbehandlung aktivieren
            ObjBag = New clsObjectBag(Me, Application.ProductVersion, Application.StartupPath())
            If Not ObjBag.InitOK Then
                Call CriticalAppExit()
                Return False
            End If
            Return True                                                         ' Hier ist alles in Ordnung
        Catch ex As Exception                                                   ' Fehler behandeln
            Call CriticalAppExit()
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function InitMainForm() As Boolean
        ' Initaialisiert / Positioniert Hauptform und dessen Steuerelemente
        Dim SpliterTop As Integer = 0
        Dim szWinState As String
        Dim szSize As String
        Try                                                                     ' Fehlerbehandlung aktivieren           
            Me.Text = ObjBag.oClsApp.AppTitle                                   ' Form Caption Setzten
            Call InitImageList()
            Call InitStatusStripMain()                                          ' Statusbar initialisieren
            Call InitMenue()                                                    ' Menü initiealisieren
            Call InitToolBar()                                                  ' Toolbar initialisieren

            'szSplitpos = objOptions.GetOptionByName(OPTION_SPLIT)               ' Spliter pos
            ' Me.SpliterMain.SplitterDistance ' muß hier gesetzt werden
            'If szSplitpos <> "" Then                                            ' Spliter pos. vorhanden
            '    curlngSplitposProz = CSng(szSplitpos)                           ' Spliter Pos setzen
            'Else
            '    curlngSplitposProz = 0.3                                        ' Default Spliter pos setzen
            'End If
            szSize = ObjBag.OptionByName(OPTION_MAINSIZE).Value                 ' Option WindowSize auslesen
            Call SetWindowSizeFromString(Me, szSize)                            ' WindowSize setzen
            szWinState = ObjBag.OptionByName(OPTION_MAINSTATE).Value            ' Option WindowState auslesen
            Call SetWindowStateFromString(Me, szWinState)                       ' Windowstate setzen

            If Me.MenuMain.Visible Then                                         ' wenn MainMenu nach Initialisierung sichtbar
                SpliterTop = Me.MenuMain.Height
            End If
            If Me.ToolStripMain.Visible Then                                    ' Wenn Toolbar nach initialisierung sichtbar
                SpliterTop = SpliterTop + Me.ToolStripMain.Height
            End If
            Me.SpliterMain.Location = New System.Drawing.Point(0, SpliterTop)
            Call InitLV()                                                       ' Listview Initialisieren
            Call InitTV()                                                       ' Treeview Initialisieren

            Return True                                                         ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "InitMainForm", ex)             ' Fehlermeldung
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function InitStatusStripMain() As Boolean
        Try                                                                     ' Fehlerbehandlung aktivieren


            Return True                                                         ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "InitStatusStripMain", ex)      ' Fehlermeldung
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function InitImageList()
        Dim img As Image
        Try                                                                     ' Fehlerbehandlung aktivieren
            For Each img In ILMain.Images                                       ' Alle bilder aus führende imagelist
                ILTree.Images.Add(img)                                          ' in TmageList Tree (kleine abbildungen) einlesen
            Next                                                                ' Nächstes Bild
            Return True                                                         ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "InitImageList", ex)            ' Fehlermeldung
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function InitLV() As Boolean
        ' List View Initialisieren
        Try                                                                     ' Fehlerbehandlung aktivieren
            LVMain.Sorting = SortOrder.None                                     ' Sortierung abschalten
            LVMain.Items.Clear()                                                ' Evtl. vorhandene Items löschen
            LVMain.Columns.Clear()                                              ' Evtl. vorhandene Spaltenköpfe löschen
            LVMain.FullRowSelect = True                                         ' Ganze Zeilen Selecten
            LVMain.View = View.Details                                          ' Erstmal detailansicht  
            LVMain.LargeImageList = ILMain                                      ' Imagelist für Große symbole
            LVMain.SmallImageList = ILTree                                      ' Imagelist für kleine symbole
            'LVMain.Icons = ILTree                                               ' Verweis auf Image List
            'LVMain.SmallIcons = ILTree
            'Call ShowLV                                                         ' Zuerst ListView anzeigen, Grid ausblenden

            Return True                                                         ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "InitLV", ex)                   ' Fehlermeldung
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function InitTV() As Boolean
        'Dim szRootNodeList As String                                            ' ; getrennte Liste mit Basisknoten
        Dim szRootNodeArray() As String                                         ' Rootnodelist als Array
        Try                                                                     ' Fehlerbehandlung aktivieren
            'szRootNodeList = ObjBag.oClsConfigXML.GetChildNodeNameList("Notarverwaltung\Tree", ";")
            'szRootNodeArray = Split(szRootNodeList, ";")                        ' Liste In Array Aufspalten
            szRootNodeArray = ObjBag.oClsConfigXML.GetChildNodeNameArray("Notarverwaltung\Tree")
            For i = 0 To szRootNodeArray.Length - 1                             ' das Arra Duchlaufen
                If szRootNodeArray(i) <> "" Then                                ' Wenn wir einen Namen haben
                    Dim tvn As New clsTreeNodeEx(ObjBag, szRootNodeArray(i))    ' Erweiterten TreeNode erstellen
                    Me.TVMain.Nodes.Add(tvn)                                    ' Hinzufügen
                    If tvn.ExTVNodeInfo.bShowSubnodes Then
                        tvn.Expand()
                    End If
                End If
            Next i                                                              ' Nächstes Array Item
            Return True                                                         ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "InitTV", ex)                   ' Fehlermeldung
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function InitMenue() As Boolean
        Dim xmlMainMenue As Xml.XmlElement                                      ' Xml Element des Main Menüs
        Dim xmlc As Xml.XmlElement                                              ' XML Element des jeweiligen Kindknotens (Menüpunkt)
        Dim Menu As ToolStripItem                                               ' Menüpunkt
        Try                                                                     ' Fehlerbehandlung aktivieren
            xmlMainMenue = ObjBag.oClsConfigXML.GetXmLNode("Notarverwaltung\Menue", "MainMenue")
            If xmlMainMenue.HasChildNodes Then                                  ' Nur wenn Kindknoten vorhanden
                For Each xmlc In xmlMainMenue.ChildNodes                        ' Alle Kindknoten durchlaufen
                    Menu = AddMenuItem(xmlc)                                    ' Menüpunkt erstellen (Rekursiv)
                    If Not IsNothing(Menu) Then                                 ' Wenn Menüpunkz forhanden
                        Me.MenuMain.Items.Add(Menu)                             ' Ans Menü Anhängen
                    End If
                Next                                                            ' Nächster Kindknoten
                Me.MenuMain.ImageList = Me.ILMain                               ' Imageliste festsetzen 
            Else
                Me.MenuMain.Visible = False                                     ' Menü ausblenden
            End If
            Return True                                                         ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "InitMenue", ex)                ' Fehlermeldung
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function InitToolBar() As Boolean
        Dim xmlMainMenue As Xml.XmlElement                                      ' Xml Element des Main Menüs
        Dim xmlc As Xml.XmlElement                                              ' XML Element des jeweiligen Kindknotens (Menüpunkt)
        Dim Menu As ToolStripItem                                               ' Menüpunkt
        Try                                                                     ' Fehlerbehandlung aktivieren
            xmlMainMenue = ObjBag.oClsConfigXML.GetXmLNode("Notarverwaltung\Menue", "ToolBar")
            If xmlMainMenue.HasChildNodes Then                                  ' Nur wenn Kindknoten vorhanden
                For Each xmlc In xmlMainMenue.ChildNodes                        ' Alle Kindknoten durchlaufen                   
                    Menu = AddMenuItem(xmlc, True)                              ' Menüpunkt erstellen (Rekursiv)
                    If Not IsNothing(Menu) Then                                 ' Wenn Menüpunke vorhanden
                        Me.ToolStripMain.Items.Add(Menu)                        ' An Toolbar Anhängen
                    End If
                Next                                                            ' Nächster Kindknoten
            Else
                Me.ToolStripMain.Visible = False                                ' Toolbar ausblenden
            End If
            Return True                                                         ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "InitToolBar", ex)              ' Fehlermeldung
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

#End Region

    Private Function AddMenuItem(ByVal xmlE As Xml.XmlElement, _
                                 Optional ByVal bToolbar As Boolean = False, _
                                 Optional ByVal bDiabled As Boolean = False) As ToolStripMenuItem
        Dim MIText As String
        Dim MIName As String
        Dim MIToolTip As String
        Dim szImageindex As String
        Dim MIEabled As String
        Dim xmlC As Xml.XmlElement
        Dim NewMenueItem As ToolStripMenuItem
        Try                                                                     ' Fehlerbehandlung aktivieren
            If IsNothing(xmlE) Then Return Nothing ' Kein XMLELEMENT -> Fertig
            If xmlE.HasChildNodes Then                                          ' Unterknoten = unter menüpunkte
                MIText = xmlE.GetAttribute("Text")                              ' Infos auslesen
                Dim NewMenue As New ToolStripMenuItem(MIText)                   ' MenüItem anlegen
                For Each xmlC In xmlE                                           ' Alle Unterknoten im XML durchlaufen
                    NewMenueItem = AddMenuItem(xmlC)                            ' Untermenüpunkte rekursiv dazu
                    If Not IsNothing(NewMenueItem) Then                         ' Wenn ergebnis nicht nothing
                        NewMenue.DropDownItems.Add(NewMenueItem)                ' Ans Menü anhängen
                    End If
                Next                                                            ' Nächster XML unterknoten
                Return NewMenue                                                 ' Menü zurückgeben
            Else                                                                ' Dies ist ein Menüpunkt 
                MIText = xmlE.GetAttribute("Text")                              ' Infos auslesen
                MIName = xmlE.GetAttribute("Name")
                szImageindex = xmlE.GetAttribute("Imageindex")
                MIToolTip = xmlE.GetAttribute("ToolTip")
                MIEabled = xmlE.GetAttribute("StartEnabled")
                If MIToolTip = "" Then MIToolTip = MIText
                NewMenueItem = New ToolStripMenuItem(MIText, Nothing, _
                                New EventHandler(AddressOf MainMenu_Click))     ' und anlegen
                NewMenueItem.Tag = MIName                                       ' Tag Setzen (wichtig für Click handler)
                NewMenueItem.Name = MIName
                NewMenueItem.ToolTipText = MIToolTip                            ' ToolTip setzen
                If MIEabled.ToUpper = "TRUE" Or MIEabled.ToUpper = "FALSE" Then
                    NewMenueItem.Enabled = CBool(MIEabled)
                End If
                If bToolbar Then
                    NewMenueItem.ImageAlign = System.Drawing.ContentAlignment.TopCenter
                    NewMenueItem.DisplayStyle = ToolStripItemDisplayStyle.Image
                Else
                    NewMenueItem.DisplayStyle = ToolStripItemDisplayStyle.ImageAndText
                End If
                If IsNumeric(szImageindex) Then
                    'NewMenueItem.ImageIndex = szImageindex
                    Dim img As Image
                    img = ILMain.Images(CType(szImageindex, Integer))
                    NewMenueItem.Image = img
                End If
                Return (NewMenueItem)                                           ' Menüpunkt zurück geben
                End If
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "AddMenuItem", ex)              ' Fehlermeldung
            Return Nothing                                                      ' Misserfolg zurück
        End Try
    End Function

    Private Function DoNav(ByVal bBack As Boolean) As Boolean
        ' Navigiert durch die Liste der gespeicherten Tree nodes        
        Dim cNode As clsTreeNodeEx                                              ' Akt TV Node
        Dim szKey As String                                                     ' Node key
        Dim NavArray() As String                                                ' Aktuelles Nav Array
        Dim i As Integer
        Try                                                                     ' Fehlerbehandlung aktivieren
            If bBack Then                                                       ' es wird Rückwärts Navigiert
                NavArray = NavArrayBack.Clone
                i = 1
            Else                                                                ' sonst wird Vorkwärts Navigiert
                NavArray = NavArrayForward.Clone
                i = 0
            End If
            szKey = NavArray(i)                                                 ' Key aus array ermitteln
            If szKey <> "" Then
                cNode = GetTreeNodeByKey(Me.TVMain, szKey, ObjBag)                  ' Node mit Key Ermitteln
                If IsNothing(cNode) Then Return False ' Kein Knoten -> Fertig
                'Call HandleNodeClick(cNode)                                         ' Click behandeln
                Call SelectTreeNode(TVMain, cNode)                                  ' Treenode auswählen      
            End If
            Call RefreshNavList(szKey, Not bBack)
            Return True                                                         ' Hier erfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "DoNav", ex)                    ' Fehlermeldung
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function RefreshNavList(ByVal szKey As String, Optional ByVal bForward As Boolean = True) As Boolean
        ' speichert szKey (NodePath im Tree) in array für Navigation
        Dim NavArrayCopy() As String
        Dim bForwardHasEntry As Boolean = False
        Dim bBackHasEntry As Boolean = False
        Try                                                                     ' Fehlerbehandlung aktivieren
            If szKey <> "" Then
                If bForward Then                                                ' Vorwärts navigiert (Durch navBar oder Node click)
                    NavArrayCopy = NavArrayBack.Clone                           ' Back Array Kopieren
                    If NavArrayBack(0) = szKey Then Return True
                    NavArrayBack(0) = szKey                                     ' Key anfügen (an pos 0)                
                    For i = 0 To NavArrayCopy.Length - 2
                        NavArrayBack(i + 1) = NavArrayCopy(i)
                    Next i                                                      ' Nächstes Array Item
                    NavArrayCopy = NavArrayForward.Clone
                    For i = 0 To NavArrayCopy.Length - 2
                        NavArrayForward(i) = NavArrayCopy(i + 1)
                    Next i                                                      ' Nächstes Array Item
                Else                                                            ' Rückwärts navigiert (nur durch navBar)
                    NavArrayCopy = NavArrayForward.Clone                        ' Forward Array Kopieren
                    NavArrayForward(0) = NavArrayBack(0)                        ' Key anfügen (an pos 0)
                    For i = 0 To NavArrayCopy.Length - 2
                        NavArrayForward(i + 1) = NavArrayCopy(i)
                    Next i                                                      ' Nächstes Array Item
                    NavArrayCopy = NavArrayBack.Clone                           ' Back Array Kopieren
                    For i = 0 To NavArrayCopy.Length - 2
                        NavArrayBack(i) = NavArrayCopy(i + 1)
                    Next i                                                      ' Nächstes Array Item                
                End If
            End If
            If (Not IsNothing(NavArrayBack(1))) Or NavArrayBack(1) <> "" Then bBackHasEntry = True
            If (Not IsNothing(NavArrayForward(0))) Or NavArrayForward(0) <> "" Then bForwardHasEntry = True
            Me.ToolStripMain.Items("mnuForward").Enabled = bForwardHasEntry     ' Vorwärts Einträge Vorhanden-> Entsprechenden Menüpunkt enablen    
            Me.ToolStripMain.Items("mnuBack").Enabled = bBackHasEntry           ' Rückwärts Einträge Vorhanden -> Entsprechenden Menüpunkt enablen
            Return True                                                         ' Ab hier erfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "RefreshNavList", ex)           ' Fehlermeldung
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

#Region "TestButtons"

    Private Sub btnShowNav_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnShowNav.Click
        Dim szNavList As String = "Nav Forward: " & vbCrLf
        For i = 0 To NavArrayForward.Length - 1
            szNavList = szNavList & "(" & i & ") " & NavArrayForward(i) & vbCrLf
        Next
        szNavList = szNavList & "Nav Back: " & vbCrLf
        For i = 0 To NavArrayBack.Length - 1
            szNavList = szNavList & "(" & i & ") " & NavArrayBack(i) & vbCrLf
        Next
        Call MsgBox(szNavList)
    End Sub

    Private Sub btnExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Call ObjBag.AskForExit()                                                ' Anwendung nach nachfrage beenden
    End Sub

    Private Sub btnEnviroment_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEnviroment.Click
        Call ObjBag.ShowEnvironment()                                           ' Alle Eviroment Werte anzeigen
    End Sub

    Private Sub btnOptions_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOptions.Click
        Call ObjBag.ShowOptions()                                               ' Alle Optionen Anzeigen
    End Sub

    Private Sub btnInfo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnInfo.Click
        Call ObjBag.ShowAboutForm()                                             ' InfoForm anzeigen
    End Sub

    Private Sub btnReadMe_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReadMe.Click

    End Sub

#End Region

#Region "Handels"

    Private Sub MainMenu_Click(ByVal sender As Object, ByVal e As EventArgs)
        Try                                                                     ' Fehlerbehandlung aktivieren
            'MsgBox(sender.tag)
            Select Case sender.tag
                Case "mnuBack"
                    Call DoNav(True)                                            ' Zurück navigieren
                Case "mnuForward"
                    Call DoNav(False)                                           ' Vorwärts navigieren
                Case "mnuExit"
                    Call ObjBag.AskForExit()                                    ' Anwendung nach nachfrage beenden
                Case "mnuAbout"
                    Call ObjBag.ShowAboutForm()                                 ' InfoForm anzeigen
                Case "mnuReadMe"

                Case "mnuOptions"
                    ObjBag.ShowOptionsForm()                                    ' Einstellungsform anzeigen
                Case "mnuDetail"
                    Me.LVMain.View = View.Details                               ' ListView ansicht auf Details
                Case "mnuSymbol"
                    Me.LVMain.View = View.LargeIcon                             ' ListView ansicht auf Große Symbole
                Case Else

            End Select
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "MainMenu_Click", ex)           ' Fehlermeldung
        End Try
    End Sub

    Private Sub HandleNodeClick(ByVal cNode As clsTreeNodeEx)
        Try                                                                     ' Fehlerbehandlung aktivieren
            Call HandleNodeSelect(cNode)                                        ' Navliste aktualisieren
            Call RefreshNavList(cNode.FullPath)
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "HandleNodeClick", ex)          ' Fehlermeldung
        End Try
    End Sub

    Private Sub HandleNodeSelect(ByVal cNode As clsTreeNodeEx)
        Try                                                                     ' Fehlerbehandlung aktivieren
            Me.LVMain.Clear()
            Call cNode.AddSubNodes()                                            ' evtl. Unterknoten anhängen
            Call RefreshListView(Me.LVMain, Me.TVMain, cNode, , ObjBag)         ' ListView aktualisieren
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "HandleNodeSelect", ex)         ' Fehlermeldung
        End Try
    End Sub

    Private Sub HandleListItemClick(ByVal lv As ListView)
        Dim szRootkey As String = ""                                            ' Identifiziert akt auswahl in Listview und treeview
        Dim szDetailKey As String = ""                                          ' Datensatz ID
        Dim szEditformName As String = ""                                       ' Evtl. Name des Bearbeiten formulars
        Try                                                                     ' Fehlerbehandlung aktivieren
            If lv.SelectedItems.Count > 0 Then                                  ' Ist überhaupt ein Listview Item ausgewählt
                'MsgBox(lv.SelectedItems(0).Tag)
                If EditPosible(LVMain) Then                                     ' Prüfen ob bearbeiten vorgesehen
                    Call GetRootKey(lv.SelectedItems(0), szRootkey, szDetailKey, szEditformName) ' DS ID & Editform namen eritteln
                    If szEditformName <> "" Then                                ' Wenn Editform angegeben
                        Call OpenEditForm(szEditformName, szDetailKey)          ' Bearbeiten Formular öffen
                    Else                                                        ' Sonst
                        Call OpenEditForm(szRootkey, szDetailKey)               ' Mit Rotkey versuchen
                    End If
                End If
            End If
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "HandleListItemClick", ex)      ' Fehlermeldung
        End Try
    End Sub

#End Region

    Private Function EditFormArrayCheck(ByVal FormName As String, _
                                        ByVal szID As String) As Form
        Try                                                                     ' Fehlerbehandlung aktivieren
            For Each oform In Application.OpenForms                             ' Auflistung Offener Formulare durchlafen
                If oform.name.toupper = FormName.ToUpper Then                   ' Ist ein Edit form offen
                    If oform.id = szID Then                                     ' Ist dessen DS ID = der gesuchten
                        Return oform                                            ' gefunden zurück
                    End If
                End If
            Next                                                                ' Nächstes Form
            Return Nothing                                                        ' Wenn wir hier ankommen haben wir nichts gefunden
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "EditFormArrayCheck", ex)       ' Fehlermeldung
            Return Nothing                                                        ' Misserfolg zurück
        End Try
    End Function

    Public Function OpenEditForm(ByVal szRootkey As String, _
                                 ByVal detailKey As String, _
                                 Optional ByVal ParentForm As Form = Nothing, _
                                 Optional ByVal bDialog As Boolean = False, _
                                 Optional ByVal Parameter As String = "") As Boolean
        ' Öffnet ein Edit Form und Trägt Ref in EditForm Array (Formsauflistung) ein
        Dim NewFrmEdit As Form                                                  ' Ref aufs neue EditForm
        Try                                                                     ' Fehlerbehandlung aktivieren
            NewFrmEdit = EditFormArrayCheck("frmEdit", detailKey)               ' Prüfen ob form mit dieser ID schon offen
            If Not IsNothing(NewFrmEdit) Then                                   ' Gefunden
                'NewFrmEdit.Show()                                               ' Anzeigen
                NewFrmEdit.BringToFront()                                       ' In Vorderrund schubsen
                Return True                                                     ' Erfolg zurück
            End If
            If ParentForm Is Nothing Then ParentForm = Me ' Akt. DB Form als Parent form
            'Select Case UCase(szRootkey)                                        ' Abhängig vom Rootkey
            '    Case UCase("Benutzer"), UCase("Richter"), UCase("Sonst. Mitarbeiter"), UCase("Rechtspfleger") ' Benutzer (a_personen)
            '        NewFrmEdit = New frmEditUser                                ' Neue Formular instanz
            '        Call NewFrmEdit.InitEditForm(ParentForm, ThisDBCon, detailKey, PArameter)
            '    Case UCase("Abteilungen"), UCase("Abteilung")
            '        NewFrmEdit = New frmEditAbt                                 ' Neue Formular instanz
            '        Call NewFrmEdit.InitEditForm(ParentForm, ThisDBCon, detailKey, PArameter)
            '    Case UCase("Register")
            '        NewFrmEdit = New frmEditReg                                 ' Neue Formular instanz
            '        Call NewFrmEdit.InitEditForm(ParentForm, ThisDBCon, detailKey, PArameter)
            '    Case UCase("Verfahrensgegenstände")
            '        NewFrmEdit = New frmEditVerfgegen                           ' Neue Formular instanz
            '        Call NewFrmEdit.InitEditForm(ParentForm, ThisDBCon, detailKey, PArameter)
            '    Case UCase("Parteibezeichnungen")
            '        NewFrmEdit = New frmEditBetBez                              ' Neue Formular instanz
            '        Call NewFrmEdit.InitEditForm(ParentForm, ThisDBCon, szRootkey, detailKey, PArameter)
            '    Case UCase("Vorinstanzparteibezeichnungen")
            '        NewFrmEdit = New frmEditBetBezVor                           ' Neue Formular instanz
            '        Call NewFrmEdit.InitEditForm(ParentForm, ThisDBCon, szRootkey, detailKey, PArameter)
            '    Case UCase("Zusatzparteibezeichnungen")
            '        NewFrmEdit = New frmEditBetBezZus                           ' Neue Formular instanz
            '        Call NewFrmEdit.InitEditForm(ParentForm, ThisDBCon, szRootkey, detailKey, PArameter)
            '    Case UCase("Zusatzvorinstanzparteibezeichnungen")
            '        NewFrmEdit = New frmEditBetBezVorZus                        ' Neue Formular instanz
            '        Call NewFrmEdit.InitEditForm(ParentForm, ThisDBCon, szRootkey, detailKey, PArameter)
            '    Case Else
            '        bDialog = True                                              ' Wertelisten Pauschal als Dialog aufrufen
            '        NewFrmEdit = New frmEdit                                    ' Neue Formular instanz
            '        Call NewFrmEdit.InitEditForm(ParentForm, ThisDBCon, szRootkey, detailKey, PArameter)
            'End Select
            NewFrmEdit = New frmEdit(ObjBag, szRootkey, detailKey)              ' Neue Formular instanz
            Call ObjBag.WriteProtokoll("OpenEditForm - RootKey: " & szRootkey & _
                    "  DetailKey: " & detailKey & " Params: " & Parameter)      ' Protolieren
            'EditFormArray(UBound(EditFormArray)) = NewFrmEdit                   ' Form ref Im Edit Form Array merken
            If bDialog Then                                                     ' Soll Modal geöffnet werden
                ''NewFrmEdit.Show 1, objObjectBag.getMainForm
                'NewFrmEdit.Show(1, objObjectBag.getmainform)                    ' Form modal anzeigen
                'OpenEditForm = NewFrmEdit.ID                                    ' ID aus form übernehmen (nur dialog)
                'Call CloseEditForm(NewFrmEdit)                                  ' Form Schliessen
            Else
                NewFrmEdit.Show()                                               ' Form anzeigen
            End If
            Return True
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "OpenEditForm", ex)             ' Fehlermeldung
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

#Region "Control Events"

    Private Sub TVMain_AfterSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TVMain.AfterSelect
        Call HandleNodeSelect(e.Node)
        'Dim n As clsTreeNodeEx
        'n = e.Node
        'Me.LVMain.Clear()
        'Call HandleNodeClick(n)
        'Call n.AddSubNodes()
        'Call RefreshNavList(n.FullPath)
        'Call RefreshListView(Me.LVMain, Me.TVMain, e.Node, , ObjBag)
    End Sub

    Private Sub TVMain_NodeMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles TVMain.NodeMouseClick
        Call HandleNodeClick(e.Node)
    End Sub

    Private Sub LVMain_ItemActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles LVMain.ItemActivate
        Call HandleListItemClick(sender)
    End Sub

#End Region

    Private Function EditPosible(ByVal LVMain As ListView) As Boolean
        Try                                                                     ' Fehlerbehandlung aktivieren
            Dim cNode As clsTreeNodeEx = GetTreeNodeByKey(TVMain, LVMain.Tag)
            If Not IsNothing(cNode) Then
                Return cNode.ExListInfo.bEdit
            End If
            Return False
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "EditPosible", ex)              ' Fehlermeldung
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function EditPosible(ByVal cNode As clsTreeNodeEx) As Boolean
        Try                                                                     ' Fehlerbehandlung aktivieren
            If Not IsNothing(cNode) Then
                Return cNode.ExListInfo.bEdit
            End If
            Return False
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "EditPosible", ex)              ' Fehlermeldung
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function GetRootKey(ByVal lvItem As ListViewItem, _
                                ByRef szRootkey As String, _
                                ByRef szDetailKey As String, _
                                Optional ByRef EditForm As String = "") As Boolean
        ' Ermittelt aus Tag (z.b. ListView.SelectedItems.tag) den RootKey (name des Editforms) und den DetailKey (datensatzID)
        ' Zusätzlich wird aus dem LV.tag der akt Treeview Node ermittelt und daraus der alternative name des edit forms
        Dim szItemTagArray() As String                                          ' Array mit ListView Item Tag elementen
        Dim szLVTagArray() As String
        Try                                                                     ' Fehlerbehandlung aktivieren
            If lvItem.Tag = "" Then Return False
            If LVMain.Tag = "" Then Return False
            szItemTagArray = Split(lvItem.Tag, TV_KEY_SEP)                      ' Item Tag aufspalten
            szLVTagArray = Split(LVMain.Tag, TV_KEY_SEP)                        ' ListView Tag aufspalten
            If szLVTagArray.Length = szItemTagArray.Length Then                 ' Beide Arrays sind gleich lang

            ElseIf szLVTagArray.Length + 1 = szItemTagArray.Length Then         ' Das ItemArray ist um 1 länger als das LvTagAray
                szDetailKey = szItemTagArray(szItemTagArray.Length - 1)         ' das letzte Item  im ItemArray ist die DS ID
                szRootkey = szItemTagArray(szItemTagArray.Length - 2)
            End If
            Dim cNode As clsTreeNodeEx = GetTreeNodeByKey(TVMain, LVMain.Tag)
            If Not IsNothing(cNode) Then
                EditForm = cNode.ExListInfo.EditFormName
            End If
            Return True
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "GetRootKey", ex)               ' Fehlermeldung
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    'Public Function GetKontextRoot(ByVal bTV As Boolean, _
    '                               ByVal szRootkey As String, _
    '                               ByVal szDetailKey As String, _
    '                               Optional ByVal szAction As String = "")
    '    ' ermittelt Rootkey und DS ID sowie mögliche Aktionen aus den Kontext (ListView) und XML

    '    Dim szItemTagArray() As String                                  ' Array mit ListView Item Tag elementen
    '    Dim szItemKeyArray() As String                                  ' Array mit ListView Item Key elementen
    '    Dim szLVTagArray() As String                                    ' Array mit ListView Tag elementen
    '    Dim szItemTag As String                                         ' Tag des ListView Items (* statt ID)
    '    Dim szItemKey As String                                         ' Key des ListView Items (enthält ID)
    '    Dim szLVTag As String                                           ' Tag des Listviews
    '    Dim szDetails As String                                         ' Details für Fehlerbehandlung
    '    'Dim TVNode As TreeViewNodeInfo                                  ' Infos über TreeNode
    '    'Dim LVInfo As ListViewInfo                                      ' Infos über ListView
    '    Try                                                                     ' Fehlerbehandlung aktivieren
    '        If bTV Then                                                     ' Aus TreeView ermitteln
    '            szItemTag = TVMain.SelectedNode.Tag
    '            'szItemKey = TVDB.SelectedItem.Key
    '        Else                                                            ' Aus ListView ermitteln
    '            szItemTag = LVMain.SelectedItems(0).Tag
    '            'szItemKey = LVDB.SelectedItem.Key
    '        End If
    '        szLVTag = LVMain.Tag
    '        'szDetails = "LVTag: " & lvmain.Tag & vbCrLf & "SelectedItem.tag: " & lvmain.SelectedItem.Key
    '        'Err.Clear()
    '        'On Error GoTo Errorhandler                                          ' Fehlerbehandlung wieder aktivieren
    '        'With LVInfo                                                     ' ListViewInfo aus XML füllen
    '        '    Call objTools.GetLVInfoFromXML(App.Path & "\" & INI_XMLFILE, LVDB.Tag, _
    '        '            .szsql, .szTag, .szWhere, .lngImage, .bValueList, .bListSubNodes, .bEdit, .bNew, .bSelectNode, _
    '        '            , , , .DelFlagField, , .bDelete)
    '        '    If .bEdit Then                                              ' Edit zulässig
    '        szAction = "Edit"
    '        '    End If
    '        '    'If .bDelete Then szAction = szAction & "Delete"
    '        '    If .bSelectNode Then szAction = szAction & "SelectNode"
    '        '    'If (Not .bEdit) And .bSelectNode Then szAction = "SelectNode"  ' Select zulässig
    '        '    If Not .bEdit And Not .bSelectNode Then szAction = "NoAction"
    '        'End With
    '        'If szItemKey = "" And szItemTag = "" Then
    '        '    szRootkey = LVInfo.szTag
    '        '    Return True
    '        'Else

    '        'End If
    '        szItemTagArray = Split(szItemTag, TV_KEY_SEP)                   ' ListView Item Tag aufspalten
    '        szItemKeyArray = Split(szItemKey, TV_KEY_SEP)                   ' ListView Item Key aufspalten
    '        szLVTagArray = Split(szLVTag, TV_KEY_SEP)                       ' ListView Tag aufspalten
    '        If UBound(szItemKeyArray) = UBound(szItemTagArray) Then

    '            If szItemTagArray(UBound(szItemTagArray)) <> "*" Then       ' Case 3 SubNode in Valuelist
    '                ' Ubound(szItemKeyArray ) = Ubound(szItemTagArray) / key mit ID / tag mit *
    '                szAction = "SelectNode"
    '                szRootkey = szItemTagArray(UBound(szItemTagArray))
    '            Else                                                        ' Case 1 Detail SubNode in ListView
    '                ' Ubound(szItemKeyArray ) = Ubound(szItemTagArray) / key mit ID / tag mit *
    '                ' -> Select Node and/or Edit
    '                szDetailKey = szItemKeyArray(UBound(szItemKeyArray))    ' ID aus ListItem.Tag ermitten
    '                szRootkey = szItemKeyArray(UBound(szItemKeyArray) - 1)
    '                If szRootkey = "*" Then                                 ' Case 2a einzelner DetailsDS (nicht in Valuelist)
    '                    ' Ubound(szLVTagArray) = Ubound(szItemTagArray) / LVTag mit ID / Itemtag mit *
    '                    ' Ubound(szItemKeyArray ) <> Ubound(szItemTagArray)
    '                    szRootkey = szItemKeyArray(UBound(szItemKeyArray) - 2)
    '                End If
    '            End If
    '            ' case 4 Relation Deatilnode eines Detailnodes in Liste
    '            ' Ubound(szItemKeyArray ) = Ubound(szItemTagArray) / key mit ID / tag mit *
    '        Else                                                            ' Case 2 einzelner DetailsDS (Valuelist)
    '            ' Ubound(szLVTagArray) = Ubound(szItemTagArray) / LVTag mit ID / Itemtag mit *
    '            szDetailKey = szItemTagArray(UBound(szItemTagArray))        ' ID aus ListItem.Tag ermitten
    '            If UBound(szItemTagArray) > 0 Then
    '                szRootkey = szItemTagArray(UBound(szItemTagArray) - 1)
    '            Else
    '                szRootkey = szLVTagArray(UBound(szLVTagArray))
    '            End If
    '        End If

    '        If szRootkey = "*" Or szRootkey = "" Then szAction = "SelectNode"
    '    Catch ex As Exception                                                   ' Fehler behandeln
    '        Call ObjBag.ErrorHandler(MODULNAME, "GetKontextRoot", ex)             ' Fehlermeldung
    '        Return False                                                        ' Misserfolg zurück
    '    End Try
    'End Function

End Class

