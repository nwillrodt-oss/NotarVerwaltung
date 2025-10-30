Imports System.Xml                                                              ' XML Klasse Importieren (Spart schreibarbeit)
Imports Notarverwaltung.clsOptionList


Public Class clsObjectBag
    Inherits clsError                                                           ' Fehler Klasse beerben
    Private Const MODULNAME = "clsObjectBag"                                    ' Modulname für Fehlerbehandlung
    Private bInitOK As Boolean                                                  ' Gibt an das die Klasse erfolgreich initialisiert wurde
    Private cmdArray() As String                                                ' Array Mit Startparametern
    Private fSplash As frmSplash                                                ' Referenz auf den Splash Screen
    Private fMain As frmMain                                                    ' Referenz aufs Hauptform

    Public oClsApp As clsApp                                                    ' Anwendungs Klasse
    Public oClsUser As clsUser                                                  ' Benutzer Klasse
    Private oClsOptions As clsOptionList                                        ' Options Klasse
    Private oClsDBCon As clsDBConnect                                           ' optionale DB VerbindungsKlasse
    Public oClsConfigXML As clsXmlFile                                          ' XML DAte mit Anwendung konfiguration 

#Region "Constructor"

    Public Sub New(ByVal MainForm As Form)
        fMain = MainForm
        bInitOK = InitAppClass(Application.ProductVersion, Application.StartupPath)
    End Sub

    Public Sub New(ByVal MainForm As Form, ByVal szProduktversion As String, ByVal szAppPath As String)
        fMain = MainForm
        bInitOK = InitAppClass(szProduktversion, szAppPath)
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

#End Region

#Region "Properties"

    Public ReadOnly Property InitOK As Boolean                                  ' Gibt zurück ob die initialieierung Fehlerfrei war
        Get
            Return bInitOK
        End Get
    End Property

    Public ReadOnly Property ObjDBConnect() As clsDBConnect                     ' Gibt die DB Klasse zurück
        Get
            ObjDBConnect = oClsDBCon
        End Get
    End Property

    Public ReadOnly Property ConfigXML() As clsXmlFile
        Get
            ConfigXML = oClsConfigXML
        End Get
    End Property

    Public ReadOnly Property ConfigXMLDoc() As System.Xml.XmlDocument           ' Gibt das Config XML Zurück
        Get
            ConfigXMLDoc = oClsConfigXML.XMLDoc
            'If ConfigXML Is Nothing Then                                        ' XML noch nicht geladen
            '    Call LoadConfigXML()                                            ' XML Laden
            'End If
            'ConfigXMLDoc = ConfigXML                                            ' Und zurück geben
        End Get
    End Property

    Public ReadOnly Property Image(ByVal index As Integer) As Image
        Get
            Image = fMain.ILMain.Images(index)
        End Get
    End Property

    Public ReadOnly Property OptionByName(ByVal OptionName As String) As OptionValue
        Get
            OptionByName = oClsOptions.OptionByName(OptionName)
        End Get
    End Property

    Public ReadOnly Property MainForm() As frmMain
        Get
            MainForm = fMain
        End Get
    End Property

#End Region

    Private Function InitAppClass(ByVal szProduktversion As String, ByVal szAppPath As String) As Boolean
        Try                                                                     ' Fehlerbehandlung aktivieren
            Call ReadCmdParams()
            oClsApp = New clsApp(Me, szProduktversion, szAppPath)               ' Application Klasse mit ObjectBag initialieren
            'Error 1                                                             ' testweise Error raisen
            If Not oClsApp.InitOK Then                                          ' Initialisierung Fehlgeschlagen
                bInitOK = False
                Return False                                                    ' Misserfolg zurück
            End If
            oClsUser = New clsUser(Me)                                          ' Benutzer Klasse mit ObjectBag initialieren
            If Not oClsUser.InitOK Then                                         ' Initialisierung Fehlgeschlagen
                bInitOK = False
                Return False                                                    ' Misserfolg zurück
            End If
            oClsOptions = New clsOptionList(Me)                                 ' Options Klasse mit ObjectBag initialisieren
            If Not oClsOptions.InitOK Then                                      ' Initialisierung Fehlgeschlagen
                bInitOK = False
                Return False                                                    ' Misserfolg zurück
            End If
            Call ShowSplashForm(True)                                           ' Ab hier können wir den Splash anzeigen

            oClsConfigXML = New clsXmlFile(Me, oClsApp.AppDir & SZ_XMLFILE)
            If Not oClsConfigXML.InitOK Then                                    ' Initialisierung Fehlgeschlagen
                bInitOK = False
                Return False                                                    ' Misserfolg zurück
            End If

            Call CheckFirstStart()                                              ' Überprüfen ob 1. Start der Anwendung

            Return True                                                         ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehlerbehandlung
            Call ErrorHandler(MODULNAME, "InitAppClass", ex)                    ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Public Function InitDBClass(Optional ByVal bSQL As Boolean = True, _
                                Optional ByVal bAccess As Boolean = True)
        Try                                                                     ' Fehlerbehandlung aktivieren
            oClsDBCon = New clsDBConnect(Me, bAccess, bSQL)                     ' DBVerbindungs Klasse mit ObjectBag initialieren
            If Not oClsDBCon.InitOK Then                                        ' Initialisierung Fehlgeschlagen
                Return False                                                    ' Misserfolg zurück
            End If
            Return True                                                         ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehlerbehandlung
            Call ErrorHandler(MODULNAME, "InitDBClass", ex)                     ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Public Sub AppExit()
        Try                                                                     ' Fehlerbehandlung aktivieren
            ' Mainform Optionen in reg
            'fMain.Size
            ' Alle Forms schliessen

            ' Optionen speichern

        Catch ex As Exception                                                   ' Fehlerbehandlung
            Call ErrorHandler(MODULNAME, "AppExit", ex)                         ' Fehlermeldung ausgeben
        Finally
            Application.Exit()                                                  ' Anwendung beenden
        End Try
    End Sub

    Private Function ReadCmdParams() As Boolean
        ' Start Parameter auslesen
        ' Array mit startparametern   
        Dim separators As String = " -"                                         ' Seperator für Startparameter
        Dim szCmdString As String                                               ' Kompletter Start String 
        Try                                                                     ' Fehlerbehandlung aktivieren
            szCmdString = Microsoft.VisualBasic.Interaction.Command()           ' Startparameter einlesen
            szCmdString = szCmdString.ToUpper                                   ' hier nur Grossbuchstaben
            cmdArray = szCmdString.Split(separators.ToCharArray)                ' In Array Spliten
            'For i = 0 To UBound(cmdArray)                                   ' Ganzes Array abarbeiten
            '    Select Case UCase(cmdArray(i))
            '        Case UCase(CMD_NOSplash)                                    ' kein Splash
            '            If Not bE Then bNotShowSplash = True
            '        Case UCase(CMD_EGG)
            '            bNotShowSplash = False
            '            bE = Not bNotShowSplash
            '        Case UCase(CMD_DOS), UCase(CMD_CMD), UCase(CMD_CONSOLE)     ' Console
            '            bConsole = True
            '        Case UCase(CMD_AUTOCON)                                     ' Autoconnect
            '            bAutoConnect = True
            '        Case UCase(CMB_EXPERT)
            '            bExpert = True                                          ' Experten modus
            '        Case UCase(CMD_MORSE)
            '            bMorse = True
            '        Case UCase(CMB_LEET)
            '            bLEET = True
            '        Case UCase(CMD_HELP), UCase(CMD_HELP2)
            '            '            Call MsgBox(CMD_HELP_TXT, vbInformation, SZ_APPTITLE & " " _
            '            '                    & App.Major & "." & App.Minor & "." & App.Revision)
            '            End                                                      ' Anwendung beenden
            '        Case Else

            '    End Select
            'Next                                                            ' nächstes Array Item
            Return True                                                         ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehlerbehandlung
            Call ErrorHandler(MODULNAME, "ReadCmdParams", ex)
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function CheckFirstStart() As Boolean
        Try                                                                     ' Fehlerbehandlung aktivieren

            Return True
        Catch ex As Exception                                                   ' Fehlerbehandlung
            Call ErrorHandler(MODULNAME, "CheckFirstStart", ex)                 ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Public Sub ShowAboutForm()
        frmAbout.ShowDialog()                                                   ' Öffnet den Infodialog
    End Sub

    Public Sub ShowSplashForm(ByVal Visible As Boolean)
        Try                                                                     ' Fehlerbehandlung aktivieren

            If Not Visible Then
                If Not IsNothing(fSplash) Then
                    fSplash.Close()
                End If
            Else
                Dim Imagepath As String = oClsOptions.OptionByName(OPTION_SPLASH_IMG).Value
                fSplash = New frmSplash(Me, Imagepath)
                fSplash.Show()
            End If
            Threading.Thread.Sleep(3000)                                        ' 3 Sekunden warten   
        Catch ex As Exception                                                   ' Fehlerbehandlung
            Call ErrorHandler(MODULNAME, "ShowSplashForm", ex)                  ' Fehlermeldung ausgeben
        End Try
    End Sub

    Public Sub ShowOptionsForm()
        oClsOptions.ShowOptionsForm()
    End Sub

    Public Function AskForExit()
        Dim szTitle As String                                                   ' Msg Title
        Dim szMSG As String                                                     ' Msg Text
        Try                                                                     ' Fehlerbehandlung aktivieren
            szTitle = "Beenden"                                                 ' Meldungstitel festlegen
            szMSG = "Möchten Sie die " & oClsApp.AppTitle & " beenden?"         ' Meldung festlegen
            If ShowErrMsg(szMSG, vbOKCancel + vbQuestion, szTitle) <> vbCancel Then
                Call AppExit()                                                  ' Anwendung Beenden
            End If
            Return True
        Catch ex As Exception                                                   ' Fehlerbehandlung
            Call ErrorHandler(MODULNAME, "AskForExit", ex)                      ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

#Region "XML Geraller"

    'Private Function LoadConfigXML() As Boolean
    '    Try                                                                     ' Fehlerbehandlung aktivieren
    '        ConfigXML = New Xml.XmlDocument                                     ' Neues Object Createn
    '        ConfigXML.Load(oClsApp.AppDir & SZ_XMLFILE)                         ' XML Datei aus Anwendungsverz. Laden
    '        Return True                                                         ' Erfolg zurück
    '    Catch ex As Exception                                                   ' Fehler behandeln
    '        Call ErrorHandler(MODULNAME, "LoadOptionXML", ex)                   ' Fehlermeldung ausgeben
    '        Return False                                                        ' Misserfolg zurück
    '    End Try
    'End Function

    'Public Function GetTreeNodeFromXML(ByVal TVNodeFullPath As String) As XmlElement
    '    ' ermittelt aus <Root><Tree><Treenode> Childnode dessen Atrib. "Tag" dem Pfad NodeTag folgt
    '    Dim cXMLNode As XmlElement                                              ' HAuptknoten <Tree>
    '    Dim cXMLChildNode As XmlElement                                         ' Ergebis XML Node
    '    Dim FullPathArray() As String                                           ' Nodes Tag in Array augespalten
    '    Dim i As Integer                                                        ' Counter
    '    Try                                                                     ' Fehler behandlung aktivieren
    '        FullPathArray = Split(TVNodeFullPath, XML_PATH_SEP)                 ' FullPath aufspalten
    '        cXMLNode = GetXMLNode("Notarverwaltung" & XML_PATH_SEP & "Tree")    ' Hauptknoten <Notarverwaltung><Tree> auswählen
    '        cXMLChildNode = GetXmlChildNode(cXMLNode, "TreeNode", "Tag", FullPathArray(0))    ' Childnode mit Attribut "Tag"=TagArray(0) auswählen
    '        If FullPathArray.Length > 0 Then                                    ' Wenn Tag abgearbeitet Fertig
    '            For i = 1 To FullPathArray.Length - 1                           ' Array Abarbeiten
    '                cXMLChildNode = GetXmlChildNode(cXMLChildNode, "TreeNode", "Tag", FullPathArray(i))
    '            Next i
    '        End If
    '        Return cXMLChildNode                                                ' Gefundenen Knoten zurück geben
    '    Catch ex As Exception                                                   ' Fehler behandeln
    '        Call ErrorHandler(MODULNAME, "GetTreeNodeFromXML", ex)              ' Fehlermeldung ausgeben
    '        Return Nothing                                                      ' Misserfolg zurück
    '    End Try
    'End Function

#End Region

#Region "Debug Hilfen"

    Public Sub ShowEnvironment()
        ' Nur als entwiklungshilfe gedacht
        Dim szEnv As String
        szEnv = oClsApp.AppTitle & " v" & oClsApp.AppVersion & vbCrLf & oClsApp.AppDesc & vbCrLf & oClsApp.Copyright & vbCrLf _
            & oClsApp.AppDesc & vbCrLf _
            & "Host: " & oClsApp.Comutername & vbCrLf & "OS: " & oClsApp.Operatingsystem & vbCrLf & "WinDir: " & oClsApp.Windir & vbCrLf _
            & "Sys32Dir: " & oClsApp.Sys32Dir & vbCrLf & "RegRoot: " & oClsApp.RegRoot & vbCrLf _
            & "AppDir: " & oClsApp.AppDir & vbCrLf & "Images: " & oClsApp.ImageDir & vbCrLf & "Sounds: " & oClsApp.SoundDir & vbCrLf _
            & "User: " & oClsUser.UserName

        Call ShowErrMsg(szEnv, MsgBoxStyle.Information, "Environment")
    End Sub

    Public Sub ShowOptions()
        Dim szOpt As String = ""
        Dim i As Integer
        With oClsOptions
            For i = 0 To .Count
                szOpt = szOpt & .OptionByIndex(i).Caption & ": " & .OptionByIndex(i).Value & vbCrLf
            Next i
        End With
        Call ShowErrMsg(szOpt, MsgBoxStyle.Information, "Optionen")
    End Sub

#End Region

End Class


Public Class clsOptionList

    Private Const MODULNAME = "clsOptionList"                                   ' Modulname für Fehlerbehandlung
    Private ObjBag As clsObjectBag                                              ' Sammelklasse
    Private bInitOK As Boolean                                                  ' Gibt an das die Klasse erfolgreich initialisiert wurde
    Private oClsOptionsXML As clsXmlFile                                        ' Akt XML Doc mit Optionen
    Private VarOptions() As OptionValue                                         ' Optionen Array
    Private fOptions As frmOptions                                              ' Optionsform

    Public Structure OptionValue
        Public Name As String                                                   ' Name der Option
        Public Caption As String                                                ' Angez. Bezeichnung der option
        Public Value As Object                                                  ' Wert
        Public bCrypt As Boolean                                                ' Verschlüsselt
        Public bEdit As Boolean                                                 ' von Anw. editierbar
        Public bDisabled As Boolean                                             ' Wir angzeigt aber nicht änderbar
        Public bBool As Boolean                                                 ' Option ist ein Boolwert
        Public bPath As Boolean                                                 ' Option ist eine Pfadangabe
        Public bFile As Boolean                                                 ' Option ist eine Dateiangabe
        Public Kategorie As String                                              ' Kategoriebez. zum zusammenfassen
        Public bExpert As Boolean
        Public szList As String
    End Structure

#Region "Constructor"

    Public Sub New(ByVal oBag As clsObjectBag)
        Try                                                                     ' Fehlerbehandlung aktivieren
            ObjBag = oBag                                                       ' ObjectBag Übergeben
            If Not InitOptions() Then
                bInitOK = False
                Exit Sub
            End If
            bInitOK = True                                                      ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            bInitOK = False                                                     ' Misserfolg zurück
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

#End Region

#Region "Properties"

    Public ReadOnly Property InitOK As Boolean                                  ' Gibt zurück ob die initialieierung Fehlerfrei war
        Get
            Return bInitOK
        End Get
    End Property

    Public ReadOnly Property OptionsXML() As clsXmlFile
        Get
            OptionsXML = oClsOptionsXML
        End Get
    End Property

    'Public Property SaveOptionValueByName(ByVal OptionName As String)
    '    Set(ByVal value)

    '    End Set
    'End Property

    Public Property OptionValueByName(ByVal OptionName As String) As Object
        Get
            Return VarOptions(OptionGetIndexByName(OptionName)).Value
        End Get
        Set(ByVal value As Object)
            VarOptions(OptionGetIndexByName(OptionName)).Value = value
        End Set
    End Property

    Public ReadOnly Property OptionByName(ByVal OptionName As String) As OptionValue
        Get
            Return VarOptions(OptionGetIndexByName(OptionName))
        End Get
        'Set(ByVal value As VariantType)

        'End Set
    End Property

    Public ReadOnly Property OptionByIndex(ByVal Optionindex As String) As OptionValue
        Get
            Return VarOptions(Optionindex)
        End Get
        'Set(ByVal value As VariantType)

        'End Set
    End Property

    Public ReadOnly Property Count() As Integer
        Get
            Return UBound(VarOptions)
        End Get
    End Property

#End Region

    Public Sub ShowOptionsForm()
        Try                                                                     ' Fehlerbehandlung aktivieren
            fOptions = New frmOptions(ObjBag, Me, VarOptions)                   ' Einstellungs Form Laden
            fOptions.ShowDialog()                                               ' Als Dialog anzeigen
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "ShowOptionsForm", ex)          ' Fehlermeldung ausgeben
        End Try
    End Sub

    Public Function SaveOptions()
        ' Speichert die akt Optionen in die Registry
        Dim val As OptionValue
        Try                                                                     ' Fehlerbehandlung aktivieren
            If IsNothing(VarOptions) Then Return False ' Keine Optionen im array -> Fertig
            For i = 0 To VarOptions.Length - 1                                  ' Alle Array items durchlaufen
                val = OptionByIndex(i)                                          ' Option ermitteln
                With val
                    Call WriteOptionReg(.Name, .Value.ToString)                 ' Nach HKCU schreiben
                End With
            Next
            Return True                                                         ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "SaveOptions", ex)              ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function InitOptions() As Boolean
        ' Liest options aus ini (XML) in Array
        ' Schreibt diese in Reg (HKCU) wenn sie nicht existieren
        ' Übernimmt werte aus Reg wenn diese existieren
        ' Erst HKLM dann HKCU
        Dim OptionsRootNode As XmlElement                                       ' Options Wurzelknoten
        Dim CategorieNode As XmlElement                                         ' Akt. Kategorieknoten
        Dim i As Integer                                                        ' Counter für die Kategorieknoten
        Dim n As Integer                                                        ' Counter für Optionen einer Kategorie
        Try                                                                     ' Fehlerbehandlung aktivieren
            oClsOptionsXML = New clsXmlFile(ObjBag, ObjBag.oClsApp.AppDir & SZ_XMLOPTION)
            If oClsOptionsXML.InitOK Then
                'If LoadOptionXML() Then                                         ' Option Datei Laden
                OptionsRootNode = oClsOptionsXML.RootElement                    ' Wurzelknoten ermittelm
                For i = 0 To OptionsRootNode.ChildNodes.Count - 1               ' Alle kategorien durchlaufen
                    CategorieNode = OptionsRootNode.ChildNodes(i)               ' Akt. Kategorieknoten ermitteln
                    For n = 0 To CategorieNode.ChildNodes.Count - 1             ' Alle Optionen dieser Kategorie durchlaufen
                        Call InitOption(CategorieNode.ChildNodes(n))            ' Option einlesen
                    Next                                                        ' Nächste Option dieser Kategorie
                Next                                                            ' Nächste Kategorie
                Return True                                                     ' Erfolg zurück
            Else
                Return False
            End If

        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "InitOptions", ex)              ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function InitOption(ByVal OptiNode As XmlElement) As Boolean
        Dim OptValue As New OptionValue
        Try                                                                     ' Fehlerbehandlung aktivieren
            If GetOptionInfoFromXML(OptiNode, OptValue) Then                    ' Option aus XML Lesen
                If OptionGetIndexByName(OptValue.Name) < 0 Then                 ' Wenn diese option noch nicht vorhanden
                    'If CheckOptionArray(VarOptions) Then                        ' Prüfen ob Array nicht leer
                    If Not IsNothing(VarOptions) Then
                        ReDim Preserve VarOptions(VarOptions.Length)            ' Dann anhängen
                    Else                                                        ' Sonst
                        ReDim VarOptions(0)                                     ' Neues Arry
                    End If
                    VarOptions(VarOptions.Length - 1) = OptValue                ' Array Item Setzen
                End If
            End If
            Return True                                                         ' Erfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "InitOption", ex)               ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function CheckOptionArray(ByRef VarArray() As OptionValue) As Boolean
        ' prüft ob das Option Array schon einen Eintrag hat
        ' zur Fehlervermeidung
        Dim i As Integer                                                        ' Array größe
        Try
            i = UBound(VarArray)                                                ' Array Größe ermitteln
            Return True                                                         ' kein fehler ->  Array Definiert
        Catch ex As Exception
            Return False                                                        ' Array ist leer
        End Try
    End Function

    Private Function OptionGetIndexByName(ByVal szOptionName As String) As Integer
        ' Liest optionswert aus. Option wird nach name ausgewählt.
        Dim i As Integer                                                        ' Counter
        Try                                                                     ' Fehlerbehandlung aktivieren
            'If Not CheckOptionArray(VarOptions) Then
            If IsNothing(VarOptions) Then
                Return -1
            Else
                For i = 0 To VarOptions.Length - 1                              ' Alle optionen durchlaufen
                    If VarOptions(i).Name.ToUpper = szOptionName.ToUpper Then   ' Wenn gesuchter Optionname gefunden
                        Return i                                                ' Index auslesen
                    End If
                Next i                                                          ' Nächste Option
                Return -1
            End If
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "InitOptions", ex)              ' Fehler behandlung aufrufen
            Return -1
        End Try
    End Function

    Private Function GetOptionInfoFromXML(ByVal OptiNode As XmlElement, ByRef NewOptValue As OptionValue) As Boolean
        ' Liest attribute aus Options knoten
        Try                                                                     ' Fehlerbehandlung aktivieren
            With NewOptValue
                .Name = OptiNode.GetAttribute("Name")
                .Caption = OptiNode.GetAttribute("Caption")
                .Value = OptiNode.GetAttribute("Value")
                .bCrypt = OptiNode.GetAttribute("bCrypt")
                .bEdit = OptiNode.GetAttribute("bEdit")
                .bDisabled = OptiNode.GetAttribute("bDisabled")
                .bPath = OptiNode.GetAttribute("bPath")
                .bFile = OptiNode.GetAttribute("bFile")
                .bBool = OptiNode.GetAttribute("bBool")
                .bExpert = OptiNode.GetAttribute("bExpert")
                .szList = OptiNode.GetAttribute("List")
                If .bPath Or .bFile Then                                        ' Wenn Pfad angaben 
                    .Value = CheckSystemFolder(.Value, .bPath, .bFile)          ' überprüfen ob Dynamische pfadangaben enthalten sind
                End If

                If .bCrypt Then                                                 ' Falls Verschlüsselt
                    '.Value = objTools.Crypt(CStr(.Value), False)                ' Entschlüsseln
                End If
                If .bBool Then                                                  ' Boolwerte behandeln
                    If .Value = "" Then .Value = False ' Wenn Leer -> False
                    .Value = CBool(.Value)                                      ' In Boolwert wandeln
                End If
            End With
            Return True                                                         ' erfolg zurück
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "GetOptionInfoFromXML", ex)     ' Fehlermeldung ausgeben
            Return False                                                        ' Misserfolg zurück
        End Try
    End Function

    Private Function CheckSystemFolder(ByVal szValue As String, _
                                       Optional ByVal bDir As Boolean = False, _
                                       Optional ByVal bFile As Boolean = False) As String
        ' Überprüft das vorkommen von Schlüsselwörten in Pfad angaben
        ' und ersetzt diese durch real exitierende Pfad anaben
        Dim NewValuePart As String = ""
        Dim NewValue As String = ""
        Try                                                                     ' Fehlerbehandlung aktivieren
            If szValue = "" Then Return ""
            If InStr(szValue.ToUpper, DIR_EigeneDateien.ToUpper) > 0 Then       ' Eigene Dateien
                NewValuePart = ObjBag.oClsUser.PersonalDir
                NewValue = Replace(szValue, DIR_EigeneDateien, NewValuePart)
            End If
            If InStr(szValue.ToUpper, DIR_WinDir.ToUpper) > 0 Then              ' Windows Dir
                NewValuePart = ObjBag.oClsApp.Windir
                NewValue = Replace(szValue, DIR_WinDir, NewValuePart)
            End If
            If InStr(szValue.ToUpper, DIR_Sysdir32.ToUpper) > 0 Then            ' Windows/system32
                NewValuePart = ObjBag.oClsApp.Sys32Dir
                NewValue = Replace(szValue, DIR_Sysdir32, NewValuePart)
            End If
            If InStr(szValue.ToUpper, DIR_AppFolder.ToUpper) > 0 Then           ' Anwendungsverz.
                NewValuePart = ObjBag.oClsApp.AppDir
                NewValue = Replace(szValue, DIR_AppFolder, NewValuePart)
            End If
            If InStr(szValue.ToUpper, DIR_AppImages.ToUpper) > 0 Then           ' Imagefoder im Anwendungsverz
                NewValuePart = ObjBag.oClsApp.ImageDir
                NewValue = Replace(szValue, DIR_AppImages, NewValuePart)
            End If
            If NewValuePart <> "" Then                                          ' Wert zum ersetzen gefunden

            Else
                NewValue = szValue
            End If
            If InStr(3, NewValue, "\\") > 0 Then NewValue = Replace(NewValue, "\\", "\", 3)
            If NewValue <> "" Then                                              ' Neuer Wert gefunden
                If bDir Then                                                    ' Ist Direkotry
                    'If Not objTools.CheckPath(NewValue, False) Then NewValue = ""
                End If
                If bFile Then                                                   ' Ist File
                    'If Not objTools.ExistFile(NewValue, False) Then NewValue = ""
                End If
            End If
            Return NewValue                                                     ' Ergebniss liefern
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "CheckSystemFolder", ex)        ' Fehlermeldung ausgeben
            Return ""                                                           ' Misserfolg zurück
        End Try
    End Function

    Private Function ReadOptionReg(ByVal szOptName As String, Optional ByVal szDefaultValue As String = "") As String
        ' Liest Option Wert aus Registry
        ' erst HKLM dann HKCU
        Dim szValue As String                                                   ' Gefundener Reg Wert
        Try                                                                     ' Fehlerbehandlung aktivieren
            szValue = ReadRegValue("SOFTWARE\" & ObjBag.oClsApp.RegRoot, _
                                   szOptName, szDefaultValue, ObjBag)           ' Erst aus HKLM lesen
            Return szValue                                                      ' Wert zurück geben
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "ReadOptionReg", ex)            ' Fehler behandlung aufrufen
            Return ""
        End Try
    End Function

    Private Function WriteOptionReg(ByVal szOptName As String, Optional ByVal szValue As String = "")
        ' Schreibt Option Wert in Registry
        ' nur HKCU
        Try                                                                     ' Fehlerbehandlung aktivieren
            szValue = WriteRegValue("SOFTWARE\" & ObjBag.oClsApp.RegRoot, _
                                   szOptName, szValue, ObjBag)                  ' in HKCU schreiben
            Return szValue                                                      ' Wert zurück geben
        Catch ex As Exception                                                   ' Fehler behandeln
            Call ObjBag.ErrorHandler(MODULNAME, "WriteOptionReg", ex)           ' Fehler behandlung aufrufen
            Return ""
        End Try
    End Function
End Class