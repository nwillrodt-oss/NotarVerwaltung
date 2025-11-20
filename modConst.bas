Attribute VB_Name = "modConst"
Option Explicit                                                     ' Variaben Deklaration erzwingen
Private Const MODULNAME = "modConst"                                ' Modulname für Fehlerbehandlung

Public Declare Sub ExitProcess Lib "kernel32" (ByVal _
        uExitCode As Long)
        
Public Declare Function GetTickCount Lib "kernel32" () As Long
                                                                    ' *****************************************
                                                                    ' Application Konstanten
Public Const SZ_APPTITLE = "Notarverwaltung"                        ' Anwendungstitel
Public Const SZ_REGROOT = "Notarverwaltung"                         ' Basis Registry Verz
Public Const SZ_COPYRIGHT = "© 2007 - 2011 OLG Schleswig"           ' Copyright text
Public Const INI_FILENAME = "Notare.ini"                            ' Ini konfigurationsfile (Alt)
Public Const INI_XMLFILE = "Notare.xml"                             ' XML konfigurationsfile
'Public Const INI_OPTIONSINI = "Options.ini"                         ' Options Konfigurations File
Public Const INI_OPTIONSINI = "Options.xml"                         ' Options Konfigurations File
Public Const SZ_SUPPORTMAIL = "mathias.wolter@olg.landsh.de"        ' Support Email
Public Const SZ_WWW = "" '"www.olg-schleswig.de"                    ' Internet adresse
Public Const SZ_HELPFILE = "NotarverwaltungHilfe.chm"               ' Hilfe datei
Public Const SZ_READMEFILE = "ReadMe.rtf"                           ' ReadMe Datei
Public Const TV_KEY_SEP = "\"                                       ' Trennzeichen für Tree Node Keys

                                                                    ' *****************************************
                                                                    ' Startparameter
Public Const CMD_NOSplash = "/nosplash"
Public Const CMD_EGG = "/egghead"
Public Const CMD_AUTOCON = "/autocon"
Public Const CMB_EXPERT = "/expert"
Public Const CMD_DOS = "/dos"
Public Const CMD_CMD = "/cmd"
Public Const CMD_MORSE = "/--"
Public Const CMD_CONSOLE = "/console"
Public Const CMB_LEET = "/1337"
Public Const CMD_TRANZ = "/tranzparent"
Public Const CMD_HELP = "/?"
Public Const CMD_HELP2 = "/Help"
                                                                    ' *****************************************
                                                                    ' Images Konstanten
Public Const IMG_MAIN = 1
Public Const IMG_COMP = 2
Public Const IMG_CPU = 4
Public Const IMG_OS = 5
Public Const IMG_NET = 6
Public Const IMG_SORTDOWN = 9
Public Const IMG_SORTUP = 10

Public Const SZ_EXPERT = "(Expertenmodus)"
                                                                    ' *****************************************
                                                                    ' Node Keys
Public Const SZ_TREENODE_MAIN = "Personen"

                                                                    ' *****************************************
                                                                    ' Optionen
Public Const OPTION_SPLASH = "Splash"
Public Const OPTION_SPLASH_IMG = "SplashImage"
Public Const OPTION_ABOUT_IMG = "AboutImage"
Public Const OPTION_AUTOCON = "AutoConnect"
Public Const OPTION_LASTCON = "LastConnection"
Public Const OPTION_CONN = "Connection"
Public Const OPTION_SPLIT = "SpliterPos"
Public Const OPTION_MAINSIZE = "MainWindowSize"
Public Const OPTION_MAINSTATE = "MainWindowState"
Public Const OPTION_DBSTATE = "DBWindowState"
Public Const OPTION_DBSIZE = "DBWindowSize"
Public Const OPTION_ERRLOG = "ErrorProtokollFile"
Public Const OPTION_APPLOG = "ProtokollFile"
Public Const OPTION_HELPFILE = "NetzMgrHelp.chm"
Public Const OPTION_TEMPLATES = "Vorlagenverzeichnis"
Public Const OPTION_ABLAGE = "Ablageverzeichnis"
Public Const OPTION_LASTNODE = "LastNode"
Public Const OPTION_STARTLASTNODE = "StartOnLastNode"
Public Const OPTION_SHOWLVBACK = "UseListViewBackground"
Public Const OPTION_LVBACK = "ListViewBackground"
Public Const OPTION_SHOWDELREL = "ShowDelRel"
Public Const OPTION_TRANZ = "Tranzparenz"
' MW 26.08.11 {
Public Const OPTION_DOCX = "DOCX"
Public Const OPTION_DOTX = "DOTX"
' MW 26.08.11 }
                                                                    ' *****************************************
                                                                    'INI Sections
Public Const INI_APP = "Application"
Public Const INI_SQL = "SQL"
Public Const INI_ROOTNODES = "MainNodeOrder"
Public Const INI_SUBNODES = "SubNodeOrder"
Public Const INI_VALUELIST = "ValueLists"
Public Const INI_RELATIONS = "Relations"
Public Const INI_NONEDIT = "NonEDIT"
Public Const INI_EDITSQL = "EditSQL"
Public Const INI_NONVIS = "NonVisible"
Public Const INI_NONEMPTY = "NonEmptyEDIT"
Public Const INI_DEFAULTVAL = "DefaultValues"
Public Const INI_VALIDATE = "ValidateValues"
Public Const INI_IMAGE = "IMAGEINDEX"
Public Const INI_SEARCH = "Search"

                                                                    ' *****************************************
                                                                    ' Protokoll
Public Const PROT_APP_START = "Application Start"
Public Const PROT_APP_END = "Application End"
Public Const PROT_DB_CON = "Connect"
Public Const PROT_DB_AUTOCON = "Autoconnect"
Public Const PROT_DB_DIS = "Disconnect"
Public Const PROT_DOC_START = "Dokument Automation Start"
Public Const PROT_DOC_END = "Dokument Automation End"
Public Const PROT_IMPORT_START = "Dokument Import Start"
Public Const PROT_IMPORT_END = "Dokument Import End"

                                                                    ' *****************************************
                                                                    ' ToolBarButtons Keys
Public Const TB_BACK = "tbBack"
Public Const TB_FORW = "tbForward"
Public Const TB_LEFT = "tbForward"
Public Const TB_RIGHT = "tbBack"
Public Const TB_NEW = "tbNew"
Public Const TB_NEWBEWERBER = "tbNewBewerber"
Public Const TB_NEWBEWERBUNG = "tbNewBewerbung"
Public Const TB_NEWSTELLE = "tbNewStelle"
Public Const TB_NEWAUSSCR = "tbNewAusschreibung"
Public Const TB_NEWDOC = "tbNewDokument"
Public Const TB_PRINT = "tbPrint"

'Public Const TB_CHECKDB = "tbCheck"
Public Const TB_REFRESH = "tbRefresh"
Public Const TB_SEARCH = "tbSearch"
Public Const TB_SEARCH_PERS = "tbSearchPerson"
Public Const TB_SEARCH_DOC = "tbSearchDoc"
Public Const TB_DOCNEW = "tbDocNew"
Public Const TB_HELP = "tbHelp"
Public Const TB_INFO = "tbInfo"
'Public Const TB_REFESH = "tbRefresh"

Public Function SignedToUnsignedLong(ByVal LongIn As Long) As Double
    If LongIn < 0 Then
        SignedToUnsignedLong = LongIn + 4294967296#
    Else
        SignedToUnsignedLong = LongIn
    End If
End Function

Public Function GetImageIndexByRootNode(szRootkey As String) As Integer

'    Select Case szRootKey
'    Case SZ_TREENODE_MAIN
'        GetImageIndexByRootNode = 1
'    Case SZ_TREENODE_REGISTERS, SZ_TREENODE_VERFGEGEN
'        GetImageIndexByRootNode = 3
'    Case SZ_TREENODE_USERS
'        GetImageIndexByRootNode = 5
'    Case SZ_TREENODE_UNITS
'        GetImageIndexByRootNode = 7
'    Case SZ_TREENODE_VERTRETER, SZ_TREENODE_VERTRITT
'        'GetImageIndexByRootNode = 6
'        GetImageIndexByRootNode = 6
'    Case SZ_TREENODE_PARTEI, SZ_TREENODE_VORPARTEI, SZ_TREENODE_ZUSPARTEI
'            GetImageIndexByRootNode = 14
'    Case SZ_TREENODE_SECURE
'            GetImageIndexByRootNode = 15
'    Case Else
'        GetImageIndexByRootNode = 0
'    End Select
    
End Function


