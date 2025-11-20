Attribute VB_Name = "modConst"
Option Explicit                                                     ' Variaben Deklaration erzwingen
Option Compare Text                                                 ' Sortierreihenfolge festlegen

Private Const MODULNAME = "modConst"                                ' Modulname für Fehlerbehandlung

                                                                    ' *****************************************
                                                                    ' Application Konstanten
Public Const SZ_APPTITLE = "DeCrypter"                              ' Anwendungstitel
Public Const SZ_COPYRIGHT = "© 2010 OLG Schleswig"                  ' Copyright text
Public Const INI_XMLFILE = "DeCryp.xml"                             ' XML konfigurationsfile
Public Const INI_OPTIONSINI = "Options.xml"                         ' Options Konfigurations File
Public Const SZ_SUPPORTMAIL = "mathias.wolter@olg.landsh.de"        ' Support Email
Public Const SZ_WWW = "" '"www.olg-schleswig.de"                    ' Internet adresse
Public Const SZ_HELPFILE = "" '"NotarverwaltungHilfe.chm"               ' Hilfe datei
Public Const SZ_READMEFILE = "ReadMe.rtf"                           ' ReadMe Datei
Public Const TV_KEY_SEP = "\"                                       ' Trennzeichen für Tree Node Keys

                                                                    ' *****************************************
                                                                    ' Startparameter
Public Const CMD_NOSplash = "/nosplash"                             ' Splash Verbergen
Public Const CMD_EGG = "/egghead"
Public Const CMD_AUTOCON = "/autocon"                               ' Autoconnect mit Last Connection
Public Const CMB_EXPERT = "/expert"                                 ' Experten einstellungen
Public Const CMD_DOS = "/dos"
Public Const CMD_CMD = "/cmd"
Public Const CMD_CONSOLE = "/console"
Public Const CMB_LEET = "/1337"
Public Const CMD_TRANZ = "/tranzparent"
Public Const CMD_WORK_PATH = "/wrkdir"                              ' Arbeitsverzeichnis
Public Const CMD_MORSE = "/--"
Public Const CMD_HELP = "/?"                                        ' Hilfe zu Startparametern
Public Const CMD_HELP2 = "/Help"                                    ' Hilfe zu Startparametern
'Public Const CMD_HELP_TXT = "Die Anwendung erlaubt folgende Parameter:" & vbCrLf & _
        "KompileHelper.exe [/nosplash /wrkdir DateiPfad ]" & vbCrLf & vbCrLf & _
        "/nosplash" & vbTab & "Ohne Splashscreen starten." & vbCrLf & _
        "/wrkdir [DateiPfad]" & vbTab & "Legt das Arbeitsverz. fest." & vbCrLf & _
        "/? oder /help" & vbTab & "Zeigt diese Hilfe an."

                                                                    ' *****************************************
                                                                    ' Optionen
Public Const OPTION_SPLASH = "Splash"                               ' Splash Anzeigen Ja/Nein
Public Const OPTION_SPLASH_IMG = "SplashImage"                      ' Splash Bild
Public Const OPTION_ABOUT_IMG = "AboutImage"                        ' About Bild
Public Const OPTION_AUTOCON = "AutoConnect"                         ' Autoconnect mit Last Connection
Public Const OPTION_LASTCON = "LastConnection"                      ' Letzte DB Verbindung
Public Const OPTION_CONN = "Connection"
Public Const OPTION_ERRLOG = "ErrorProtokollFile"                   ' Error Protokoll
Public Const OPTION_APPLOG = "ProtokollFile"                        ' Anwendungsprotokoll
Public Const OPTION_HELPFILE = "" ' "NetzMgrHelp.chm"
Public Const OPTION_TRANZ = "Tranzparenz"
Public Const OPTION_LASTPROT = "LastProtokoll"
Public Const OPTION_CHECKFIELDS = "CheckFields"
Public Const OPTION_CHECKIDIZES = "CheckIndizes"
Public Const OPTION_PROTUSER = "UserInProt"
Public Const OPTION_PROTCOMP = "CompInProt"
Public Const OPTION_PROTDATE = "DateInProt"
Public Const OPTION_OLDTABLES = "CheckOldTables"                    ' Auf nicht mehr benötigte Tabellen prüfen
Public Const OPTION_COUNTTABLES = "CountTables"                     ' DS in TAbellen zählen
Public Const OPTION_SCHEMA = "DBSchemaFile"                         ' Akt Shema Datei
Public Const OPTION_SCHEMADIR = "DBSchemaPath"                      ' Schema Verzeichnis
Public Const OPTION_CHECKPROTDIR = "ProtDir"                        ' Vergleichs protokoll
Public Const OPTION_PROTDIFF_ONLY = "DiffOnly"                      ' Nur abweichungen protokolieren
                                                                    ' *****************************************
                                                                    ' Protokoll
Public Const PROT_APP_START = "Application Start"                   ' Anwendung gestartet
Public Const PROT_APP_END = "Application End"                       ' Anwendung beendet
Public Const PROT_DB_CON = "Connect"                                ' Mit DB Verbunden
Public Const PROT_DB_AUTOCON = "Autoconnect"                        ' Automatisch verbunden mit LastConnection
Public Const PROT_DB_DIS = "Disconnect"                             ' Verbindung getrent
Public Const PROT_DOC_START = "Dokument Automation Start"
Public Const PROT_DOC_END = "Dokument Automation End"
Public Const PROT_IMPORT_START = "Dokument Import Start"
Public Const PROT_IMPORT_END = "Dokument Import End"

                                                                    ' *****************************************
                                                                    'INI Sections
Public Const INI_APP = "Application"
Public Const INI_SQL = "SQL"
Public Const INI_SQLNODE = "SQLNode"
Public Const INI_SQLLIST = "SQLList"
Public Const INI_ROOTNODES = "MainNodeOrder"
Public Const INI_VALUELIST = "ValueLists"
Public Const INI_RELATIONS = "Relations"
Public Const INI_NONEDIT = "NonEDIT"
Public Const INI_SUBNODES = "SubNodeOrder"
Public Const INI_EDITSQL = "EditSQL"
Public Const INI_NONVIS = "NonVisible"
Public Const INI_NONEMPTY = "NonEmptyEDIT"
Public Const INI_DEFAULTVAL = "DefaultValues"
Public Const INI_VALIDATE = "ValidateValues"
Public Const INI_IMAGE = "IMAGEINDEX"

' Images Konstanten
Public Const IMG_MAIN = 1






