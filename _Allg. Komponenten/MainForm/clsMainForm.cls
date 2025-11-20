VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const MODULNAME = "clsMainForm"
Public objObjectBag As Object                                       ' ObjectBag object
Public objError  As Object                                          ' Glob Error object
Public objRegTools As Object                                        ' Registry Tools
Public objTools As Object                                           ' Hilfreiches
Public objSQLTools As Object                                        '
Public objOptions As Object                                         ' Optionen einlesen & Speichern

Private ThisFormOptions As FormOptions

Private Type FormOptions
    IsMDIContainer As Boolean                                       ' Gib an ob das Main Form ein MDI Container ist
    ThisForm As Form                                                ' Ref auf aktuelle Formular
    Caption As String                                               ' Form Caption
    Image As Integer                                                ' Imageindex des FormIcons
End Type

Private Sub Class_Initialize()
'

End Sub

Private Sub Class_Terminate()
'

End Sub

                                                                    ' *****************************************
                                                                    ' Properties
Public Property Get IsMDI() As Boolean
    Set IsMDI = ThisFormOptions.IsMDIContainer
End Property

'Public Property Set SetDBConn(DBCon As Object)
'    Set ThisDBCon = DBCon
'End Property
                                                                    ' *****************************************
                                                                    ' Methoden
Public Function InitMainForm(ObjectBag As Object, bMDI As Boolean)

    Set objObjectBag = ObjectBag                                    ' Übergebenen OBag übernehmen
                                                                    ' Sonstige benötigte Objekte holen
                                                                    
    Set objError = objObjectBag.GetErrorObj()                       ' Fehlerbehandlung
    Set objRegTools = objObjectBag.GetRegToolsObj()                 ' Registry Tools
    Set objTools = objObjectBag.GetToolsObj()                       ' Allg. Tools
    Set objSQLTools = objObjectBag.GetSqlToolsobj()
    Set objOptions = objObjectBag.GetOptionsObj()
    
    With ThisFormOptions
        .IsMDIContainer = bMDI
        If .IsMDIContainer Then
            Set .ThisForm = frmMDIMain
        Else
            Set ThisForm = frmMain
        End If
    End With
    
End Function

Public Function Show() As Boolean

    ThisFormOptions.ThisForm.Show
    
End Function

Public Function Hide() As Boolean

    ThisFormOptions.ThisForm.Hide
    
End Function
