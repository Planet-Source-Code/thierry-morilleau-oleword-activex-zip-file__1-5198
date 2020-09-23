VERSION 5.00
Begin VB.UserControl OleWord 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   MaskPicture     =   "UsrWord.ctx":0000
   Picture         =   "UsrWord.ctx":014A
   PropertyPages   =   "UsrWord.ctx":05CA
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "UsrWord.ctx":05FE
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2775
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   360
      Width           =   1260
   End
End
Attribute VB_Name = "OleWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
Private objetw As Object
Const m_def_Preview = True
Const m_def_NomFichier = ""
Dim m_NomFichier As String
Dim m_Preview As Boolean
Dim m_FileName As String
Dim m_DatabaseName As Database
Public Enum RecType
    Table = 0
    Dynaset = 1
    SnapShot = 2
End Enum

Public Sub ShowAbout()
Attribute ShowAbout.VB_Description = "Affiche la boîte de dialogue A propos de"
Attribute ShowAbout.VB_UserMemId = -552
    frmAbout.Show vbModal
    Unload frmAbout
    Set frmAbout = Nothing
End Sub


Private Sub UserControl_Resize()
    Size 430, 430
End Sub

Private Sub UserControl_Terminate()
    Set objetw = Nothing
End Sub
Public Property Get NomFichier() As String
Attribute NomFichier.VB_ProcData.VB_Invoke_Property = "PropertyPage1;Text"
Attribute NomFichier.VB_UserMemId = -520
    NomFichier = m_NomFichier
End Property

Public Property Let NomFichier(ByVal New_NomFichier As String)

    m_NomFichier = New_NomFichier
    PropertyChanged "NomFichier"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_NomFichier = m_def_NomFichier
    m_Preview = m_def_Preview
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim Index As Integer

    m_NomFichier = PropBag.ReadProperty("NomFichier", m_def_NomFichier)
    Data1.Connect = PropBag.ReadProperty("Connect", "Access")
    Data1.DatabaseName = PropBag.ReadProperty("DatabaseName", "")
    Data1.RecordsetType = PropBag.ReadProperty("RecordsetType", 1)
    Data1.RecordSource = PropBag.ReadProperty("RecordSource", "")
    m_Preview = PropBag.ReadProperty("Preview", m_def_Preview)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim Index As Integer

    Call PropBag.WriteProperty("NomFichier", m_NomFichier, m_def_NomFichier)
    Call PropBag.WriteProperty("Connect", Data1.Connect, "Access")
    Call PropBag.WriteProperty("DatabaseName", Data1.DatabaseName, "")
    Call PropBag.WriteProperty("RecordsetType", Data1.RecordsetType, 1)
    Call PropBag.WriteProperty("RecordSource", Data1.RecordSource, "")
    Call PropBag.WriteProperty("Preview", m_Preview, m_def_Preview)
End Sub

Public Function OuvrirFichier() As Variant
    Set objetw = CreateObject("Word.Application")
    objetw.Visible = True
    objetw.Documents.Open m_NomFichier
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Data1,Data1,-1,Connect
Public Property Get Connect() As String
Attribute Connect.VB_Description = "Indicates the source of an open database, a database used in a pass-through query, or an attached table."
Attribute Connect.VB_MemberFlags = "4"
    Connect = Data1.Connect
End Property

Public Property Let Connect(ByVal New_Connect As String)
    Data1.Connect() = New_Connect
    PropertyChanged "Connect"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Data1,Data1,-1,RecordsetType
Public Property Get RecordsetType() As RecType
Attribute RecordsetType.VB_Description = "Returns/sets a value indicating the type of Recordset object you want the Data control to create."
Attribute RecordsetType.VB_ProcData.VB_Invoke_Property = ";Data"
    RecordsetType = Data1.RecordsetType
End Property

Public Property Let RecordsetType(ByVal New_RecordsetType As RecType)
    Data1.RecordsetType() = New_RecordsetType
    PropertyChanged "RecordsetType"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Data1,Data1,-1,RecordSource
Public Property Get RecordSource() As String
Attribute RecordSource.VB_Description = "Returns/sets the underlying table, SQL statement, or QueryDef object for a Data control."
Attribute RecordSource.VB_ProcData.VB_Invoke_Property = "PropertyPage2;Data"
Attribute RecordSource.VB_MemberFlags = "4"
    RecordSource = Data1.RecordSource
End Property

Public Property Let RecordSource(ByVal New_RecordSource As String)
    Data1.RecordSource() = New_RecordSource
    PropertyChanged "RecordSource"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Data1,Data1,-1,DatabaseName
Public Property Get DatabaseName() As String
Attribute DatabaseName.VB_Description = "Returns/sets the name and location of the source of data for a Data control."
Attribute DatabaseName.VB_ProcData.VB_Invoke_Property = "PropertyPage3;Data"
Attribute DatabaseName.VB_MemberFlags = "4"
    DatabaseName = Data1.DatabaseName
End Property

Public Property Let DatabaseName(ByVal New_DatabaseName As String)
    Dim Msg
    Data1.DatabaseName() = New_DatabaseName
    If New_DatabaseName = "" Then
        On Error Resume Next
        Err.Clear
        Err.Raise vbObjectError + 1050, "oleWord", "Pas de base de données sélectionnée"
        If Err.Number <> 0 Then
         '   Msg = "Erreur # " & Str(Err.Number) & " was generated by " _
                    & Err.Source & Chr(13) & Err.Description
        Msg = Err.Description
        MsgBox Msg, , "Error", Err.HelpFile, Err.HelpContext
        End If
    End If
    PropertyChanged "DatabaseName"
End Property

'
Public Function Merge() As Variant
      Set objetw = CreateObject("Word.Application")
      If m_Preview = True Then
      objetw.Application.Visible = True
      End If
'ouverture du document
      objetw.Documents.Open m_NomFichier
'creation du mailing
      objetw.ActiveDocument.MailMerge.OpenDataSource Name:=DatabaseName, _
        ConfirmConversions:=False, ReadOnly:=False, LinkToSource:=True, _
        AddToRecentFiles:=False, PasswordDocument:="", PasswordTemplate:="", _
        WritePasswordDocument:="", WritePasswordTemplate:="", Revert:=False, _
        Format:=False, Connection:="TABLE " & RecordSource, SQLStatement:= _
        "SELECT * FROM [" & RecordSource & "]", SQLStatement1:=""
      objetw.ActiveDocument.MailMerge.EditMainDocument
      objetw.ActiveDocument.MailMerge.ViewMailMergeFieldCodes = False
 If m_Preview = False Then
      objetw.ActiveDocument.PrintOut
      objetw.Application.Quit 0
 End If
End Function

Public Property Get Preview() As Boolean
    Preview = m_Preview
End Property

Public Property Let Preview(ByVal New_Preview As Boolean)
    m_Preview = New_Preview
    PropertyChanged "Preview"
End Property
