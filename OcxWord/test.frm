VERSION 5.00
Object = "{6810544D-40F9-11D2-A6DE-00C06C770162}#216.0#0"; "TWord.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   ScaleHeight     =   1830
   ScaleWidth      =   3885
   StartUpPosition =   3  'Windows Default
   Begin wordOle.OleWord OleWord1 
      Left            =   120
      Top             =   15
      _ExtentX        =   767
      _ExtentY        =   767
      NomFichier      =   "C:\OcxWord\essai.doc"
      DatabaseName    =   "C:\OcxWord\Comm97.mdb"
      RecordSource    =   "Client"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Mailing"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
OleWord1.Merge
End Sub

