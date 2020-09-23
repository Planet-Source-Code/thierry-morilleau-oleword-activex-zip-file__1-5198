VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "A propos"
   ClientHeight    =   2430
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4305
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1677.229
   ScaleMode       =   0  'User
   ScaleWidth      =   4042.618
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2640
      TabIndex        =   0
      Top             =   2040
      Width           =   1260
   End
   Begin VB.Label Label3 
      Caption         =   "This component is freeware, you can use it as ever you need !"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "tmorilleau@france-mail.com"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Copyright 1998 Thierry Morilleau"
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   1410
      Left            =   120
      Picture         =   "About.frx":0000
      Top             =   240
      Width           =   1380
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   3718.645
      Y1              =   1242.392
      Y2              =   1242.392
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   112.686
      X2              =   3718.645
      Y1              =   1242.392
      Y2              =   1242.392
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOK_Click()
  Unload Me
End Sub

