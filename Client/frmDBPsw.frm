VERSION 5.00
Begin VB.Form frmDBPsw 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database password required:"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   4200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      Height          =   495
      Left            =   1620
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   2820
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   60
      Width           =   3135
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   60
      Picture         =   "frmDBPsw.frx":0000
      Top             =   60
      Width           =   480
   End
End
Attribute VB_Name = "frmDBPsw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
DBPassword = Text1.Text
Unload Me
End Sub

Private Sub Command2_Click()
DBPassword = ""
Unload Me
End Sub
