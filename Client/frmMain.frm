VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JSQL Client V 1.0"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10860
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   10860
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   4320
      Left            =   8460
      TabIndex        =   7
      Top             =   420
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   8520
      Top             =   5400
   End
   Begin VB.CommandButton CmdConnect 
      BackColor       =   &H0000C000&
      Caption         =   "Connect"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Text1"
      ToolTipText     =   "Enter Server IP Address"
      Top             =   60
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SEND"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   315
      Left            =   8700
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   900
      TabIndex        =   1
      Top             =   4920
      Width           =   7695
   End
   Begin MSWinsockLib.Winsock W 
      Index           =   0
      Left            =   7860
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ListView LV 
      Height          =   4395
      Left            =   60
      TabIndex        =   0
      Top             =   420
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   7752
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12648447
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "TABLES :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8520
      TabIndex        =   8
      Top             =   120
      Width           =   840
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   60
      Width           =   4515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SQL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   4980
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MaxCN As Long

Private Sub Command1_Click()
If Text1.Text <> "" Then
DoEvents: DoEvents: DoEvents
W(MaxCN).SendData "sqlst~~" & Text1.Text & "~~"
End If
End Sub

Private Sub Form_Load()
Text2.Text = W(MaxCN).LocalIP
LV.FullRowSelect = True
LV.View = lvwReport
LV.GridLines = True
LV.ColumnHeaders.Clear
LV.ListItems.Clear
End Sub

Private Sub CmdConnect_Click()
On Error GoTo SckErr
MaxCN = MaxCN + 1
Load W(MaxCN)
W(MaxCN).Connect Text2.Text, 9456
WinsockStatus
DoEvents: DoEvents: DoEvents
Exit Sub

SckErr:
Call WinsockStatus
CmdConnect.Enabled = True
End Sub

Private Sub List1_Click()
If List1.Text <> "" Then
DoEvents: DoEvents: DoEvents
W(MaxCN).SendData "sqlst~~" & "select * from " & List1.Text & "~~"
End If
End Sub

Private Sub W_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim GetRecieve As String
W(MaxCN).GetData GetRecieve
DoEvents: DoEvents: DoEvents
Ertelmez GetRecieve
End Sub

Sub WinsockStatus()
If W(MaxCN).State = sckConnected Then
Label2.Caption = "Connection Established"
CmdConnect.Enabled = False
Command1.Enabled = True
End If
If W(MaxCN).State = sckClosed Then
Label2.Caption = "Connection is Closed"
CmdConnect.Enabled = True
Command1.Enabled = False
LV.ListItems.Clear
List1.Clear
End If


If W(MaxCN).State = sckError Then
    Label2.Caption = "Network Sock Err"
    W(MaxCN).Close
    CmdConnect.Enabled = True
    Command1.Enabled = False
End If

If W(MaxCN).State = sckInProgress Then Label2.Caption = "Connection Open"
If W(MaxCN).State = sckConnecting Then Label2.Caption = "Connecting"

If W(MaxCN).State = 8 Then
    Label2.Caption = "Connection is Closed"
    W(MaxCN).Close
    CmdConnect.Enabled = True
    Command1.Enabled = False
End If

End Sub

Private Sub Timer1_Timer()
WinsockStatus
End Sub

Private Sub Ertelmez(ByVal ReData As String)
Select Case Left(ReData, 5)
Case "table" 'table list
List1.Clear
Dim nrT As Long
nrT = Val(Split(ReData, "~~")(1))
For i = 2 To nrT + 1
    List1.AddItem Split(ReData, "~~")(i)
Next i
Case "errxx" 'general error
If Split(ReData, "~~")(1) = "Not a valid password." Then
frmDBPsw.Show 1
If DBPassword <> "" Then
DoEvents: DoEvents: DoEvents
DoEvents: DoEvents: DoEvents
W(MaxCN).SendData "dbpsw~~" & DBPassword & "~~"
DoEvents: DoEvents: DoEvents
DoEvents: DoEvents: DoEvents
W(MaxCN).SendData "sqlst~~" & Text1.Text & "~~"
End If
Exit Sub
End If
MsgBox "Error on server:" & vbCrLf & Split(ReData, "~~")(1), vbCritical
Case "recva" 'values from database
Dim ColsC As Long, recC As Long
ColsC = Val((Split(ReData, "~~")(1)))
recC = Val((Split(ReData, "~~")(2)))

LV.ColumnHeaders.Clear
LV.ListItems.Clear

With LV.ColumnHeaders
For i = 1 To ColsC
    .Add , , Split(ReData, "~~")(i + 2), 2400
Next i
End With

Dim osszSzam As Single
Dim a As Long
a = 1
osszSzam = recC * ColsC
For i = 3 + ColsC To osszSzam + 3 + ColsC Step ColsC
On Error Resume Next
    With LV
    .ListItems.Add , , Split(ReData, "~~")(i)
    For ii = 1 To ColsC
    .ListItems(a).ListSubItems.Add , , Split(ReData, "~~")(i + ii)
    Next ii
    End With
    a = a + 1
Next i
End Select
End Sub
