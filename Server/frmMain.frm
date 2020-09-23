VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data server"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   7515
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1290
      Left            =   0
      TabIndex        =   4
      Top             =   660
      Width           =   7515
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3180
      Width           =   1635
   End
   Begin MSComDlg.CommonDialog Co 
      Left            =   660
      Top             =   3060
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   315
      Left            =   6300
      TabIndex        =   2
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   900
      TabIndex        =   1
      Top             =   0
      Width           =   5295
   End
   Begin MSWinsockLib.Winsock W 
      Index           =   0
      Left            =   5280
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1440
      TabIndex        =   3
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Database:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CurID As Long
Private MaxID As Long
Private sndTxt As String
Private DB As Database
Private FS As String
Private LastSelect As String
Private LB As New CLBHScroll
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As RECT, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA = 48
Option Base 1

Private Sub PutFormInLowerRight(ByVal frm As Form, ByVal right_margin As Single, ByVal bottom_margin As Single)
Dim wa_info As RECT
    If SystemParametersInfo(SPI_GETWORKAREA, _
        0, wa_info, 0) <> 0 _
    Then
        frm.Left = ScaleX(wa_info.Right, vbPixels, vbTwips) - _
            Width - right_margin
        frm.Top = ScaleY(wa_info.Bottom, vbPixels, vbTwips) - _
            Height - bottom_margin
    Else
        frm.Left = Screen.Width - Width - right_margin
        frm.Top = Screen.Height - Height - bottom_margin
    End If
End Sub

Private Sub Command1_Click()
On Error GoTo l1
Co.ShowOpen
Text1.Text = Co.FileName
SaveSetting App.Title, "set", "db", Co.FileName

Data1.DatabaseName = Co.FileName
Data1.Refresh
Label2.Caption = "Data linked"

FS = Co.FileName
LB.AddItem "Data linked: " & FS
Set DB = OpenDatabase(FS)
SendTables

Exit Sub
l1:
MsgBox Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo l1
LB.Init List1
PutFormInLowerRight Me, 0, 0
W(0).LocalPort = 9456
W(0).Listen
MaxID = 0: CurID = 0
Me.Caption = "JSQL Server V 1.0 [LOCAL IP: " & W(0).LocalIP & "]"
LB.AddItem "Server ready for connections."
If GetSetting(App.Title, "set", "db") <> Empty Then
Text1.Text = GetSetting(App.Title, "set", "db")
Data1.DatabaseName = GetSetting(App.Title, "set", "db")
FS = GetSetting(App.Title, "set", "db")
Label2.Caption = "Data linked"
Set DB = OpenDatabase(FS)
LB.AddItem "Data linked: " & FS
End If

l1:
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
For i = 0 To MaxID
W(i).Close
Next i
Set DB = Nothing
Data1.Database.Close
End Sub

Private Sub W_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim aa As Long
MaxID = MaxID + 1      '
aa = CheckW
Load W(aa)
W(aa).LocalPort = 0
W(aa).Accept requestID
LB.AddItem "Client no. " & aa & " connected from " & W(aa).RemoteHostIP & "[" & W(aa).RemoteHost & "]"
SendTables
End Sub

Private Function CheckW() As Long
On Error GoTo l1
Dim i As Long, a As Variant
For i = 0 To MaxID + 1
    a = W(i).RemotePort
Next i
l1:
CheckW = i
End Function

Private Sub W_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim GetRecieve As String
W(Index).GetData GetRecieve
CurID = Index
LB.AddItem "msg: " & GetRecieve
Ertelmezes GetRecieve
End Sub

Private Sub Ertelmezes(ByVal ReData As String)
Dim SQLStr As String
On Error GoTo l1
Select Case Left(ReData, 5)
Case "sqlst"
SQLStr = (Split(ReData, "~~")(1))
If Left(UCase(Trim(SQLStr)), 6) = "SELECT" Then
Data1.RecordSource = SQLStr
LastSelect = SQLStr
Else
If LastSelect <> "" Then
DB.Execute SQLStr
Data1.RecordSource = LastSelect
SendTables
End If
End If
Data1.Refresh
MakeRec
DoEvents: DoEvents: DoEvents
DoEvents: DoEvents: DoEvents
W(CurID).SendData sndTxt
Case "dbpsw"
Set DB = OpenDatabase(FS, False, gnreadonly, ";pwd=" & (Split(ReData, "~~")(1)))
Data1.DatabaseName = FS
End Select
Exit Sub
l1:
DoEvents: DoEvents: DoEvents
W(CurID).SendData "errxx~~" & Err.Description & "~~"
End Sub

Private Sub MakeRec()
Dim txt As String
Dim FieldCount As Long
Dim RecCount As Long
Dim i As Long, ii As Long
Dim T As TableDef
Dim R As Recordset
Dim n As Integer
Dim F As Field
Dim FName() As String
n = 0
For Each R In Data1.Database.Recordsets
If R.Name = Data1.Recordset.Name Then
    For Each F In R.Fields
    n = n + 1
    ReDim Preserve FName(n)
    FName(n) = F.Name
    Next
    Exit For
End If
Next
FieldCount = n
If Data1.Recordset.RecordCount <> 0 Then
Data1.Recordset.MoveLast
End If
RecCount = Data1.Recordset.RecordCount
txt = "recva~~"
'#### fileds & rows counts
txt = txt & FieldCount & "~~" & RecCount & "~~"
'#######field names
For i = 1 To UBound(FName)
txt = txt & FName(i) & "~~"
Next i
'############records
If Data1.Recordset.RecordCount <> 0 Then
Data1.Recordset.MoveFirst
Do While Not Data1.Recordset.EOF
For i = 0 To UBound(FName) - 1
        txt = txt & Trim(Data1.Recordset(i).Value) & "~~"
Next i
Data1.Recordset.MoveNext
Loop
End If
sndTxt = txt
End Sub

Private Sub W_Close(Index As Integer)
On Error GoTo l1
If W(Index).State <> sckConnected Then
LB.AddItem "Disconected user no. " & Index
End If
l1:
End Sub

Private Sub SendTables()
On Error GoTo l1
Set DB = OpenDatabase(FS)
Data1.DatabaseName = FS
Data1.Refresh
Dim T As TableDef
Dim txtx As String
Dim txt As String
Dim n As Long
txtx = "table~~"
txt = ""
For Each T In Data1.Database.TableDefs
If Left(T.Name, 4) <> "MSys" Then
    txt = txt & T.Name & "~~"
    n = n + 1
End If
Next
txtx = txtx & n & "~~" & txt
For i = 1 To MaxID
On Error Resume Next
DoEvents: DoEvents: DoEvents
DoEvents: DoEvents: DoEvents
W(i).SendData txtx
Next i
Exit Sub
l1:
DoEvents: DoEvents: DoEvents
DoEvents: DoEvents: DoEvents
DoEvents: DoEvents: DoEvents

W(MaxID).SendData "errxx~~" & Err.Description & "~~"

End Sub
