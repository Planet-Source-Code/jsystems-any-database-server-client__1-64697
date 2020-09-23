VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLBHScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Const LB_SETHORIZONTALEXTENT = &H194
Private Const WM_VSCROLL = &H115
Private Const SB_BOTTOM = 7
Private Const SM_CXVSCROLL = 2
Private mvarListBox As ListBox
Private m_lMaxItemWidth As Long

Public Sub Init(pListBox As ListBox)
Set mvarListBox = pListBox
mvarListBox.Clear
m_lMaxItemWidth = 0
End Sub

Public Sub AddItem(ByRef psItemText As String)
Dim m As Long
mvarListBox.AddItem psItemText
m = mvarListBox.Parent.TextWidth(psItemText)
If m > m_lMaxItemWidth Then
m_lMaxItemWidth = m

SendMessage mvarListBox.hwnd, LB_SETHORIZONTALEXTENT, m / Screen.TwipsPerPixelX + GetSystemMetrics(SM_CXVSCROLL) + 2, 0
End If

SendMessage mvarListBox.hwnd, WM_VSCROLL, SB_BOTTOM, 0
End Sub
