VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Form1"
   ClientHeight    =   12660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20385
   LinkTopic       =   "Form1"
   ScaleHeight     =   12660
   ScaleWidth      =   20385
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnCopyNamesToTextBox 
      Caption         =   "Copy Names to TextBox >"
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   0
      Width           =   2415
   End
   Begin VB.CommandButton BtnFilterLocale 
      Caption         =   "Filter LOCALE"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton BtnGetConsts 
      Caption         =   "Get Constants"
      Height          =   375
      Left            =   7320
      TabIndex        =   4
      Top             =   0
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   12135
      Left            =   7320
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   3
      Top             =   360
      Width           =   9375
   End
   Begin VB.CommandButton BtnReadWinNlsh 
      Caption         =   "Read WinNls.h"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   12150
      Left            =   0
      MultiSelect     =   2  'Erweitert
      TabIndex        =   0
      Top             =   360
      Width           =   7215
   End
   Begin VB.Label Label1 
      Caption         =   " "
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_PFN As PathFileName
Private m_WinNlsh As ConstsWinNlsh

Private Sub Form_Load()
    Me.Caption = App.EXEName & " v" & App.Major & "." & App.Minor & "." & App.Revision
    'Set m_PFN = MNew.PathFileName("C:\Program Files (x86)\Microsoft SDKs\Windows\v7.1A\Include\WinNls.h")
    Set m_PFN = MNew.PathFileName("C:\Program Files\Microsoft Visual Studio\2022\Preview\SDK\ScopeCppSDK\vc15\SDK\include\um\WinNls.h")
End Sub

Private Sub Form_Resize()
    Dim L As Single, T As Single, W As Single, H As Single
    L = List1.Left
    T = List1.Top
    W = List1.Width
    H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then List1.Move L, T, W, H
    L = Text1.Left
    W = Me.ScaleWidth - L
    If W > 0 And H > 0 Then Text1.Move L, T, W, H
End Sub

Sub UpdateView()
    m_WinNlsh.ToLB List1
    Label1.Caption = m_WinNlsh.Constants.Count
End Sub

Private Sub List1_DblClick()
    Dim s As String: s = List1.List(List1.ListIndex)
    s = InputBox("Constant: ", "Copy constant info:", s)
End Sub

Private Sub BtnReadWinNlsh_Click()
    If Not m_PFN.Exists Then MsgBox "File not found: " & m_PFN.Value: Exit Sub
    Set m_WinNlsh = MNew.ConstsWinNlsh(m_PFN)
    Label1.Caption = m_WinNlsh.Constants.Count
    UpdateView
End Sub

Private Sub BtnFilterLocale_Click()
    If m_WinNlsh Is Nothing Then
        BtnReadWinNlsh_Click
    End If
    If m_WinNlsh Is Nothing Then
        MsgBox "No data maybe could not read winnls.h"
        Exit Sub
    End If

    Dim flt As String: flt = "LOCALE_"
    flt = InputBox("Filter every string starting with: ", "Filter", flt)
    If Len(flt) = 0 Then Exit Sub
    Set m_WinNlsh = m_WinNlsh.Filter(flt)
    UpdateView
End Sub

Private Sub BtnGetConsts_Click()
    Dim s    As String:  s = Text1.Text
    Dim sa() As String: sa = Split(s, vbCrLf)
    Dim cw As ConstsWinNlsh: Set cw = m_WinNlsh.FilterSA(sa)
    Dim v As Long, i As Long, c As CConstant
    Dim n As String: s = ""
    For i = 0 To UBound(sa) 'cw.Constants.Count - 1
        n = sa(i)
        If Not cw.Constants.ContainsKey(n) Then
            s = s & n & vbCrLf
        Else
            Set c = cw.Constants.ItemByKey(n)
            's = s & c.Key
            s = s & c.ToStr(0, vbTab) & vbCrLf
            'If c.TryGetValueAsLong(v) Then
            '    s = s & vbTab & "&H" & Hex(v)
            'End If
        End If
    Next
    Text1.Text = s
End Sub

Private Sub BtnCopyNamesToTextBox_Click()
Try: On Error GoTo Catch
    Dim s As String, sl As String, i As Long, c As Long
    Dim sa() As String
    ReDim cna(List1.ListCount - 1) As String
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) Then
            sl = List1.List(i)
            sa = Split(sl, " ")
            cna(c) = sa(0)
            c = c + 1
        End If
    Next
    ReDim Preserve cna(0 To c - 1)
    Text1.Text = Join(cna, vbCrLf)
Catch:
End Sub

