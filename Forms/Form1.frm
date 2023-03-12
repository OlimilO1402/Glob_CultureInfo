VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Form1"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16110
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   16110
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text1 
      Height          =   4815
      Left            =   8040
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   2
      Top             =   360
      Width           =   8055
   End
   Begin VB.CommandButton BtnListLCIDString 
      Caption         =   "List All Languages"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2415
   End
   Begin VB.ListBox List1 
      Height          =   4785
      ItemData        =   "Form1.frx":1782
      Left            =   0
      List            =   "Form1.frx":1784
      TabIndex        =   0
      Top             =   360
      Width           =   8055
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_CultureInfos As Collection 'Of CultureInfo
Private m_ci As CultureInfo
'Private m_de As CultureInfo
'Private m_at As CultureInfo

Private Sub Form_Load()
    Me.Caption = App.EXEName & " v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Resize()
    Dim l As Single, T As Single, W As Single, H As Single
    l = List1.Left: T = List1.Top: W = List1.Width
    H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then List1.Move l, T, W, H
    l = Text1.Left: W = Me.ScaleWidth - l
    If W > 0 And H > 0 Then Text1.Move l, T, W, H
End Sub

Private Sub BtnListLCIDString_Click()
    CreateCultureInfos
    UpdateViewList
End Sub

Private Sub CreateCultureInfos()
    Set m_CultureInfos = New Collection
    Dim LCID As Long, ci As CultureInfo
    Dim i As Long, j As Long, langnam As String
    For i = 0 To &HFF
        For j = 0 To &H10
            LCID = j * &H400 + i
            Set ci = MNew.CultureInfo(LCID)
            langnam = ci.LanguageName
            If Len(langnam) Then
                If Col_TryAddObject(m_CultureInfos, ci, langnam) Then
                    'OK
                End If
            End If
        Next
    Next
End Sub

Private Sub UpdateViewList()
    Dim i As Long, ci As CultureInfo
    For Each ci In m_CultureInfos
        List1.AddItem Format(i, "000") & " | " & ci.ToStr
        i = i + 1
    Next
End Sub

Private Sub List1_Click()
    Dim i As Long: i = List1.ListIndex
    Dim ln As String: ln = ParseLanguageName(List1.List(i))
    If ln = "????? (?????)" Then Set m_ci = Nothing Else Set m_ci = m_CultureInfos.Item(ln)
    UpdateViewDetails
End Sub

Private Function ParseLanguageName(s As String) As String
    Dim sa() As String: sa = Split(s, " | ")
    ParseLanguageName = sa(3)
End Function

Sub UpdateViewDetails()
    Dim s As String
    If m_ci Is Nothing Then
        s = ""
    Else
        With m_ci
            Dim llcid As Long: llcid = .LCID
            s = s & "LCID        : " & CStr(llcid) & "(d) = &H" & Hex(llcid) & vbCrLf
            s = s & "Name        : " & .Name & vbCrLf
            s = s & "Languagename: " & .LanguageName & vbCrLf
            s = s & .ToInfoStr & vbCrLf
        End With
    End If
    Text1.Text = s
End Sub