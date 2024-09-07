VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Form1"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19230
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   19230
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   18360
      TabIndex        =   11
      Top             =   0
      Width           =   735
   End
   Begin VB.ComboBox CmbISO3 
      Height          =   375
      Left            =   11160
      TabIndex        =   10
      Text            =   "CmbISO3"
      Top             =   0
      Width           =   2295
   End
   Begin VB.ComboBox CmbLCIDVal 
      Height          =   375
      Left            =   15960
      TabIndex        =   9
      Text            =   "CmbLCIDVal"
      Top             =   0
      Width           =   2295
   End
   Begin VB.ComboBox CmbLCIDNam 
      Height          =   375
      Left            =   13560
      TabIndex        =   8
      Text            =   "CmbLCIDNam"
      Top             =   0
      Width           =   2295
   End
   Begin VB.CommandButton BtnTestELCID 
      Caption         =   "Test ELCID"
      Height          =   375
      Left            =   8160
      TabIndex        =   5
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton BtnGetMaxNameLen 
      Caption         =   "MaxNameLen"
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton BtnTestDefaults 
      Caption         =   "Test Defaults"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton BtnLCIDEnum 
      Caption         =   "lcid-Enum"
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton BtnListLCIDString 
      Caption         =   "List All Languages"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   10200
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   2
      Top             =   360
      Width           =   8175
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4785
      ItemData        =   "FMain.frx":1782
      Left            =   0
      List            =   "FMain.frx":1784
      TabIndex        =   0
      Top             =   360
      Width           =   10215
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'Kein
      Height          =   375
      Left            =   10320
      ScaleHeight     =   375
      ScaleWidth      =   735
      TabIndex        =   7
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_CultureInfos As Collection 'Of CultureInfo
Private m_ci    As CultureInfo
Private m_ISO3s As Collection

Private Sub Form_Load()
    Me.Caption = App.EXEName & " v" & App.Major & "." & App.Minor & "." & App.Revision
    FillCombos
End Sub

Sub FillCombos()
    FillCombosLCID
    FillComboISO3
End Sub

Private Sub Form_Resize()
    Dim l As Single, t As Single, w As Single, h As Single
    l = List1.Left: t = List1.Top: w = List1.Width
    h = Me.ScaleHeight - t
    If w > 0 And h > 0 Then List1.Move l, t, w, h
    l = Text1.Left: w = Me.ScaleWidth - l
    If w > 0 And h > 0 Then Text1.Move l, t, w, h
End Sub

Private Sub BtnListLCIDString_Click()
    'CreateCultureInfos
    'Set m_CultureInfos = MLocale.GetCultureInfos(Mode_EnumSystemLocale, MLocale.LCID_INSTALLED)
    Set m_CultureInfos = MLocale.GetCultureInfos(Mode_EnumSystemLocaleEx, MLocale.LOCALE_ALL)
    UpdateViewList
    'MFlags.ReadFlagPics
End Sub

Private Sub BtnLCIDEnum_Click()
    Text1.Text = MLocale.LCIDEnumToStr
End Sub

Private Sub BtnTestDefaults_Click()
    If m_ci Is Nothing Then Set m_ci = New CultureInfo
    Dim ciud As CultureInfo: Set ciud = m_ci.UserDefaultCulture
    Dim cidt As CultureInfo: Set cidt = m_ci.DefaultThreadCurrentCulture
    Dim ciiv As CultureInfo: Set ciiv = m_ci.InvariantCulture
    Dim cipa As CultureInfo: Set cipa = m_ci.Parent
    
    MsgBox ciud.Name & "  " & ciud.lcid & "  " & ciud.LCID_ToHex & "  " & "UserDefaultCulture"
    MsgBox cidt.Name & "  " & cidt.lcid & "  " & cidt.LCID_ToHex & "  " & "DefaultThreadCurrentCulture"
    MsgBox ciiv.Name & "  " & ciiv.lcid & "  " & ciiv.LCID_ToHex & "  " & "InvariantCulture"
    MsgBox cipa.Name & "  " & cipa.lcid & "  " & cipa.LCID_ToHex & "  " & "Parent"
    
End Sub

Private Sub BtnGetMaxNameLen_Click()
    MsgBox "MaxNameLen: " & GetMaxNameLen
End Sub

Private Sub BtnTestELCID_Click()
    
    Dim lcid As Long
    
    'lcid = MLang.MAKELANGID(7, &H400)
    
    'MsgBox Hex(lcid)
    
    'lcid = MLang.MAKELCID(lcid, 1)
    'MsgBox Hex(lcid)
    
    
    'Dim s As String
    
    'lcid = &H7
    's = MLang.LCID_ToLocaleName(lcid)
    'MsgBox s 'de-DE
    
    'lcid = &H10407
    's = MLang.LCID_ToLocaleName(lcid)
    'MsgBox s 'de
    
    'lcid = &H9
    's = MLang.LCID_ToLocaleName(lcid)
    'MsgBox s 'en
    
    'lcid = &H409
    's = MLang.LCID_ToLocaleName(lcid)
    'MsgBox s 'en-US
    
    
'VB.net-code:
'        Dim ci_de As Globalization.CultureInfo
'
'        ci_de = New Globalization.CultureInfo(&H7)
'        MsgBox (ci_de.Name) 'de
'
'        ci_de = New Globalization.CultureInfo(&H407)
'        MsgBox (ci_de.Name) 'de-DE
'
'        Dim ci_en As Globalization.CultureInfo
'
'        ci_en = New Globalization.CultureInfo(&H9)
'        MsgBox (ci_en.Name) 'en
'
'        ci_en = New Globalization.CultureInfo(&H409)
'        MsgBox (ci_en.Name) 'en-US

    'Dim lcid0 As ELCID:  lcid0 = lcid_FI
    'Dim sLcid As String: sLcid = MLang.LCID_ToRfc1766(lcid0)
    
    'MsgBox "LCID: " & lcid0 & " = " & MLCID.ELCID_ToStr(lcid0) & " = " & sLcid
    
'    Dim ci_de As CultureInfo
'
'    Set ci_de = MNew.CultureInfo(&H7)
'    MsgBox ci_de.Name
'
'    Set ci_de = MNew.CultureInfo(&H407)
'    MsgBox ci_de.Name
'
    Dim ci_en As CultureInfo

    Set ci_en = MNew.CultureInfo(&H9)
    MsgBox ci_en.Name

    Set ci_en = MNew.CultureInfo(&H409)
    
    MsgBox ci_en.Name
    MsgBox ci_en.ToStr
    
    MsgBox ci_en.DisplayName
    MsgBox ci_en.EnglishName
    
End Sub

Private Sub FillComboISO3()
    CmbISO3.Clear
    Dim ci As CultureInfo
    Dim Key As String
    Set m_ISO3s = New Collection
    If m_CultureInfos Is Nothing Then
        BtnListLCIDString_Click
    End If
    For Each ci In m_CultureInfos
        Key = ci.AbbrevCountryName
        If Not Col_Contains(m_ISO3s, Key) Then
            m_ISO3s.Add ci.AbbrevCountryName, ci.AbbrevCountryName
        End If
    Next
    MPtr.Col_Sort m_ISO3s
    'MsgBox m_ISO3s.Count & " different countries"
    Dim v, s As String
    For Each v In m_ISO3s
        s = v
        CmbISO3.AddItem s
    Next
    'Text1.Text = s
    Set Picture1.Picture = Nothing
End Sub

Private Sub CreateCultureInfos()
    Set m_CultureInfos = New Collection
    Dim lcid As Long, ci As CultureInfo
    Dim i As Long, j As Long, langnam As String
    For i = 0 To &HFF
        For j = 0 To &H10
            lcid = j * &H400 + i
            Set ci = MNew.CultureInfo(lcid)
            langnam = ci.LanguageName
            'If langnam <> "Sprachneutral" Then Debug.Print """" & langnam & """"
            If Len(langnam) Then
                'If Col_TryAddObject(m_CultureInfos, ci, langnam) Then
                If langnam <> "Sprachneutral" Then
                If Col_TryAddObject(m_CultureInfos, ci, ci.LCID_ToHex) Then
                    'OK
                End If
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
    'Dim ln As String: ln = ParseLanguageName(List1.List(i))
    Dim ll As String: ll = ParseName(List1.List(i))
    'If ln = "????? (?????)" Then Set m_ci = Nothing Else Set m_ci = m_CultureInfos.Item(ln)
    Set m_ci = m_CultureInfos.Item(ll)
    UpdateViewDetails
End Sub

Private Function ParseLCID(s As String) As String
    Dim sa() As String: sa = Split(s, " | ")
    ParseLCID = Trim(sa(1))
End Function

Private Function ParseName(s As String) As String
    Dim sa() As String: sa = Split(s, " | ")
    ParseName = Trim(sa(2))
End Function

Private Function ParseLanguageName(s As String) As String
    Dim sa() As String: sa = Split(s, " | ")
    ParseLanguageName = Trim(sa(3))
End Function

Sub UpdateViewDetails()
    Dim s As String
    If m_ci Is Nothing Then
        s = ""
    Else
        With m_ci
            Dim sISO3 As String: sISO3 = .AbbrevCountryName
            Dim llcid As Long: llcid = .lcid
            s = s & "LCID        : " & CStr(llcid) & "(d) = &H" & Hex(llcid) & vbCrLf
            s = s & "Name        : " & .Name & vbCrLf
            s = s & "Languagename: " & .LanguageName & vbCrLf
            s = s & .ToInfoStr & vbCrLf
        End With
    End If
    Text1.Text = s
    'Set Picture1.Picture = MFlags.Flag(sIso3)
    Set Picture1.Picture = GetPicFromRes(sISO3) 'LoadResPicture(sISO3, VBRUN.LoadResConstants.vbResBitmap)
End Sub

Function GetPicFromRes(ByVal sISO3 As String) As StdPicture
Try: On Error GoTo Catch
    If Len(sISO3) <> 3 Then Exit Function
    Set GetPicFromRes = LoadResPicture(sISO3, VBRUN.LoadResConstants.vbResBitmap)
Catch:
End Function

Sub FillCombosLCID()
    Dim s As String
    Dim i As Long
    CmbLCIDNam.Clear
    CmbLCIDVal.Clear
    For i = 0 To 65536
        s = MLCID.ELCID_ToStr(i)
        If Len(s) Then
            CmbLCIDNam.AddItem s
            CmbLCIDVal.AddItem "&H" & Hex(i)
        End If
    Next
End Sub

Private Sub CmbLCIDNam_Click()
    Dim s As String: s = CmbLCIDNam.Text
    If Len(s) = 0 Then Exit Sub
    Dim h As Long: h = MLCID.ELCID_Parse(s)
    CmbLCIDVal.Text = IIf(h = 0, "", "&H" & Hex(h))
    Dim i As Long
    Set m_ci = FilterCultureInfo_ByLCID(h, i)
    List1.ListIndex = i
End Sub
Private Sub CmbLCIDNam_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> KeyCodeConstants.vbKeyReturn Then Exit Sub
    CmbLCIDNam_Click
End Sub

Private Sub CmbLCIDVal_Click()
    Dim s As String: s = Trim(CmbLCIDVal.Text)
    If Len(s) = 0 Then Exit Sub
    Dim h As Long: h = CLng(s)
    If h = 0 Then
        CmbLCIDNam.Text = ""
    Else
        s = MLCID.ELCID_ToStr(h)
        If Len(s) = 0 Then Exit Sub
        CmbLCIDNam.Text = s
    End If
End Sub
Private Sub CmbLCIDVal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> KeyCodeConstants.vbKeyReturn Then Exit Sub
    CmbLCIDVal_Click
End Sub

Private Sub CmbISO3_Click()
    Dim s As String: s = LCase(CmbISO3.Text)
    If Len(s) = 0 Then Exit Sub
    Dim i As Long
    Set m_ci = FilterCultureInfo_ByAbbrevCountryName(s, i)
    List1.ListIndex = i
'    For i = 1 To m_CultureInfos.Count
'        Set ci = m_CultureInfos.Item(i)
'        If LCase(ci.AbbrevCountryName) = s Then
'            Text1.Text = ci.ToInfoStr
'        End If
'    Next
End Sub

Private Sub CmbISO3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> KeyCodeConstants.vbKeyReturn Then Exit Sub
    CmbISO3_Click
End Sub

Function FilterCultureInfo_ByAbbrevCountryName(sISO3 As String, ByRef i_out As Long) As CultureInfo
    Dim ci As CultureInfo
    Dim i As Long
    For i = 1 To m_CultureInfos.Count
        Set ci = m_CultureInfos.Item(i)
        If LCase(ci.AbbrevCountryName) = sISO3 Then Exit For
    Next
    i_out = i - 1
    Set FilterCultureInfo_ByAbbrevCountryName = ci
End Function

Function FilterCultureInfo_ByLCID(aLCID As ELCID, ByRef i_out As Long) As CultureInfo
    Dim ci As CultureInfo
    Dim i As Long, c As Long: c = m_CultureInfos.Count
    For i = 1 To c
        Set ci = m_CultureInfos.Item(i)
        If ci.lcid = aLCID Then Exit For
    Next
    If c < i Then Exit Function
    i_out = i - 1
    Set FilterCultureInfo_ByLCID = ci
End Function

