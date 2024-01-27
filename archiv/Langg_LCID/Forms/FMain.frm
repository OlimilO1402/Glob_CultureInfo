VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8760
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnTestLcid 
      Caption         =   "TestLcid"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   1455
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
      Height          =   4560
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   8535
   End
   Begin VB.CommandButton BtnListAllLanggs 
      Caption         =   "List All Languages"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_lcids As Collection

'also see repo Glob_CultureInfo
Private Sub BtnListAllLanggs_Click()
    Set m_lcids = MLang.ListLCIDStrings
    Dim v, s As String
    For Each v In m_lcids
        s = v
        List1.AddItem s
    Next
End Sub

Private Sub BtnTestLcid_Click()
    
    Dim lcid1 As Long: lcid1 = CLng("&H" & Hex(Int(Rnd * 100))) ' & Hex(Int(Rnd * 256)))
    
    Dim sLcid As String: sLcid = MLang.LCID_ToRfc1766(lcid1)
    MsgBox "lcid1: &H" & Hex(lcid1) & " = " & sLcid
    
    Dim lcid2 As Long: lcid2 = MLang.Rfc1766String_ToLCID(sLcid)
    MsgBox "lcid2: &H" & Hex(lcid2) & " = " & sLcid
    
    sLcid = "de-de"
    lcid1 = MLang.Rfc1766String_ToLCID(sLcid)
    sLcid = MLang.LCID_ToRfc1766(lcid1)
    MsgBox "lcid1: &H" & Hex(lcid1) & " = " & sLcid
    
    sLcid = "en-us"
    lcid2 = MLang.Rfc1766String_ToLCID(sLcid)
    sLcid = MLang.LCID_ToRfc1766(lcid2)
    MsgBox "lcid2: &H" & Hex(lcid2) & " = " & sLcid
        
End Sub

