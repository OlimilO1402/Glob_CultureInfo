VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ListBox List1 
      Height          =   5130
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   6015
   End
   Begin VB.CommandButton BtnListLCIDString 
      Caption         =   "ListAllLanguages"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function VerLanguageNameW Lib "kernel32" (ByVal wLang As Long, ByVal pSzLang As Any, ByVal nSize As Long) As Long

Private Sub BtnListLCIDString_Click()
    Dim i As Long, j As Long
    Dim sBuf As String * 256
    Dim HashCol As New Collection
    Dim c As Long
    Dim sLen As Long
    Dim sLCItem As String
    Dim sH As String
    Dim sItem As String
    Dim lcid As Long
  
    For i = 0 To &HFF
        For j = 0 To &H10
            lcid = j * &H400& + i
            sLen = VerLanguageNameW(lcid, ByVal StrPtr(sBuf), 512)
            If sLen > 0 Then
                sLCItem = Left$(sBuf, sLen)
                If Not IsInList(HashCol, sLCItem) Then
                    sH = Hex$(lcid)
                    sH = String$(4 - Len(sH), "0") & sH
                    sItem = "&H" & sH & "   " & sLCItem
                    c = c + 1
                    sH = CStr(c)
                    sH = String$(3 - Len(sH), "0") & sH & "    "
                    sItem = sH & sItem
                    Call List1.AddItem(sItem)
                End If
            End If
        Next
    Next
End Sub

Private Function IsInList(aCol As Collection, strKey As String) As Boolean
    On Error GoTo Err1
    Call aCol.Add(strKey, strKey)
    If Err.Number = 0 Then Exit Function
Err1:
    IsInList = True
End Function
