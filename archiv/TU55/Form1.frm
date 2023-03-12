VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  
  MsgBox LCID_ToRfc1766(&H407) 'de
  MsgBox LCID_ToRfc1766(1031)  'de
  MsgBox LCID_ToRfc1766(&HC07) 'de-AT
  MsgBox LCID_ToRfc1766(&H7)   'de-DE
  MsgBox LCID_ToRfc1766(2077)  'sv-FI
  MsgBox LCID_ToRfc1766(&H1D)  'sv-SE
  MsgBox LCID_ToRfc1766(&H81D) 'sv-FI
  
  MsgBox Hex$(Rfc1766_ToLCID("de"))
  MsgBox Hex$(Rfc1766_ToLCID("de-de"))
  MsgBox Hex$(Rfc1766_ToLCID("de-at"))
  MsgBox Hex$(Rfc1766_ToLCID("de-AT"))
  MsgBox Hex$(Rfc1766_ToLCID("sv-fi"))
  MsgBox CStr(Rfc1766_ToLCID("sv-fi"))
  MsgBox CStr(Rfc1766_ToLCID("sv-SE"))
  
End Sub

