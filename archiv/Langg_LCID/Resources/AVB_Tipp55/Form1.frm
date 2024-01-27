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
  
  MsgBox ConvertLCIDToRfc1766(&H407) 'de
  MsgBox ConvertLCIDToRfc1766(1031)  'de
  MsgBox ConvertLCIDToRfc1766(&HC07) 'de-AT
  MsgBox ConvertLCIDToRfc1766(&H7)   'de-DE
  MsgBox ConvertLCIDToRfc1766(2077)  'sv-FI
  MsgBox ConvertLCIDToRfc1766(&H1D)  'sv-SE
  MsgBox ConvertLCIDToRfc1766(&H81D) 'sv-FI

  MsgBox Hex$(ConvertRfc1766ToLCID("de"))
  MsgBox Hex$(ConvertRfc1766ToLCID("de-de"))
  MsgBox Hex$(ConvertRfc1766ToLCID("de-at"))
  MsgBox Hex$(ConvertRfc1766ToLCID("de-AT"))
  MsgBox Hex$(ConvertRfc1766ToLCID("sv-fi"))
  MsgBox CStr(ConvertRfc1766ToLCID("sv-fi"))
  MsgBox CStr(ConvertRfc1766ToLCID("sv-SE"))
  
End Sub

