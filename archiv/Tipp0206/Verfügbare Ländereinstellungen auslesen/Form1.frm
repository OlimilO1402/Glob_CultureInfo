VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "www.activevb.de"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   ScaleHeight     =   2850
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3015
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Installierte Einstellungen"
      Height          =   195
      Index           =   0
      Left            =   3240
      TabIndex        =   2
      Top             =   1980
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Unterstütze Einstellungen"
      Height          =   195
      Index           =   1
      Left            =   3240
      TabIndex        =   1
      Top             =   2220
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Alternatives Sortiment"
      Height          =   195
      Index           =   2
      Left            =   3240
      TabIndex        =   0
      Top             =   2460
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Option1(0).Value = True
End Sub

Private Sub Option1_Click(Index As Integer)
  Call EnumLocales(Index, List1)
End Sub
