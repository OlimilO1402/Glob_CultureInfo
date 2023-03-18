VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "www.activevb.de"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   ScaleHeight     =   7155
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1455
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3375
      Begin VB.OptionButton Option1 
         Caption         =   "Alternatives Sortiment"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Unterstütze Einstellungen"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Installierte Einstellungen"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.ListBox List1 
      Height          =   5715
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   5415
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   360
      Width           =   1455
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
  Label1.Caption = List1.ListCount
End Sub
