VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CConstant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Name  As String
Private m_Value As String
Private m_Comm  As String

Friend Sub New_(ByVal aName As String, ByVal aValue As String, ByVal aComment As String)
    'if left(aValue,2)="0x" then aValue =
    m_Name = aName: m_Value = aValue: m_Comm = aComment
End Sub

Public Property Get Key() As String
    Key = m_Name
End Property

Public Property Get Value() As String
    Value = m_Value
End Property

Public Property Get IsCHex() As Boolean
    IsCHex = MString.StartsWith(CStr(m_Value), "0x")
End Property

Public Property Get TryGetValueAsLong(aValue As Long) As Boolean
Try: On Error GoTo Catch
    If IsCHex Then
        aValue = CLng(CHexToVBHex(m_Value))
        TryGetValueAsLong = True
    Else
        If IsNumeric(m_Value) Then
            aValue = CLng(m_Value)
            TryGetValueAsLong = True
        End If
    End If
    Exit Property
Catch:
End Property

Public Function ToVBCode(Optional ByVal NameLen As Long) As String
    ToVBCode = "Public Const " & MString.PadRight(m_Name, NameLen) & " As Long = &H" & Hex(Me.Value) 'AsLong)
End Function

Friend Function ToStr(Optional ByVal NameLen As Long, Optional ByVal sep As String = " = ") As String
    Dim v As Long, sv As String: sv = m_Value
    If TryGetValueAsLong(v) Then sv = "&H" & Hex(v)
    ToStr = PadRight(m_Name, NameLen) & sep & sv
End Function

