VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConstsWinNlsh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_PFN       As PathFileName
Private m_Constants As List 'Of CConstant

Private Sub Class_Initialize()
    Set m_Constants = MNew.List(vbObject, , True)
    Set m_PFN = New PathFileName
End Sub

Friend Sub New_(aPFN As PathFileName)
    Set m_PFN = aPFN
    ParseWinNlsh
End Sub

Friend Sub NewC(other As ConstsWinNlsh)
    Set m_PFN = other.pfn.Clone
    Set m_Constants = other.Constants.Clone
End Sub

Friend Function Clone() As ConstsWinNlsh
    Set Clone = New ConstsWinNlsh: Clone.NewC Me
End Function

Private Sub ParseWinNlsh()
    Dim lines() As String
    If Not m_PFN.TryReadAllLines(lines) Then MsgBox "Could not read lines": Exit Sub
    Dim line As String, c As CConstant
    Dim i As Long, u As Long: u = UBound(lines)
    If u = 0 Then Exit Sub
    For i = 0 To u
        line = Trim(lines(i))
        If LCase(Left(line, 8)) = "#define " Then
            line = Mid(line, 9)
            Set c = ParseCConstant(line)
            If Not c Is Nothing Then m_Constants.Add c, c.Key
        End If
    Next
End Sub

Private Function ParseCConstant(ByVal aLine As String) As CConstant
    Dim nm As String, v As String, cmnt As String
    'first read comment from right
    Dim pos As Long: pos = InStr(aLine, "//")
    If pos > 0 Then
        cmnt = Mid(aLine, pos)
        aLine = Left(aLine, pos)
    End If
    aLine = MString.RecursiveReplace(aLine, "  ", " ")
    Dim elems() As String: elems = Split(aLine, " ")
    nm = elems(0)
    If UBound(elems) > 0 Then v = elems(1)
    Set ParseCConstant = MNew.CConstant(nm, v, cmnt)
End Function

Public Function GetMaxNameLen() As Long
    Dim i As Long, c As CConstant
    For i = 0 To m_Constants.Count - 1
        Set c = m_Constants.Item(i)
        GetMaxNameLen = Max(GetMaxNameLen, Len(c.Key))
    Next
End Function

Public Property Get pfn() As PathFileName
    Set pfn = m_PFN
End Property

Public Property Get Constants() As List 'Of CConstant
    Set Constants = m_Constants
End Property

Public Sub ConstsSortByValue()
    m_Constants.Sort Me, "CompareByValue"
End Sub

Public Function CompareByValue(obj1 As CConstant, obj2 As CConstant) As Long
    If obj1.Value > obj2.Value Then
        CompareByValue = 1
    ElseIf obj1.Value < obj2.Value Then
        CompareByValue = -1
    End If
End Function

Public Sub ConstsSortByName()
    m_Constants.Sort Me, "CompareByName"
End Sub

Public Function CompareByName(obj1 As CConstant, obj2 As CConstant) As Long

End Function

Public Sub ToLB(aLBCB) ' As ListBox)
    Dim mx As Long: mx = GetMaxNameLen
    aLBCB.Clear
    Dim i As Long, c As CConstant
    For i = 0 To m_Constants.Count - 1
        Set c = m_Constants.Item(i)
        aLBCB.AddItem c.ToStr(mx)
    Next
End Sub

Public Function Filter(ByVal NameStartsWith As String) As ConstsWinNlsh
    Set Filter = New ConstsWinNlsh
    Filter.pfn.Value = m_PFN.Value
    Dim i As Long, c As CConstant
    For i = 0 To m_Constants.Count - 1
        Set c = m_Constants.Item(i)
        If MString.StartsWith(c.Key, NameStartsWith) Then
            Filter.Constants.Add c, c.Key
        End If
    Next
End Function

Public Function FilterSA(aNames() As String) As ConstsWinNlsh
    Set FilterSA = New ConstsWinNlsh
    FilterSA.pfn.Value = m_PFN.Value
    Dim i As Long, nm As String, c As CConstant
    For i = 0 To UBound(aNames) 'm_Constants.Count - 1
        nm = aNames(i)
        If m_Constants.ContainsKey(nm) Then
            Set c = m_Constants.ItemByKey(nm)
            FilterSA.Constants.Add c, c.Key
        End If
    Next
End Function

