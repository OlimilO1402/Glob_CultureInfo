Attribute VB_Name = "ModMlang"
Option Explicit

'MLang.dll
Public Declare Function LcidToRfc1766A Lib "mlang" ( _
    ByVal Locale As Long, _
    ByVal pszRfc1766 As String, _
    ByVal nChar As Long) As Long
'Public Declare Function LcidToRfc1766W Lib "mlang" ( _
'    ByVal Locale As Long, _
'    ByVal pszRfc1766 As Long, _
'    ByVal nChar As Long) As Long

Public Declare Function Rfc1766ToLcidA Lib "mlang" ( _
    ByRef pLocale As Long, _
    ByVal pszRfc1766 As String) As Long
'Public Declare Function Rfc1766ToLcidW Lib "mlang" ( _
'    ByRef pLocale As Long, _
'    ByVal pszRfc1766 As Long) As Long

Public Function ConvertLCIDToRfc1766(aLCID As Long) As String
  
  Dim hr As Long
  Dim i As Long
  
  ConvertLCIDToRfc1766 = String$(6, vbNullChar)
  hr = LcidToRfc1766A(aLCID, ConvertLCIDToRfc1766, 6)
  
  If hr = 0 Then
    For i = 0 To 1
      If Len(ConvertLCIDToRfc1766) > 3 + i Then
        Mid$(ConvertLCIDToRfc1766, 4 + i, 1) = UCase$(Mid$(ConvertLCIDToRfc1766, 4 + i, 1))
      End If
    Next
  Else
    'Msgbox "Fehler"
  End If
End Function
Public Function ConvertRfc1766ToLCID(aStrrfc1766 As String) As Long
  Dim hr As Long
  
  hr = Rfc1766ToLcidA(ConvertRfc1766ToLCID, aStrrfc1766)
End Function

