Attribute VB_Name = "Mlang"
Option Explicit

'MLang.dll
Public Declare Function LcidToRfc1766W Lib "mlang" (ByVal Locale As Long, pszRfc1766 As Any, ByVal nChar As Long) As Long
Public Declare Function Rfc1766ToLcidW Lib "mlang" (ByRef pLocale As Long, pszRfc1766 As Any) As Long


Public Function LCID_ToRfc1766(aLCID As Long) As String
    
    Dim hr As Long
    Dim i As Long
    
    LCID_ToRfc1766 = String$(6, vbNullChar)
    hr = LcidToRfc1766W(aLCID, ByVal StrPtr(LCID_ToRfc1766), 6)
    
    If hr = 0 Then
        For i = 0 To 1
            If Len(LCID_ToRfc1766) > 3 + i Then
                Mid$(LCID_ToRfc1766, 4 + i, 1) = UCase$(Mid$(LCID_ToRfc1766, 4 + i, 1))
            End If
        Next
    Else
      'Msgbox "Fehler"
    End If
End Function

Public Function Rfc1766_ToLCID(aStrRfc1766 As String) As Long
    Dim hr As Long: hr = Rfc1766ToLcidW(Rfc1766_ToLCID, ByVal StrPtr(aStrRfc1766))
End Function

