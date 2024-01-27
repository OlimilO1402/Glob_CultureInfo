Attribute VB_Name = "MLang"
Option Explicit

'https://learn.microsoft.com/de-de/openspecs/windows_protocols/ms-lcid/70feba9f-294e-491e-b6eb-56532684c37f

'MLang.dll
'https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa741205(v=vs.85)
Private Declare Function LcidToRfc1766W Lib "MLang" (ByVal LocaleID As Long, pSzRfc1766 As Any, ByVal nChar As Long) As Long

'https://learn.microsoft.com/de-de/windows/win32/api/winnls/nf-winnls-lcidtolocalename
Private Declare Function LCIDToLocaleName Lib "kernel32" (ByVal LocaleID As Long, lpName As Any, ByVal cchName As Long, ByVal dwFlags As Long) As Long

'https://learn.microsoft.com/de-de/windows/win32/api/winnls/nf-winnls-localenametolcid
'https://learn.microsoft.com/de-de/windows/win32/intl/national-language-support

'https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa741208(v=vs.85)
Private Declare Function Rfc1766ToLcidW Lib "MLang" (ByRef pLocaleID_out As Any, pSzRfc1766 As Any) As Long

'https://learn.microsoft.com/en-us/windows/win32/api/winver/nf-winver-verlanguagenamew
Private Declare Function VerLanguageNameW Lib "kernel32" (ByVal LocaleID As Long, ByRef pSzLang_out As Any, ByVal nSize As Long) As Long
'DWORD VerLanguageNameW(
'  [in]  DWORD  wLang,
'  [out] LPWSTR szLang,
'  [in]  DWORD  cchLang
');
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Any) As Long


Public Function LCID_ToRfc1766(ByVal aLCID As Long) As String
    Dim l As Long: l = 32
    Dim s As String: s = Space$(l) 'String$(6, vbNullChar)
    Dim hr As Long
    Dim bMLang As Boolean: bMLang = True
    If bMLang Then
        hr = LcidToRfc1766W(aLCID, ByVal StrPtr(s), 6)
        If hr <> 0 Then
        ' Msgbox "Fehler"
            Exit Function
        End If
        Dim i As Long
        For i = 0 To 1
            If Len(s) > 3 + i Then
                Mid$(s, 4 + i, 1) = UCase$(Mid$(s, 4 + i, 1))
            End If
        Next
        s = Trim0(s)
    Else
        hr = LCIDToLocaleName(aLCID, ByVal StrPtr(s), l, 0)
        If hr = 0 Then
            ' Msgbox "Fehler"
            Exit Function
        End If
        s = Left(s, hr - 1)
    End If
    LCID_ToRfc1766 = s
End Function

Public Function LCID_ToLanguageName(ByVal aLCID As Long) As String
    'retutrns the fullqualified name of the language in the users language
    Dim l As Long:   l = 256
    Dim s As String: s = String$(l, vbNullChar)
    l = VerLanguageNameW(aLCID, ByVal StrPtr(s), l)
    If l = 0 Then Exit Function
    LCID_ToLanguageName = Trim0(Left$(s, l))
End Function

Private Function Trim0(ByVal s As String) As String
    Trim0 = VBA.Strings.Trim$(Left$(s, lstrlenW(ByVal StrPtr(s))))
End Function

Public Function Rfc1766String_ToLCID(ByVal sRfc1766 As String) As Long
    Dim hr As Long: hr = Rfc1766ToLcidW(Rfc1766String_ToLCID, ByVal StrPtr(sRfc1766))
End Function

Public Function ListLCIDStrings() As Collection
    
    Dim HashCol As New Collection
    Dim lcids   As New Collection
    
    Dim c As Long
    Dim i As Long
    For i = 1 To &HFF
        Dim j As Long
        For j = 0 To &HFF
            'Dim sBuf As String * 256
            Dim lcid  As Long: lcid = j * &H400 + i
            Dim sLcid As String: sLcid = LCID_ToRfc1766(lcid)
            lcid = Rfc1766String_ToLCID(sLcid)
            sLcid = LCID_ToRfc1766(lcid)
            If Len(sLcid) < 5 Then sLcid = sLcid & Space$(5 - Len(sLcid))
            'Dim sLen As Long: sLen = VerLanguageNameW(lcid, StrPtr(sBuf), Len(sBuf))
            'If sLen > 0 Then
                Dim sLCItem As String: sLCItem = LCID_ToLanguageName(lcid) 'Left$(sBuf, sLen)
                'lcid = LCID_ToRfc1766(sLCItem)
                If Not Col_Contains(HashCol, sLCItem) Then
                    HashCol.Add sLCItem, sLCItem
                    Dim sH As String: sH = Hex$(lcid)
                    sH = String$(4 - Len(sH), "0") & sH
                    Dim sItem As String: sItem = "&H" & sH & "    " & sLcid & "   " & sLCItem
                    c = c + 1
                    sH = CStr(c)
                    sH = String$(3 - Len(sH), "0") & sH & "    "
                    sItem = sH & sItem
                    lcids.Add sItem
                End If
            'End If
        Next
    Next
    Set ListLCIDStrings = lcids
End Function

Private Function Col_Contains(col As Collection, Key As String) As Boolean
    'for this Function all credits go to the incredible www.vb-tec.de alias Jost Schwider
    'you can find the original version of this function here: https://vb-tec.de/collctns.htm
    On Error Resume Next
'  '"Extras->Optionen->Allgemein->Unterbrechen bei Fehlern->Bei nicht verarbeiteten Fehlern"
    If IsEmpty(col(Key)) Then: 'DoNothing
    Col_Contains = (Err.Number = 0)
    On Error GoTo 0
End Function

