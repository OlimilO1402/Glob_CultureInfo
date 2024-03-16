Attribute VB_Name = "MLang"
Option Explicit

Private Const LOCALE_NAME_MAX_LENGTH        As Long = 85
Private Const LOCALE_ALLOW_NEUTRAL_NAMES    As Long = &H8000000

'https://learn.microsoft.com/de-de/openspecs/windows_protocols/ms-lcid/70feba9f-294e-491e-b6eb-56532684c37f

'MLang.dll
'https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa741205(v=vs.85)
'Private Declare Function LcidToRfc1766W Lib "MLang" (ByVal LocaleID As Long, pSzRfc1766 As Any, ByVal nChar As Long) As Long

'https://learn.microsoft.com/de-de/windows/win32/api/winnls/nf-winnls-lcidtolocalename
Private Declare Function LCIDToLocaleName Lib "kernel32" (ByVal LocaleID As Long, lpName As Any, ByVal cchName As Long, ByVal dwFlags As Long) As Long

'https://learn.microsoft.com/de-de/windows/win32/api/winnls/nf-winnls-localenametolcid
Private Declare Function LocaleNameToLCID Lib "kernel32" (ByVal lpName As Any, ByVal dwFlags As Long) As Long
'LCID LocaleNameToLCID(
'  [in] LPCWSTR lpName,
'  [in] DWORD   dwFlags
');

'https://learn.microsoft.com/de-de/windows/win32/intl/national-language-support

'https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa741208(v=vs.85)
'Private Declare Function Rfc1766ToLcidW Lib "MLang" (ByRef pLocaleID_out As Any, pSzRfc1766 As Any) As Long

'https://learn.microsoft.com/en-us/windows/win32/api/winver/nf-winver-verlanguagenamew
Private Declare Function VerLanguageNameW Lib "kernel32" (ByVal LocaleID As Long, ByRef pSzLang_out As Any, ByVal nSize As Long) As Long
'DWORD VerLanguageNameW(
'  [in]  DWORD  wLang,
'  [out] LPWSTR szLang,
'  [in]  DWORD  cchLang
');
'Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Any) As Long


Public Function LCID_ToLocaleName(ByVal aLCID As Long) As String
    Dim l As Long:   l = LOCALE_NAME_MAX_LENGTH
    Dim s As String: s = Space$(l)
    l = LCIDToLocaleName(aLCID, ByVal StrPtr(s), l, LOCALE_ALLOW_NEUTRAL_NAMES)
    If l = 0 Then Exit Function
    LCID_ToLocaleName = Left(s, l - 1)
End Function

Public Function LCID_ToLanguageName(ByVal aLCID As Long) As String
    'retutrns the fullqualified name of the language in the users language
    Dim l As Long:   l = 256
    Dim s As String: s = String$(l, vbNullChar)
    l = VerLanguageNameW(aLCID, ByVal StrPtr(s), l)
    If l = 0 Then Exit Function
    LCID_ToLanguageName = Trim0(Left$(s, l))
End Function

Public Function LocaleName_ToLCID(ByVal sLocaleName As String) As Long
    LocaleName_ToLCID = LocaleNameToLCID(StrPtr(sLocaleName), LOCALE_ALLOW_NEUTRAL_NAMES)
End Function

'                              1                               2                             3    |
'0  1  2  3  4  5  6  7  8  9  0 | 1  2  3  4  5 | 6  7  8  9  0  1  2  3  4  5  6  7  8  9  0  1 |
'           Reserved             |    Sort ID    |                   Language ID                  |
Public Function MAKELCID(ByVal lgid As Long, ByVal srtid As Long) As Long
    MAKELCID = lgid Or (srtid * 65536) '2^16 = 65536
End Function

Public Function MAKELANGID(ByVal p As Long, ByVal s As Long) As Long
    'p: value in the range between 0x03FF - 0x0200
    's: value in the range between   0x20 -   0x3F
    MAKELANGID = s Or p
End Function

Private Function Trim0(ByVal s As String) As String
    Trim0 = VBA.Strings.Trim$(Left$(s, lstrlenW(ByVal StrPtr(s))))
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

'
'static void get_name_record_locale(enum OPENTYPE_PLATFORM_ID platform, USHORT lang_id, WCHAR *locale, USHORT locale_len)
'{
'    static const WCHAR enusW[] = {'e','n','-','U','S',0};
'
'    switch (platform) {
'    Case OPENTYPE_PLATFORM_MAC:
'    {
'        const char *locale_name = NULL;
'
'        if (lang_id > TT_NAME_MAC_LANGID_AZER_ROMAN)
'            ERR("invalid mac lang id %d\n", lang_id);
'        else if (!name_mac_langid_to_locale[lang_id][0])
'            FIXME("failed to map mac lang id %d to locale name\n", lang_id);
'        Else
'            locale_name = name_mac_langid_to_locale[lang_id];
'
'        if (locale_name)
'            MultiByteToWideChar(CP_ACP, 0, name_mac_langid_to_locale[lang_id], -1, locale, locale_len);
'        Else
'            strcpyW(locale, enusW);
'        break;
'    }
'    Case OPENTYPE_PLATFORM_WIN:
'        if (!LCIDToLocaleName(MAKELCID(lang_id, SORT_DEFAULT), locale, locale_len, 0)) {
'            FIXME("failed to get locale name for lcid=0xx\n", MAKELCID(lang_id, SORT_DEFAULT));
'            strcpyW(locale, enusW);
'        }
'        break;
'default:
'        FIXME("unknown platform %d\n", platform);
'    }
'}
