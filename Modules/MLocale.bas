Attribute VB_Name = "MLocale"
Option Explicit

Public Enum EEnumLocalesMode
    Mode_EnumSystemLocale
    Mode_EnumSystemLocaleEx
End Enum

'//  Locale Enumeration Flags.
Public Const LCID_INSTALLED       As Long = &H1 ' Aufzählen nur installierte Gebietsschemabezeichner. Dieser Wert kann nicht mit LCID_SUPPORTED verwendet werden.
Public Const LCID_SUPPORTED       As Long = &H2 ' Aufzählen aller unterstützten Gebietsschemabezeichner. Dieser Wert kann nicht mit LCID_INSTALLED verwendet werden.
Public Const LCID_ALTERNATE_SORTS As Long = &H4 ' Aufzählen Sie nur die alternativen Sortierschemabezeichner. Wenn dieser Wert entweder mit LCID_INSTALLED oder LCID_SUPPORTED verwendet wird, werden die installierten oder unterstützten Gebietsschemas sowie die alternativen Sortierschemabezeichner abgerufen.

'BOOL EnumSystemLocalesW(
'  [in] LOCALE_ENUMPROCW lpLocaleEnumProc,
'  [in] DWORD            dwFlags
');
Private Declare Function EnumSystemLocalesA Lib "kernel32" (ByVal lpLocaleEnumProc As LongPtr, ByVal dwFlags As Long) As Long
Private Declare Function EnumSystemLocalesW Lib "kernel32" (ByVal lpLocaleEnumProc As LongPtr, ByVal dwFlags As Long) As Long

'Flags identifying the locales to enumerate. The flags can be used singly or combined using a binary OR. If the application specifies 0 for this parameter, the function behaves as for LOCALE_ALL.
'//  Named based enumeration flags.
Public Const LOCALE_ALL             As Long = &H0   '// enumerate all named based locales
Public Const LOCALE_WINDOWS         As Long = &H1   '// shipped locales and/or replacements for them
Public Const LOCALE_SUPPLEMENTAL    As Long = &H2   '// supplemental locales only
Public Const LOCALE_ALTERNATE_SORTS As Long = &H4   '// alternate sort locales,  see Remarks
Public Const LOCALE_REPLACEMENT     As Long = &H8   '// locales that replace shipped locales (callback flag only)
Public Const LOCALE_NEUTRALDATA     As Long = &H10  '// Locales that are "neutral" (language only, region data is default)
Public Const LOCALE_SPECIFICDATA    As Long = &H20  '// Locales that contain language and region data

'https://learn.microsoft.com/en-us/windows/win32/api/winnls/nf-winnls-enumsystemlocalesex
'BOOL EnumSystemLocalesEx(
'  [in]           LOCALE_ENUMPROCEX lpLocaleEnumProcEx,
'  [in]           DWORD             dwFlags,
'  [in]           LPARAM            lParam,
'  [in, optional] LPVOID            lpReserved
');
Private Declare Function EnumSystemLocalesEx Lib "kernel32" (ByVal lpLocaleEnumProcEx As LongPtr, ByVal dwFlags As Long, Optional ByVal lParam As Long = 0, Optional ByVal lpReserved As LongPtr = 0) As Long
'https://learn.microsoft.com/en-us/windows/win32/intl/nls--name-based-apis-sample

Private m_Locales As Collection

Public Function GetMaxNameLen() As Long
    Dim v, ci As CultureInfo
    If m_Locales Is Nothing Then Exit Function
    For Each v In m_Locales
        Set ci = v
        GetMaxNameLen = Max(GetMaxNameLen, Len(ci.Name))
    Next
End Function

Public Function GetCultureInfos(mode As EEnumLocalesMode, ByVal Flags As Long) As Collection
    Set m_Locales = New Collection
    Dim hr As Long
    Select Case mode
    Case Mode_EnumSystemLocale:   hr = EnumSystemLocalesW(FncPtr(AddressOf CALLBACK_EnumLocalesProcW), Flags)
    Case Mode_EnumSystemLocaleEx: hr = EnumSystemLocalesEx(FncPtr(AddressOf CALLBACK_EnumLocalesProcEx), Flags)
    End Select
    Set GetCultureInfos = m_Locales
End Function

'https://learn.microsoft.com/en-us/previous-versions/windows/desktop/legacy/dd317822(v=vs.85)
'BOOL CALLBACK EnumLocalesProc(
'  _In_ LPTSTR lpLocaleString
');
Private Function CALLBACK_EnumLocalesProcW(ByVal lpLocaleString As LongPtr) As Long
    Dim s As String: s = MString.PtrToString(lpLocaleString)
    Dim ci As CultureInfo: Set ci = MNew.CultureInfo(CLng("&H" & s))
    m_Locales.Add ci, ci.Name
    CALLBACK_EnumLocalesProcW = 1
End Function

'https://learn.microsoft.com/en-us/windows/win32/api/winnls/nc-winnls-locale_enumprocex
'LOCALE_ENUMPROCEX LocaleEnumprocex;
'BOOL LocaleEnumprocex(
'  LPWSTR unnamedParam1,
'  DWORD unnamedParam2,
'  lParam unnamedParam3
')
'{...}
Private Function CALLBACK_EnumLocalesProcEx(ByVal unnamedParam1 As LongPtr, ByVal unnamedParam2 As Long, ByVal unnamedParam3 As Long) As Long
    Dim s As String: s = MString.PtrToString(unnamedParam1)
    Dim ci As CultureInfo: Set ci = MNew.CultureInfoN(s)
    m_Locales.Add ci, ci.Name
    CALLBACK_EnumLocalesProcEx = 1
End Function

