VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CultureInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'https://learn.microsoft.com/de-de/windows/win32/intl/locale-information-constants
'Class CultureInfo
'Stellt Informationen über eine bestimmte Kultur bereit
'(die bei der nicht verwalteten Codeentwicklung als "locale" bezeichnet wird).
'Zu diesen Informationen gehören
' * der Name der Kultur,
' * das Schriftsystem,
' * der verwendete Kalender sowie
' * die Formatierung für Datumsangaben und
' * sortierte Zeichenfolgen.

Private Const LOCALE_NEUTRAL                As Long = &H0
Private Const LOCALE_ILANGUAGE              As Long = &H1

Private Const LOCALE_SLANGUAGE              As Long = &H2
Private Const LOCALE_SLOCALIZEDDISPLAYNAME  As Long = &H2

Private Const LOCALE_SABBREVLANGNAME        As Long = &H3
Private Const LOCALE_SNATIVELANGUAGENAME    As Long = &H4

Private Const LOCALE_ICOUNTRY               As Long = &H5
Private Const LOCALE_IDIALINGCODE           As Long = &H5

Private Const LOCALE_SCOUNTRY               As Long = &H6
Private Const LOCALE_SLOCALIZEDCOUNTRYNAME  As Long = &H6

Private Const LOCALE_SABBREVCTRYNAME        As Long = &H7
Private Const LOCALE_SABBREVCOUNTRYNAME     As Long = &H7

Private Const LOCALE_SNATIVECTRYNAME        As Long = &H8
Private Const LOCALE_SNATIVECOUNTRYNAME     As Long = &H8

Private Const LOCALE_IDEFAULTLANGUAGE       As Long = &H9
Private Const LOCALE_IDEFAULTCOUNTRY        As Long = &HA
Private Const LOCALE_IDEFAULTCODEPAGE       As Long = &HB
Private Const LOCALE_SLIST                  As Long = &HC
Private Const LOCALE_IMEASURE               As Long = &HD

Private Const LOCALE_SDECIMAL               As Long = &HE
Private Const LOCALE_STHOUSAND              As Long = &HF
Private Const LOCALE_SGROUPING              As Long = &H10
Private Const LOCALE_IDIGITS                As Long = &H11
Private Const LOCALE_ILZERO                 As Long = &H12
Private Const LOCALE_SNATIVEDIGITS          As Long = &H13

Private Const LOCALE_SCURRENCY              As Long = &H14
Private Const LOCALE_SINTLSYMBOL            As Long = &H15
Private Const LOCALE_SMONDECIMALSEP         As Long = &H16
Private Const LOCALE_SMONTHOUSANDSEP        As Long = &H17
Private Const LOCALE_SMONGROUPING           As Long = &H18
Private Const LOCALE_ICURRDIGITS            As Long = &H19
Private Const LOCALE_IINTLCURRDIGITS        As Long = &H1A
Private Const LOCALE_ICURRENCY              As Long = &H1B
Private Const LOCALE_INEGCURR               As Long = &H1C

Private Const LOCALE_SDATE                  As Long = &H1D
Private Const LOCALE_STIME                  As Long = &H1E
Private Const LOCALE_SSHORTDATE             As Long = &H1F
Private Const LOCALE_SLONGDATE              As Long = &H20
Private Const LOCALE_IDATE                  As Long = &H21
Private Const LOCALE_ILDATE                 As Long = &H22
Private Const LOCALE_ITIME                  As Long = &H23
Private Const LOCALE_ICENTURY               As Long = &H24
Private Const LOCALE_ITLZERO                As Long = &H25
Private Const LOCALE_IDAYLZERO              As Long = &H26
Private Const LOCALE_IMONLZERO              As Long = &H27

Private Const LOCALE_S1159                  As Long = &H28
Private Const LOCALE_SAM                    As Long = &H28

Private Const LOCALE_S2359                  As Long = &H29
Private Const LOCALE_SPM                    As Long = &H29

Private Const LOCALE_SDAYNAME1              As Long = &H2A
Private Const LOCALE_SDAYNAME2              As Long = &H2B
Private Const LOCALE_SDAYNAME3              As Long = &H2C
Private Const LOCALE_SDAYNAME4              As Long = &H2D
Private Const LOCALE_SDAYNAME5              As Long = &H2E
Private Const LOCALE_SDAYNAME6              As Long = &H2F
Private Const LOCALE_SDAYNAME7              As Long = &H30
Private Const LOCALE_SABBREVDAYNAME1        As Long = &H31
Private Const LOCALE_SABBREVDAYNAME2        As Long = &H32
Private Const LOCALE_SABBREVDAYNAME3        As Long = &H33
Private Const LOCALE_SABBREVDAYNAME4        As Long = &H34
Private Const LOCALE_SABBREVDAYNAME5        As Long = &H35
Private Const LOCALE_SABBREVDAYNAME6        As Long = &H36
Private Const LOCALE_SABBREVDAYNAME7        As Long = &H37
Private Const LOCALE_SMONTHNAME1            As Long = &H38
Private Const LOCALE_SMONTHNAME2            As Long = &H39
Private Const LOCALE_SMONTHNAME3            As Long = &H3A
Private Const LOCALE_SMONTHNAME4            As Long = &H3B
Private Const LOCALE_SMONTHNAME5            As Long = &H3C
Private Const LOCALE_SMONTHNAME6            As Long = &H3D
Private Const LOCALE_SMONTHNAME7            As Long = &H3E
Private Const LOCALE_SMONTHNAME8            As Long = &H3F
Private Const LOCALE_SMONTHNAME9            As Long = &H40
Private Const LOCALE_SMONTHNAME10           As Long = &H41
Private Const LOCALE_SMONTHNAME11           As Long = &H42
Private Const LOCALE_SMONTHNAME12           As Long = &H43
Private Const LOCALE_SABBREVMONTHNAME1      As Long = &H44
Private Const LOCALE_SABBREVMONTHNAME2      As Long = &H45
Private Const LOCALE_SABBREVMONTHNAME3      As Long = &H46
Private Const LOCALE_SABBREVMONTHNAME4      As Long = &H47
Private Const LOCALE_SABBREVMONTHNAME5      As Long = &H48
Private Const LOCALE_SABBREVMONTHNAME6      As Long = &H49
Private Const LOCALE_SABBREVMONTHNAME7      As Long = &H4A
Private Const LOCALE_SABBREVMONTHNAME8      As Long = &H4B
Private Const LOCALE_SABBREVMONTHNAME9      As Long = &H4C
Private Const LOCALE_SABBREVMONTHNAME10     As Long = &H4D
Private Const LOCALE_SABBREVMONTHNAME11     As Long = &H4E
Private Const LOCALE_SABBREVMONTHNAME12     As Long = &H4F

Private Const LOCALE_SPOSITIVESIGN          As Long = &H50
Private Const LOCALE_SNEGATIVESIGN          As Long = &H51
Private Const LOCALE_IPOSSIGNPOSN           As Long = &H52
Private Const LOCALE_INEGSIGNPOSN           As Long = &H53
Private Const LOCALE_IPOSSYMPRECEDES        As Long = &H54
Private Const LOCALE_IPOSSEPBYSPACE         As Long = &H55
Private Const LOCALE_INEGSYMPRECEDES        As Long = &H56
Private Const LOCALE_INEGSEPBYSPACE         As Long = &H57

Private Const LOCALE_FONTSIGNATURE          As Long = &H58

Private Const LOCALE_SISO639LANGNAME        As Long = &H59
Private Const LOCALE_SISO3166CTRYNAME       As Long = &H5A

Private Const LOCALE_IGEOID                 As Long = &H5B ' new
Private Const LOCALE_SNAME                  As Long = &H5C ' new
Private Const LOCALE_SDURATION              As Long = &H5D ' new
Private Const LOCALE_SKEYBOARDSTOINSTALL    As Long = &H5E ' new
'Private Const ?                             As Long = &H5F
Private Const LOCALE_SSHORTESTDAYNAME1      As Long = &H60 ' new
Private Const LOCALE_SSHORTESTDAYNAME2      As Long = &H61 ' new
Private Const LOCALE_SSHORTESTDAYNAME3      As Long = &H62 ' new
Private Const LOCALE_SSHORTESTDAYNAME4      As Long = &H63 ' new
Private Const LOCALE_SSHORTESTDAYNAME5      As Long = &H64 ' new
Private Const LOCALE_SSHORTESTDAYNAME6      As Long = &H65 ' new
Private Const LOCALE_SSHORTESTDAYNAME7      As Long = &H66 ' new
Private Const LOCALE_SISO639LANGNAME2       As Long = &H67 ' new
Private Const LOCALE_SISO3166CTRYNAME2      As Long = &H68 ' new
Private Const LOCALE_SNAN                   As Long = &H69 ' new
Private Const LOCALE_SPOSINFINITY           As Long = &H6A ' new
Private Const LOCALE_SNEGINFINITY           As Long = &H6B ' new
Private Const LOCALE_SSCRIPTS               As Long = &H6C ' new
Private Const LOCALE_SORTNAME               As Long = &H6D ' new

Private Const LOCALE_SPARENT                As Long = &H6D

Private Const LOCALE_SCONSOLEFALLBACKNAME   As Long = &H6E ' new
Private Const LOCALE_SLANGDISPLAYNAME       As Long = &H6F ' new
Private Const LOCALE_SLOCALIZEDLANGUAGENAME As Long = &H6F ' new
Private Const LOCALE_IREADINGLAYOUT         As Long = &H70 ' new
Private Const LOCALE_INEUTRAL               As Long = &H71 ' new
Private Const LOCALE_SENGLISHDISPLAYNAME    As Long = &H72 ' new
Private Const LOCALE_SNATIVEDISPLAYNAME     As Long = &H73 ' new
Private Const LOCALE_INEGATIVEPERCENT       As Long = &H74 ' new
Private Const LOCALE_IPOSITIVEPERCENT       As Long = &H75 ' new
Private Const LOCALE_SPERCENT               As Long = &H76 ' new
Private Const LOCALE_SPERMILLE              As Long = &H77 ' new
Private Const LOCALE_SMONTHDAY              As Long = &H78 ' new
Private Const LOCALE_SSHORTTIME             As Long = &H79 ' new
Private Const LOCALE_SOPENTYPELANGUAGETAG   As Long = &H7A ' new
Private Const LOCALE_SSORTLOCALE            As Long = &H7B ' new
Private Const LOCALE_SRELATIVELONGDATE      As Long = &H7C '124
'Private Const ???                           As Long = &H7D '125
Private Const LOCALE_SSHORTESTAM            As Long = &H7E ' new
Private Const LOCALE_SSHORTESTPM            As Long = &H7F ' new

'Private Const LOCALE_NEUTRAL                As Long = 0
'Private Const LOCALE_INVARIANT              As Long = &H7F     '    127

Private Const LOCALE_USER_DEFAULT           As Long = &H400    '   1024
Private Const LOCALE_TRADITIONAL_SPANISH    As Long = &H40A    '   1034
Private Const LOCALE_SYSTEM_DEFAULT         As Long = &H800    '   2048
Private Const LOCALE_CUSTOM_DEFAULT         As Long = &HC00    '   3072
Private Const LOCALE_CUSTOM_UNSPECIFIED     As Long = &H1000   '   4096

Private Const LOCALE_SENGLANGUAGE           As Long = &H1001 ' new
Private Const LOCALE_SENGLISHLANGUAGENAME   As Long = &H1001 ' new

Private Const LOCALE_SENGCOUNTRY            As Long = &H1002 ' new
Private Const LOCALE_STIMEFORMAT            As Long = &H1003 ' new
Private Const LOCALE_IDEFAULTANSICODEPAGE   As Long = &H1004 ' new
Private Const LOCALE_ITIMEMARKPOSN          As Long = &H1005 ' new
Private Const LOCALE_SYEARMONTH             As Long = &H1006 ' new
Private Const LOCALE_SENGCURRNAME           As Long = &H1007 ' new
Private Const LOCALE_SNATIVECURRNAME        As Long = &H1008 ' new
Private Const LOCALE_ICALENDARTYPE          As Long = &H1009 ' new
Private Const LOCALE_IPAPERSIZE             As Long = &H100A ' new

Private Const LOCALE_INVARIANT              As Long = &H100B ' new
Private Const LOCALE_IOPTIONALCALENDAR      As Long = &H100B ' new
Private Const LOCALE_IFIRSTDAYOFWEEK        As Long = &H100C ' new
Private Const LOCALE_IFIRSTWEEKOFYEAR       As Long = &H100D ' new
Private Const LOCALE_SMONTHNAME13           As Long = &H100E ' new
Private Const LOCALE_SABBREVMONTHNAME13     As Long = &H100F ' new
Private Const LOCALE_INEGNUMBER             As Long = &H1010 ' new
Private Const LOCALE_IDEFAULTMACCODEPAGE    As Long = &H1011 ' new
Private Const LOCALE_IDEFAULTEBCDICCODEPAGE As Long = &H1012 ' new
Private Const LOCALE_SSORTNAME              As Long = &H1013 ' new
Private Const LOCALE_IDIGITSUBSTITUTION     As Long = &H1014 ' new
Private Const LOCALE_CUSTOM_UI_DEFAULT      As Long = &H1400 ' new

Private Const LOCALE_SORTID_MASK            As Long = &HF0000  ' 983040
Private Const LOCALE_ALLOW_NEUTRAL_NAMES    As Long = &H8000000
Private Const LOCALE_RETURN_GENITIVE_NAMES  As Long = &H10000000
Private Const LOCALE_RETURN_NUMBER          As Long = &H20000000
Private Const LOCALE_USE_CP_ACP             As Long = &H40000000
Private Const LOCALE_NOUSEROVERRIDE         As Long = &H80000000

Private Const LOCALE_NAME_MAX_LENGTH        As Long = 85

'https://learn.microsoft.com/de-de/windows/win32/intl/sort-order-identifiers
Private Const SORT_CHINESE_BIG5           As Long = &H0 'Chinese BIG5 order 'Default sort for Simplified Chinese locales 'Use with zh-TW, zh-HK, or zh-MO
Private Const SORT_CHINESE_PRCP           As Long = &H0 'PRC Chinese phonetic order 'Default sort for Traditional Chinese locales. 'Use with zn-CN or zh-SG
Private Const SORT_DEFAULT                As Long = &H0 'Default sort order
Private Const SORT_GEORGIAN_TRADITIONAL   As Long = &H0 'Georgian traditional order 'Default sort for Georgian, use with ka-GE
Private Const SORT_HUNGARIAN_DEFAULT      As Long = &H0 'Hungarian default order 'Default sort for Hungarian, use with hu-HU
Private Const SORT_JAPANESE_XJIS          As Long = &H0 'Japanese XJIS order 'Use with ja-JP
Private Const SORT_KOREAN_KSC             As Long = &H0 'Korean KSC order Default sort for Korean, use with ko-KR
Private Const SORT_CHINESE_UNICODE        As Long = &H1 'Chinese Unicode order<4> 'Do not use, deprecated
Private Const SORT_GEORGIAN_MODERN        As Long = &H1 'Georgian modern order 'Use with ka-GE, resolves to ka-GE_modern
Private Const SORT_GERMAN_PHONE_BOOK      As Long = &H1 'German phone book order 'Use with de-DE, resolves to de-DE_phoneb
Private Const SORT_HUNGARIAN_TECHNICAL    As Long = &H1 'Hungarian technical order 'Use with hu-HU, resolves to hu-HU_technl
Private Const SORT_JAPANESE_UNICODE       As Long = &H1 'Japanese Unicode order<5> 'Do not use, deprecated
Private Const SORT_KOREAN_UNICODE         As Long = &H1 'Korean Unicode order<6> 'Do not use, deprecated
Private Const SORT_CHINESE_PRC            As Long = &H2 'PRC Chinese stroke count order 'Use with zh-CN, resolves to zh-CN_stroke 'Use with zh-SG, resolves to zh-SG_stroke
Private Const SORT_CHINESE_BOPOMOFO       As Long = &H3 'Traditional Chinese Bopomofo order 'Use with zh-TW, resolves to zh-TW_pronun
Private Const SORT_CHINESE_RADICALSTROKE  As Long = &H4 'Chinese radical/stroke order<7> 'Use with zh-HK, resolves to zh-HK_radstr 'Use with zh-MO, resolves to zh-MO_radstr 'Use with zh-TW, resolves to zh-TW_radstr
Private Const SORT_JAPANESE_RADICALSTROKE As Long = &H4 'Japanese radical/stroke sort order 'Use with ja-JP, resolves to ja-JP_radstr

Private Type NUMBERFMTW
    NumDigits        As Long    ' number of decimal digits
    LeadingZero      As Long    ' if leading zero decimal fields
    Grouping         As Long    ' group size left of decimal
    lpDecimalSep     As LongPtr 'String  ' ptr to decimal separator string
    lpThousandSep    As LongPtr 'String  ' ptr to thousand separator string
    NegativeOrder    As Long    ' negative number ordering
    PositiveOrder    As Long
    lpCurrencySymbol As LongPtr 'String
End Type

'typedef struct _numberfmtW {
'  UINT   NumDigits;
'  UINT   LeadingZero;
'  UINT   Grouping;
'  LPWSTR lpDecimalSep;
'  LPWSTR lpThousandSep;
'  UINT   NegativeOrder;
'} NUMBERFMTW, *LPNUMBERFMTW;
'
'Public Type NumberFormat
'    NumDigits     As Long   ' number of decimal digits
'    LeadingZero   As Long   ' if leading zero decimal fields
'    Grouping      As Long   ' group size left of decimal
'    DecimalSep    As String ' decimal separator string
'    ThousandSep   As String ' thousand separator string
'    NegativeOrder As Long   ' negative number ordering
'End Type

'typedef struct _currencyfmtW {
'  UINT   NumDigits;
'  UINT   LeadingZero;
'  UINT   Grouping;
'  LPWSTR lpDecimalSep;
'  LPWSTR lpThousandSep;
'  UINT   NegativeOrder;
'  UINT   PositiveOrder;
'  LPWSTR lpCurrencySymbol;
'} CURRENCYFMTW, *LPCURRENCYFMTW;

Private Type CURRENCYFMT
    NumDigits        As Long   ' number of decimal digits
    LeadingZero      As Long   ' if leading zero decimal fields
    Grouping         As Long   ' group size left of decimal
    lpDecimalSep     As String ' ptr to decimal separator string
    lpThousandSep    As String ' ptr to thousand separator string
    NegativeOrder    As Long   ' negative currency ordering
    PositiveOrder    As Long   ' positive currency ordering
    lpCurrencySymbol As String ' ptr to currency symbol string
End Type

'MLang.dll
'https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa741205(v=vs.85)
'Private Declare Function LcidToRfc1766W Lib "MLang" (ByVal LocaleID As Long, pSzRfc1766 As Any, ByVal nChar As Long) As Long

'https://learn.microsoft.com/de-de/windows/win32/api/winnls/nf-winnls-lcidtolocalename
Private Declare Function LCIDToLocaleName Lib "kernel32" (ByVal LocaleID As Long, lpName As Any, ByVal cchName As Long, ByVal dwFlags As Long) As Long
'int LCIDToLocaleName(
'  [in]            LCID   Locale,
'  [out, optional] LPWSTR lpName,
'  [in]            int    cchName,
'  [in]            DWORD  dwFlags
');
'LocaleID: lcid or LOCALE_INVARIANT, LOCALE_SYSTEM_DEFAULT, LOCALE_USER_DEFAULT
'lpName: pointer to string, or LOCALE_NAME_INVARIANT, LOCALE_NAME_SYSTEM_DEFAULT, LOCALE_NAME_USER_DEFAULT

'https://learn.microsoft.com/de-de/windows/win32/api/winnls/nf-winnls-localenametolcid
Private Declare Function LocaleNameToLCID Lib "kernel32" (ByVal lpName As Any, ByVal dwFlags As Long) As Long
'LCID LocaleNameToLCID(
'  [in] LPCWSTR lpName,
'  [in] DWORD   dwFlags
');


'https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa741208(v=vs.85)
'Private Declare Function Rfc1766ToLcidW Lib "MLang" (ByRef pLocaleID_out As Any, pSzRfc1766 As Any) As Long

'https://learn.microsoft.com/en-us/windows/win32/api/winver/nf-winver-verlanguagenamew
Private Declare Function VerLanguageNameW Lib "kernel32" (ByVal LocaleID As Long, ByRef pSzLang_out As Any, ByVal nSize As Long) As Long
'DWORD VerLanguageNameW(
'  [in]  DWORD  wLang,
'  [out] LPWSTR szLang,
'  [in]  DWORD  cchLang
');

Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Private Declare Function GetSystemDefaultLangID Lib "kernel32" () As Long
Private Declare Function GetSystemDefaultLocaleName Lib "kernel32" (ByVal lpLocaleName_out As LongPtr, ByVal cchLocaleName As Long) As Long

Private Declare Function GetUserDefaultUILanguage Lib "kernel32" () As Long

'https://learn.microsoft.com/en-us/windows/win32/api/winnls/nf-winnls-getlocaleinfow
Private Declare Function GetLocaleInfoW Lib "kernel32" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As LongPtr, ByVal cchData As Long) As Long

'https://learn.microsoft.com/en-us/windows/win32/api/winnls/nf-winnls-getlocaleinfoex
Private Declare Function GetLocaleInfoEx Lib "kernel32" (ByVal lpLocaleName As LongPtr, ByVal LCType As Long, ByVal lpLCData As LongPtr, ByVal cchData As Long) As Long
'int GetLocaleInfoEx(
'  [in, optional]  LPCWSTR lpLocaleName,
'  [in]            LCTYPE  LCType,
'  [out, optional] LPWSTR  lpLCData,
'  [in]            int     cchData
');

Private Declare Function GetNumberFormatW Lib "kernel32" (ByVal Locale As Long, ByVal dwFlags As Long, ByVal lpValue As LongPtr, lpFormat As LongPtr, ByVal lpNumberStr As LongPtr, ByVal cchNumber As Long) As Long
Private Declare Function GetCalendarInfoW Lib "kernel32" (ByVal Locale As Long, ByVal Calendar As Long, ByVal CalType As Long, ByVal lpCalData_out As LongPtr, ByVal cchData As Long, ByVal lpValue As LongPtr) As Long

'https://learn.microsoft.com/de-de/windows/win32/api/winnls/nf-winnls-getcurrencyformatw
Private Declare Function GetCurrencyFormatW Lib "kernel32" (ByVal Locale As Long, ByVal dwFlags As Long, ByVal lpValue As LongPtr, ByRef lpFormat As Any, ByVal lpCurrencyStr As LongPtr, ByVal cchCurrency As Long) As Long

'https://learn.microsoft.com/de-de/windows/win32/api/winnls/nf-winnls-getcurrencyformatex
Private Declare Function GetCurrencyFormatEx Lib "kernel32" (ByVal lpLocaleName As LongPtr, ByVal dwFlags As Long, ByVal lpValue As LongPtr, ByRef lpFormat As Any, lpCurrencyStr As LongPtr, ByVal cchCurrency As Long) As Long
'int GetCurrencyFormatEx(
'  [in, optional]  LPCWSTR            lpLocaleName,
'  [in]            DWORD              dwFlags,
'  [in]            LPCWSTR            lpValue,
'  [in, optional]  const CURRENCYFMTW *lpFormat,
'  [out, optional] LPWSTR             lpCurrencyStr,
'  [in]            int                cchCurrency
');

Private m_LCID   As Long   'RFC1766 LocaleID  437
Private m_Name   As String 'RFC1766 String    de-DE
Private m_LgName As String 'Sprachname        Deutsch (Deutschland)

Private m_parent                        As CultureInfo
Private m_userDefaultCulture            As CultureInfo
Private m_InvariantCultureInfo          As CultureInfo
Private m_userDefaultUICulture          As CultureInfo
Private m_InstalledUICultureInfo        As CultureInfo
Private m_DefaultThreadCurrentCulture   As CultureInfo
Private m_DefaultThreadCurrentUICulture As CultureInfo

'DateTimeFormat
Private m_DTF_MonthNames()              As String
Private m_DTF_DayNames()                As String
Private m_DTF_FirstDayOfWeek            As String
Private m_DTF_FullDateTimePattern       As String
Private m_DTF_AbbrevDayNames()          As String
Private m_DTF_AbbrevMonthGentvNames()   As String
Private m_DTF_AbbrevMonthNames()        As String
Private m_DTF_CalWeekRuleFirstDay       As String
Private m_DTF_CalWeekRuleFirst4DayWeek  As String
Private m_DTF_CalWeekRuleFirstFullWeek  As String
Private m_DTF_YearMonthlyPattern        As String
Private m_DTF_LongDatePattern           As String
Private m_DTF_LongTimePattern           As String
Private m_DTF_AllDateTimePatterns()     As String

Private m_NFT_CurcyDecDigits            As String
Private m_NFT_CurcyDecSeps              As String
Private m_CurrencyFmt                   As CURRENCYFMT


Friend Sub New_(ByVal aLCID As Long)
    m_LCID = aLCID
    m_Name = LCID_ToLocaleName(m_LCID)
    m_LgName = LCID_ToLanguageName(m_LCID)
End Sub

Friend Sub NewN(ByVal aLcidName As String)
    'If aLcidName = "ar-MR" Then
    '    Debug.Assert True
    'End If
    m_Name = aLcidName
    'If m_Name = "de-DE_phoneb" Then
    'If m_Name = "ar-MR" Then
    '    Debug.Assert False 'True
    'End If
    'm_LCID = Rfc1766String_ToLCID(aLcidName)
    m_LCID = CLng("&H" & LocaleInfo(LOCALE_ILANGUAGE))
    
    'Dim sortID As Long
    'sortID = CLng("&H" & LocaleInfo(LOCALE_SLANGUAGE))
    'm_LCID = MLang.MAKELCID(m_LCID, sortID)
    
    m_LgName = Trim(LocaleInfo(LOCALE_SLANGUAGE))
    If Len(m_LgName) = 0 Then m_LgName = LCID_ToLanguageName(m_LCID)
    If Len(m_LgName) = 0 Then m_LgName = LocaleInfo(LOCALE_SENGLANGUAGE)
    'erst am Schluss
    If Len(m_Name) = 0 Then m_Name = LCID_ToLocaleName(m_LCID)
End Sub

Friend Sub NewC(other As CultureInfo)
    m_LCID = other.lcid
    m_Name = other.Name
    m_LgName = other.LanguageName
End Sub

Public Function Clone() As CultureInfo
    Set Clone = New CultureInfo: Clone.NewC Me
End Function

Private Function LocaleInfo(ByVal ID As Long) As String
    'Dim Length As Long:   Length = GetLocaleInfoW(m_LCID, ID, 0, 0)
    Dim Length As Long:   Length = GetLocaleInfoEx(StrPtr(m_Name), ID, 0, Length)
    Dim Buffer As String: Buffer = Space(Length + 1) & vbNullChar & vbNullChar
    'Dim Result As Long:   Result = GetLocaleInfoW(m_LCID, ID, StrPtr(Buffer), Length)
    Dim Result As Long:   Result = GetLocaleInfoEx(StrPtr(m_Name), ID, StrPtr(Buffer), Length)
    LocaleInfo = Trim0(Left$(Buffer, Length))
End Function

'function _GetCountryFlag(): string;
'Var
'  Buffer : PChar;
'  Size : integer;
'begin
'  Size := GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVCTRYNAME, nil, 0);
'  GetMem(Buffer, Size);
'  Try
'    GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVCTRYNAME, Buffer, Size);
'    Result := Copy(Buffer,0,2); { return US / use  Result := Buffer to return USA }
'  finally
'    FreeMem(Buffer);
'  end;
'end;

Private Function IsValid() As Boolean
    IsValid = Len(m_Name)
End Function

Public Property Get LanguageName() As String
    LanguageName = m_LgName
End Property

' v ############################## v '    MLang    ' v ############################## v '
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
' ^ ############################## ^ '    MLang    ' ^ ############################## ^ '

Friend Function ToStr() As String
    'ToStr = LCID_ToHex & " = " & m_Name & "; Language: " & m_LgName
    ToStr = LCID_ToHex & " | " & MString.PadRight(m_Name, 15) & " | " & m_LgName
End Function

Friend Function ToInfoStr() As String
    Dim i As Long, s As String
    Dim maxlen As Long: maxlen = 28
    For i = 0 To 127 '90
        s = s & Format(i, "0000") & "| " & PadRight(ConstToStr(i), maxlen) & ": " & LocaleInfo(i) & vbCrLf
    Next
    For i = 4097 To 4116
        s = s & Format(i, "0000") & "| " & PadRight(ConstToStr(i), maxlen) & ": " & LocaleInfo(i) & vbCrLf
    Next
    ToInfoStr = s
End Function

Private Function Trim0(s As String) As String
    Trim0 = s
    Dim pos As Long: pos = InStr(1, Trim0, vbNullChar)
    If pos = 0 Then Exit Function
    Trim0 = Left$(Trim0, pos - 1)
End Function

'Static Properties
Public Property Get CurrentCulture() As CultureInfo
    '
End Property

Public Property Get UserDefaultCulture() As CultureInfo
    If m_userDefaultCulture Is Nothing Then Set m_userDefaultCulture = New CultureInfo: m_userDefaultCulture.New_ LOCALE_USER_DEFAULT
    Set UserDefaultCulture = m_userDefaultCulture
End Property

'Public Property Get CurrentUICulture() As CultureInfo
'    If m_userDefaultUICulture Is Nothing Then Set m_userDefaultUICulture = New CultureInfo: m_userDefaultUICulture.NewN
'    Set CurrentUICulture = m_userDefaultUICulture
'End Property
'
'Public Property Get InstalledUICulture() As CultureInfo
'    If m_InstalledUICultureInfo Is Nothing Then Set m_InstalledUICultureInfo = New CultureInfo: m_InstalledUICultureInfo.NewN
'    Set InstalledUICulture = m_InstalledUICultureInfo
'End Property
'
Public Property Get DefaultThreadCurrentCulture() As CultureInfo
    If m_DefaultThreadCurrentCulture Is Nothing Then Set m_DefaultThreadCurrentCulture = New CultureInfo: m_DefaultThreadCurrentCulture.New_ &H400
    Set DefaultThreadCurrentCulture = m_DefaultThreadCurrentCulture
End Property
'
'Public Property Get DefaultThreadCurrentUICulture() As CultureInfo
'    If m_DefaultThreadCurrentUICulture Is Nothing Then Set m_DefaultThreadCurrentUICulture = New CultureInfo: m_DefaultThreadCurrentUICulture.New_
'    Set DefaultThreadCurrentUICulture = m_DefaultThreadCurrentUICulture
'End Property

Public Property Get InvariantCulture() As CultureInfo
    If m_InvariantCultureInfo Is Nothing Then Set m_InvariantCultureInfo = New CultureInfo: m_InvariantCultureInfo.New_ &H7F 'LOCALE_INVARIANT
    Set InvariantCulture = m_InvariantCultureInfo
End Property

Public Property Get Parent() As CultureInfo
    If m_parent Is Nothing Then Set m_parent = New CultureInfo: m_parent.NewN LocaleInfo(LOCALE_SPARENT)
    Set Parent = m_parent
End Property

Public Property Get lcid() As Long
    lcid = m_LCID
End Property
Public Function LCID_ToHex() As String
    Dim sH As String: sH = Hex$(m_LCID)
    'sH = String$(4 - Len(sH), "0") & sH
    LCID_ToHex = "&H" & String$(4 - Len(sH), "0") & sH
End Function

Public Property Get Name() As String
    Name = m_Name ' en-EN ' en-US ' de-DE ' de-AT ' de-CH
End Property

Public Property Get AbbrevCountryName() As String
    AbbrevCountryName = LocaleInfo(LOCALE_SABBREVCOUNTRYNAME)
End Property

Public Property Get Calendar_ToStr() As String
    'GregorianCalendar
End Property

Private Function ConstToStr(ByVal ID As Long) As String
    Dim s As String
    Select Case ID
    Case LOCALE_NEUTRAL:                s = "LOCALE_NEUTRAL"                '&H0
    Case LOCALE_ILANGUAGE:              s = "LOCALE_ILANGUAGE"              '&H1
    Case LOCALE_SLANGUAGE:              s = "LOCALE_SLANG./LOC.DISP.NAME"   '&H2
    Case LOCALE_SABBREVLANGNAME:        s = "LOCALE_SABBREVLANGNAME"        '&H3
    Case LOCALE_SNATIVELANGUAGENAME:    s = "LOCALE_SNATIVELANGUAGENAME"    '&H4
    Case LOCALE_ICOUNTRY:               s = "LOCALE_ICOUNTRY/IDIALOGCODE"   '&H5
    Case LOCALE_SCOUNTRY:               s = "LOCALE_SCOUNTRY"               '&H6
    Case LOCALE_SABBREVCOUNTRYNAME:     s = "LOCALE_SABBREVCOUNTRYNAME"     '&H7
    Case LOCALE_SNATIVECOUNTRYNAME:     s = "LOCALE_SNATIVECOUNTRYNAME"     '&H8
    Case LOCALE_IDEFAULTLANGUAGE:       s = "LOCALE_IDEFAULTLANGUAGE"       '&H9
    Case LOCALE_IDEFAULTCOUNTRY:        s = "LOCALE_IDEFAULTCOUNTRY"        '&HA
    Case LOCALE_IDEFAULTCODEPAGE:       s = "LOCALE_IDEFAULTCODEPAGE"       '&HB
    Case LOCALE_SLIST:                  s = "LOCALE_SLIST"                  '&HC
    Case LOCALE_IMEASURE:               s = "LOCALE_IMEASURE"               '&HD
    Case LOCALE_SDECIMAL:               s = "LOCALE_SDECIMAL"               '&HE
    Case LOCALE_STHOUSAND:              s = "LOCALE_STHOUSAND"              '&HF
    Case LOCALE_SGROUPING:              s = "LOCALE_SGROUPING"              '&H10
    Case LOCALE_IDIGITS:                s = "LOCALE_IDIGITS"                '&H11
    Case LOCALE_ILZERO:                 s = "LOCALE_ILZERO"                 '&H12
    Case LOCALE_SNATIVEDIGITS:          s = "LOCALE_SNATIVEDIGITS"          '&H13
    Case LOCALE_SCURRENCY:              s = "LOCALE_SCURRENCY"              '&H14
    Case LOCALE_SINTLSYMBOL:            s = "LOCALE_SINTLSYMBOL"            '&H15
    Case LOCALE_SMONDECIMALSEP:         s = "LOCALE_SMONDECIMALSEP"         '&H16
    Case LOCALE_SMONTHOUSANDSEP:        s = "LOCALE_SMONTHOUSANDSEP"        '&H17
    Case LOCALE_SMONGROUPING:           s = "LOCALE_SMONGROUPING"           '&H18
    Case LOCALE_ICURRDIGITS:            s = "LOCALE_ICURRDIGITS"            '&H19
    Case LOCALE_IINTLCURRDIGITS:        s = "LOCALE_IINTLCURRDIGITS"        '&H1A
    Case LOCALE_ICURRENCY:              s = "LOCALE_ICURRENCY"              '&H1B
    Case LOCALE_INEGCURR:               s = "LOCALE_INEGCURR"               '&H1C
    Case LOCALE_SDATE:                  s = "LOCALE_SDATE"                  '&H1D
    Case LOCALE_STIME:                  s = "LOCALE_STIME"                  '&H1E
    Case LOCALE_SSHORTDATE:             s = "LOCALE_SSHORTDATE"             '&H1F
    Case LOCALE_SLONGDATE:              s = "LOCALE_SLONGDATE"              '&H20
    Case LOCALE_IDATE:                  s = "LOCALE_IDATE"                  '&H21
    Case LOCALE_ILDATE:                 s = "LOCALE_ILDATE"                 '&H22
    Case LOCALE_ITIME:                  s = "LOCALE_ITIME"                  '&H23
    Case LOCALE_ICENTURY:               s = "LOCALE_ICENTURY"               '&H24
    Case LOCALE_ITLZERO:                s = "LOCALE_ITLZERO"                '&H25
    Case LOCALE_IDAYLZERO:              s = "LOCALE_IDAYLZERO"              '&H26
    Case LOCALE_IMONLZERO:              s = "LOCALE_IMONLZERO"              '&H27
    Case LOCALE_S1159:                  s = "LOCALE_S1159/SAM"              '&H28
    Case LOCALE_S2359:                  s = "LOCALE_S2359/SPM"              '&H29
    Case LOCALE_SDAYNAME1:              s = "LOCALE_SDAYNAME1"              '&H2A
    Case LOCALE_SDAYNAME2:              s = "LOCALE_SDAYNAME2"              '&H2B
    Case LOCALE_SDAYNAME3:              s = "LOCALE_SDAYNAME3"              '&H2C
    Case LOCALE_SDAYNAME4:              s = "LOCALE_SDAYNAME4"              '&H2D
    Case LOCALE_SDAYNAME5:              s = "LOCALE_SDAYNAME5"              '&H2E
    Case LOCALE_SDAYNAME6:              s = "LOCALE_SDAYNAME6"              '&H2F
    Case LOCALE_SDAYNAME7:              s = "LOCALE_SDAYNAME7"              '&H30
    Case LOCALE_SABBREVDAYNAME1:        s = "LOCALE_SABBREVDAYNAME1"        '&H31
    Case LOCALE_SABBREVDAYNAME2:        s = "LOCALE_SABBREVDAYNAME2"        '&H32
    Case LOCALE_SABBREVDAYNAME3:        s = "LOCALE_SABBREVDAYNAME3"        '&H33
    Case LOCALE_SABBREVDAYNAME4:        s = "LOCALE_SABBREVDAYNAME4"        '&H34
    Case LOCALE_SABBREVDAYNAME5:        s = "LOCALE_SABBREVDAYNAME5"        '&H35
    Case LOCALE_SABBREVDAYNAME6:        s = "LOCALE_SABBREVDAYNAME6"        '&H36
    Case LOCALE_SABBREVDAYNAME7:        s = "LOCALE_SABBREVDAYNAME7"        '&H37
    Case LOCALE_SMONTHNAME1:            s = "LOCALE_SMONTHNAME1"            '&H38
    Case LOCALE_SMONTHNAME2:            s = "LOCALE_SMONTHNAME2"            '&H39
    Case LOCALE_SMONTHNAME3:            s = "LOCALE_SMONTHNAME3"            '&H3A
    Case LOCALE_SMONTHNAME4:            s = "LOCALE_SMONTHNAME4"            '&H3B
    Case LOCALE_SMONTHNAME5:            s = "LOCALE_SMONTHNAME5"            '&H3C
    Case LOCALE_SMONTHNAME6:            s = "LOCALE_SMONTHNAME6"            '&H3D
    Case LOCALE_SMONTHNAME7:            s = "LOCALE_SMONTHNAME7"            '&H3E
    Case LOCALE_SMONTHNAME8:            s = "LOCALE_SMONTHNAME8"            '&H3F
    Case LOCALE_SMONTHNAME9:            s = "LOCALE_SMONTHNAME9"            '&H40
    Case LOCALE_SMONTHNAME10:           s = "LOCALE_SMONTHNAME10"           '&H41
    Case LOCALE_SMONTHNAME11:           s = "LOCALE_SMONTHNAME11"           '&H42
    Case LOCALE_SMONTHNAME12:           s = "LOCALE_SMONTHNAME12"           '&H43
    Case LOCALE_SABBREVMONTHNAME1:      s = "LOCALE_SABBREVMONTHNAME1"      '&H44
    Case LOCALE_SABBREVMONTHNAME2:      s = "LOCALE_SABBREVMONTHNAME2"      '&H45
    Case LOCALE_SABBREVMONTHNAME3:      s = "LOCALE_SABBREVMONTHNAME3"      '&H46
    Case LOCALE_SABBREVMONTHNAME4:      s = "LOCALE_SABBREVMONTHNAME4"      '&H47
    Case LOCALE_SABBREVMONTHNAME5:      s = "LOCALE_SABBREVMONTHNAME5"      '&H48
    Case LOCALE_SABBREVMONTHNAME6:      s = "LOCALE_SABBREVMONTHNAME6"      '&H49
    Case LOCALE_SABBREVMONTHNAME7:      s = "LOCALE_SABBREVMONTHNAME7"      '&H4A
    Case LOCALE_SABBREVMONTHNAME8:      s = "LOCALE_SABBREVMONTHNAME8"      '&H4B
    Case LOCALE_SABBREVMONTHNAME9:      s = "LOCALE_SABBREVMONTHNAME9"      '&H4C
    Case LOCALE_SABBREVMONTHNAME10:     s = "LOCALE_SABBREVMONTHNAME10"     '&H4D
    Case LOCALE_SABBREVMONTHNAME11:     s = "LOCALE_SABBREVMONTHNAME11"     '&H4E
    Case LOCALE_SABBREVMONTHNAME12:     s = "LOCALE_SABBREVMONTHNAME12"     '&H4F
    Case LOCALE_SPOSITIVESIGN:          s = "LOCALE_SPOSITIVESIGN"          '&H50
    Case LOCALE_SNEGATIVESIGN:          s = "LOCALE_SNEGATIVESIGN"          '&H51
    Case LOCALE_IPOSSIGNPOSN:           s = "LOCALE_IPOSSIGNPOSN"           '&H52
    Case LOCALE_INEGSIGNPOSN:           s = "LOCALE_INEGSIGNPOSN"           '&H53
    Case LOCALE_IPOSSYMPRECEDES:        s = "LOCALE_IPOSSYMPRECEDES"        '&H54
    Case LOCALE_IPOSSEPBYSPACE:         s = "LOCALE_IPOSSEPBYSPACE"         '&H55
    Case LOCALE_INEGSYMPRECEDES:        s = "LOCALE_INEGSYMPRECEDES"        '&H56
    Case LOCALE_INEGSEPBYSPACE:         s = "LOCALE_INEGSEPBYSPACE"         '&H57
    Case LOCALE_FONTSIGNATURE:          s = "LOCALE_FONTSIGNATURE"          '&H58
    Case LOCALE_SISO639LANGNAME:        s = "LOCALE_SISO639LANGNAME"        '&H59
    Case LOCALE_SISO3166CTRYNAME:       s = "LOCALE_SISO3166CTRYNAME"       '&H5A
                                   
    Case LOCALE_IGEOID:                 s = "LOCALE_IGEOID"                 '&H5B
    Case LOCALE_SNAME:                  s = "LOCALE_SNAME"                  '&H5C
    Case LOCALE_SDURATION:              s = "LOCALE_SDURATION"              '&H5D
    Case LOCALE_SKEYBOARDSTOINSTALL:    s = "LOCALE_SKEYBOARDSTOINSTALL"    '&H5E
    Case LOCALE_SSHORTESTDAYNAME1:      s = "LOCALE_SSHORTESTDAYNAME1"      '&H60
    Case LOCALE_SSHORTESTDAYNAME2:      s = "LOCALE_SSHORTESTDAYNAME2"      '&H61
    Case LOCALE_SSHORTESTDAYNAME3:      s = "LOCALE_SSHORTESTDAYNAME3"      '&H62
    Case LOCALE_SSHORTESTDAYNAME4:      s = "LOCALE_SSHORTESTDAYNAME4"      '&H63
    Case LOCALE_SSHORTESTDAYNAME5:      s = "LOCALE_SSHORTESTDAYNAME5"      '&H64
    Case LOCALE_SSHORTESTDAYNAME6:      s = "LOCALE_SSHORTESTDAYNAME6"      '&H65
    Case LOCALE_SSHORTESTDAYNAME7:      s = "LOCALE_SSHORTESTDAYNAME7"      '&H66
    Case LOCALE_SISO639LANGNAME2:       s = "LOCALE_SISO639LANGNAME2"       '&H67
    Case LOCALE_SISO3166CTRYNAME2:      s = "LOCALE_SISO3166CTRYNAME2"      '&H68
    Case LOCALE_SNAN:                   s = "LOCALE_SNAN"                   '&H69
    Case LOCALE_SPOSINFINITY:           s = "LOCALE_SPOSINFINITY"           '&H6A
    Case LOCALE_SNEGINFINITY:           s = "LOCALE_SNEGINFINITY"           '&H6B
    Case LOCALE_SSCRIPTS:               s = "LOCALE_SSCRIPTS"               '&H6C
    Case LOCALE_SORTNAME:               s = "LOCALE_SORTNAME"               '&H6D
    
    Case LOCALE_SPARENT:                s = "LOCALE_SPARENT"                '&H6D
    
    Case LOCALE_SCONSOLEFALLBACKNAME:   s = "LOCALE_SCONSOLEFALLBACKNAME"   '&H6E ' new
    Case LOCALE_SLANGDISPLAYNAME:       s = "LOCALE_SLANGDISPLAYNAME"       '&H6F ' new
    Case LOCALE_SLOCALIZEDLANGUAGENAME: s = "LOCALE_SLOCALIZEDLANGUAGENAME" '&H6F ' new
    Case LOCALE_IREADINGLAYOUT:         s = "LOCALE_IREADINGLAYOUT"         '&H70 ' new
    Case LOCALE_INEUTRAL:               s = "LOCALE_INEUTRAL"               '&H71 ' new
    Case LOCALE_SENGLISHDISPLAYNAME:    s = "LOCALE_SENGLISHDISPLAYNAME"    '&H72 ' new
    Case LOCALE_SNATIVEDISPLAYNAME:     s = "LOCALE_SNATIVEDISPLAYNAME"     '&H73 ' new
    Case LOCALE_INEGATIVEPERCENT:       s = "LOCALE_INEGATIVEPERCENT"       '&H74 ' new
    Case LOCALE_IPOSITIVEPERCENT:       s = "LOCALE_IPOSITIVEPERCENT"       '&H75 ' new
    Case LOCALE_SPERCENT:               s = "LOCALE_SPERCENT"               '&H76 ' new
    Case LOCALE_SPERMILLE:              s = "LOCALE_SPERMILLE"              '&H77 ' new
    Case LOCALE_SMONTHDAY:              s = "LOCALE_SMONTHDAY"              '&H78 ' new
    Case LOCALE_SSHORTTIME:             s = "LOCALE_SSHORTTIME"             '&H79 ' new
    Case LOCALE_SOPENTYPELANGUAGETAG:   s = "LOCALE_SOPENTYPELANGUAGETAG"   '&H7A ' new
    Case LOCALE_SSORTLOCALE:            s = "LOCALE_SSORTLOCALE"            '&H7B ' new
    Case LOCALE_SRELATIVELONGDATE:      s = "LOCALE_SRELATIVELONGDATE"      '&H7C ' new
    'Case ?:                             s = "?"                            '&H7D ' new
    Case LOCALE_SSHORTESTAM:            s = "LOCALE_SSHORTESTAM"            '&H7E ' new
    Case LOCALE_SSHORTESTPM:            s = "LOCALE_SSHORTESTPM"            '&H7F ' new
    
    Case LOCALE_SENGLANGUAGE:           s = "LOCALE_SENGLANGUAGE"           '&H1001  ' 4097
    Case LOCALE_SENGCOUNTRY:            s = "LOCALE_SENGCOUNTRY"            '&H1002
    Case LOCALE_STIMEFORMAT:            s = "LOCALE_STIMEFORMAT"            '&H1003
    Case LOCALE_IDEFAULTANSICODEPAGE:   s = "LOCALE_IDEFAULTANSICODEPAGE"   '&H1004
    Case LOCALE_ITIMEMARKPOSN:          s = "LOCALE_ITIMEMARKPOSN"          '&H1005
    Case LOCALE_SYEARMONTH:             s = "LOCALE_SYEARMONTH"             '&H1006
    Case LOCALE_SENGCURRNAME:           s = "LOCALE_SENGCURRNAME"           '&H1007
    Case LOCALE_SNATIVECURRNAME:        s = "LOCALE_SNATIVECURRNAME"        '&H1008
    Case LOCALE_ICALENDARTYPE:          s = "LOCALE_ICALENDARTYPE"          '&H1009
    Case LOCALE_IPAPERSIZE:             s = "LOCALE_IPAPERSIZE"             '&H100A
    Case LOCALE_IOPTIONALCALENDAR:      s = "LOCALE_IOPTIONALCALENDAR"      '&H100B
    Case LOCALE_IFIRSTDAYOFWEEK:        s = "LOCALE_IFIRSTDAYOFWEEK"        '&H100C
    Case LOCALE_IFIRSTWEEKOFYEAR:       s = "LOCALE_IFIRSTWEEKOFYEAR"       '&H100D
    Case LOCALE_SMONTHNAME13:           s = "LOCALE_SMONTHNAME13"           '&H100E
    Case LOCALE_SABBREVMONTHNAME13:     s = "LOCALE_SABBREVMONTHNAME13"     '&H100F
    Case LOCALE_INEGNUMBER:             s = "LOCALE_INEGNUMBER"             '&H1010
    Case LOCALE_IDEFAULTMACCODEPAGE:    s = "LOCALE_IDEFAULTMACCODEPAGE"    '&H1011
    Case LOCALE_IDEFAULTEBCDICCODEPAGE: s = "LOCALE_IDEFAULTEBCDICCODEPAGE" '&H1012
    Case LOCALE_SSORTNAME:              s = "LOCALE_SSORTNAME"              '&H1013
    Case LOCALE_IDIGITSUBSTITUTION:     s = "LOCALE_IDIGITSUBSTITUTION"     '&H1014  '4116
    
    Case LOCALE_CUSTOM_UI_DEFAULT:      s = "LOCALE_CUSTOM_UI_DEFAULT"      '&H1400
    Case LOCALE_ALLOW_NEUTRAL_NAMES:    s = "LOCALE_ALLOW_NEUTRAL_NAMES"    '&H8000000
    Case LOCALE_RETURN_GENITIVE_NAMES:  s = "LOCALE_RETURN_GENITIVE_NAMES"  '&H10000000
    
    Case LOCALE_RETURN_NUMBER:          s = "LOCALE_RETURN_NUMBER"          '&H20000000
    Case LOCALE_USE_CP_ACP:             s = "LOCALE_USE_CP_ACP"             '&H40000000
    Case LOCALE_NOUSEROVERRIDE:         s = "LOCALE_NOUSEROVERRIDE"         '&H80000000
    
    End Select
    ConstToStr = s
End Function


'Calendar_ToStr As String                                    'System.Globalization.GregorianCalendar
'DateTimeFormat_ToStr As DateTimeFormatInfo                  '
'DisplayName As String                                       'Deutsch (Schweiz)
Public Property Get DisplayName() As String
    'DisplayName = LocaleInfo(LOCALE_SLANGUAGE)
    DisplayName = LocaleInfo(LOCALE_SNATIVEDISPLAYNAME)
End Property
'EnglishName As String                                       'German (Switzerland)
Public Property Get EnglishName() As String
    EnglishName = LocaleInfo(LOCALE_SENGLISHDISPLAYNAME)
End Property

'IsNeutralCulture As Boolean                                 'False
Public Property Get IsNeutralCulture() As Boolean
    'IsNeutralCulture = LocaleInfo(LOCALE_)
End Property
'KeyboardLayoutId As Integer                                 '2055
Public Property Get KeyboardLayoutId() As Long
Try: On Error GoTo Catch
    Dim s As String: s = LocaleInfo(LOCALE_SKEYBOARDSTOINSTALL)
    Dim sa() As String: sa = Split(s, ":")
    If UBound(sa) Then
        KeyboardLayoutId = CLng("&H" & sa(0))
    End If
Catch:
End Property
'=> m_cultureData.IINPUTLANGUAGEHANDLE;

'LCID As Integer                                             '2055
'Name As String                                              'de-CH
'NativeName As String                                        'Deutsch (Schweiz)
Public Property Get NativeName() As String
    NativeName = LocaleInfo(LOCALE_SNATIVELANGUAGENAME)
End Property

'NumberFormat As NumberFormatInfo                            '
'NumberFormat_ToStr As String                                'System.Globalization.NumberFormatInfo
'OptionalCalendars As Calendar()                             '
'OptionalCalendars_Length As Integer                         '1
'Parent As CultureInfo                                       '
'Parent.Name As String                                       'de
'TextInfo As TextInfo
'TextInfo.CultureName As String                              'de-CH
'TwoLetterISOLanguageName As String                          'de
Public Property Get TwoLetterISOLanguageName() As String
    TwoLetterISOLanguageName = LocaleInfo(LOCALE_SISO639LANGNAME)
End Property

'ThreeLetterISOLanguageName As String                        'deu
Public Property Get ThreeLetterISOLanguageName() As String
    ThreeLetterISOLanguageName = LocaleInfo(LOCALE_SISO639LANGNAME2)
End Property

'ThreeLetterWindowsLanguageName As String                    'DES
Public Property Get ThreeLetterWindowsLanguageName() As String
    ThreeLetterWindowsLanguageName = LocaleInfo(LOCALE_SABBREVLANGNAME)
End Property

'
'        s = s & ci.Calendar.ToString & vbCrLf
'        's = s & ci.Calendar. & vbCrLf
'        s = s & ci.DateTimeFormat.MonthNames.Count
Public Property Get DatFmt_MonthNames_Count() As Long
    
End Property
'        s = s & ci.DateTimeFormat.DayNames.Count
'        s = s & ci.DateTimeFormat.FirstDayOfWeek
'        s = s & ci.DateTimeFormat.FullDateTimePattern
'        s = s & ci.DateTimeFormat.AbbreviatedDayNames.Count
'        s = s & ci.DateTimeFormat.AbbreviatedMonthGenitiveNames.Count
'        s = s & ci.DateTimeFormat.AbbreviatedMonthNames.Count
'        s = s & ci.DateTimeFormat.CalendarWeekRule.FirstDay.ToString
'        s = s & ci.DateTimeFormat.CalendarWeekRule.FirstFourDayWeek
'        s = s & ci.DateTimeFormat.CalendarWeekRule.FirstFullWeek
'        s = s & ci.DateTimeFormat.YearMonthPattern
'        s = s & ci.DateTimeFormat.LongDatePattern
'        s = s & ci.DateTimeFormat.LongTimePattern
'        s = s & ci.DateTimeFormat.GetAllDateTimePatterns
'
'        s = s & ci.DisplayName & vbCrLf
'        s = s & ci.EnglishName & vbCrLf
'        s = s & ci.IsNeutralCulture & vbCrLf
'        s = s & ci.KeyboardLayoutId & vbCrLf
'        s = s & ci.LCID & vbCrLf
'        s = s & ci.Name & vbCrLf
'        s = s & ci.NativeName & vbCrLf
'
'        s = s & ci.NumberFormat.ToString & vbCrLf
'        s = s & ci.NumberFormat.CurrencyDecimalDigits & vbCrLf
'        s = s & ci.NumberFormat.CurrencyDecimalSeparator & vbCrLf
'        s = s & ci.NumberFormat.CurrencyGroupSeparator & vbCrLf
'        s = s & ci.NumberFormat.CurrencyGroupSizes.Count & vbCrLf
'        s = s & ci.NumberFormat.CurrencyNegativePattern & vbCrLf
'        s = s & ci.NumberFormat. & vbCrLf
'
'        s = s & ci.OptionalCalendars.Length & vbCrLf

'        s = s & ci.Parent.Name & vbCrLf
'        s = s & ci.Parent. & vbCrLf
'
'        s = s & ci.TextInfo.CultureName & vbCrLf
'        s = s & ci.TwoLetterISOLanguageName & vbCrLf
'        s = s & ci.ThreeLetterISOLanguageName & vbCrLf
'        s = s & ci.ThreeLetterWindowsLanguageName & vbCrLf
'
'        s = s & CultureInfo.CurrentCulture.Name & vbCrLf
'        s = s & CultureInfo.CurrentUICulture.Name & vbCrLf
'        s = s & CultureInfo.DefaultThreadCurrentCulture.Name & vbCrLf
'        s = s & CultureInfo.DefaultThreadCurrentUICulture.Name & vbCrLf
'
'        s = s & CultureInfo.InstalledUICulture.Name & vbCrLf
'        s = s & CultureInfo.InvariantCulture.Name
'        's = s & CultureInfo.

''NumberFormatInfo
''NumFmt
'
'Private NumFmt_InvariantInfo       As NumberFormatInfo
'Private m_numberGroupSizes()       As Long '() = Array( 3 )
'Private m_currencyGroupSizes()     As Long '() = Array( 3 )
'Private m_percentGroupSizes()      As Long '() = Array( 3 )
'Private m_positiveSign             As String '= "+"
'Private m_negativeSign             As String '= "-"
'Private m_numberDecimalSeparator   As String '= "."
'Private m_numberGroupSeparator     As String '= ","
'Private m_currencyGroupSeparator   As String '= ","
'Private m_currencyDecimalSeparator As String '= "."
'Private m_currencySymbol           As String '= "¤"
'Private m_ansiCurrencySymbol       As String
'Private m_nanSymbol                As String '= "NaN"
'Private m_positiveInfinitySymbol   As String '= "Infinity";
'Private m_negativeInfinitySymbol   As String '= "-Infinity";
'Private m_percentDecimalSeparator  As String '= ".";
'Private m_percentGroupSeparator    As String '= ",";
'Private m_percentSymbol            As String '= "%";
'Private m_perMilleSymbol           As String '= "‰"
''[OptionalField(VersionAdded = 2)]
'Private m_nativeDigits() As Strng '() = Array( "0", "1", "2", "3", "4", "5", "6", "7", "8", "9" )
''[OptionalField(VersionAdded = 1)]
'Private m_dataItem               As Long
'Private m_numberDecimalDigits      As Long ' = 2
'Private m_currencyDecimalDigits    As Long ' = 2
'Private m_CurrencyPositivePattern  As Long
'Private m_CurrencyNegativePattern  As Long
'Private m_numberNegativePattern    As Long ' = 1
'Private m_PercentPositivePattern   As Long
'Private m_PercentNegativePattern   As Long
'Private m_percentDecimalDigits     As Long ' = 2
'Private m_isReadOnly               As Boolean
'
'''[OptionalField(VersionAdded = 2)]
'''internal int digitSubstitution = 1
'''[OptionalField(VersionAdded = 1)]
'''internal bool m_useUserOverride
'''[OptionalField(VersionAdded = 2)]
'''internal bool m_isInvariant
'''[OptionalField(VersionAdded = 1)]
'''internal bool validForParseAsNumber = true
'''[OptionalField(VersionAdded = 1)]
'''internal bool validForParseAsCurrency = true;
'''private const NumberStyles InvalidNumberStyles = ~(NumberStyles.Any | NumberStyles.AllowHexSpecifier);
''
'Public Property Get NumFmt_InvariantInfo() As NumberFormatInfo
'    Set NumFmt_InvariantInfo = NumFmt_InvariantInfo
'End Property
''Public Property Get NumFmt_CurrentInfo() As NumberFormatInfo
''
''End Property
'Public Property Get NumFmt_CurrencyDecimalDigits() As Long
'    NumFmt_CurrencyDecimalDigits = m_currencyDecimalDigits
'End Property
'Public Property Get NumFmt_CurrencyDecimalSeparator() As String
'    NumFmt_CurrencyDecimalSeparator = m_currencyDecimalSeparator
'End Property
'Public Property Get NumFmt_IsReadOnly() As Boolean
'    NumFmt_IsReadOnly = m_isReadOnly
'End Property
'Public Property Get NumFmt_CurrencyGroupSizes() As Long()
'    NumFmt_CurrencyGroupSizes = m_currencyGroupSizes '()
'End Property
'Public Property Get NumFmt_NumberGroupSizes() As Long()
'    NumFmt_NumberGroupSizes = m_numberGroupSizes '()
'End Property
'Public Property Get NumFmt_PercentGroupSizes() As Long()
'    NumFmt_PercentGroupSizes = m_percentGroupSizes '()
'End Property
'Public Property Get NumFmt_CurrencyGroupSeparator() As String
'    NumFmt_CurrencyGroupSeparator = m_currencyGroupSeparator
'End Property
'Public Property Get NumFmt_CurrencySymbol() As String
'    NumFmt_CurrencySymbol = m_currencySymbol
'End Property
'Public Property Get NumFmt_NaNSymbol() As String
'    NumFmt_NaNSymbol = m_nanSymbol
'End Property
'Public Property Get NumFmt_CurrencyNegativePattern() As Long
'    NumFmt_CurrencyNegativePattern = m_CurrencyNegativePattern
'End Property
'Public Property Get NumFmt_NumberNegativePattern() As Long
'    NumFmt_NumberNegativePattern = m_numberNegativePattern
'End Property
'Public Property Get NumFmt_PercentPositivePattern() As Long
'    NumFmt_PercentPositivePattern = m_PercentPositivePattern
'End Property
'Public Property Get NumFmt_PercentNegativePattern() As Long
'    NumFmt_PercentNegativePattern = m_PercentNegativePattern
'End Property
'Public Property Get NumFmt_NegativeInfinitySymbol() As String
'    NumFmt_NegativeInfinitySymbol = m_negativeInfinitySymbol
'End Property
'Public Property Get NumFmt_NegativeSign() As String
'    NumFmt_NegativeSign = m_negativeSign
'End Property
'Public Property Get NumFmt_NumberDecimalDigits() As Long
'    NumFmt_NumberDecimalDigits = m_numberDecimalDigits
'End Property
'Public Property Get NumFmt_NumberDecimalSeparator() As String
'    NumFmt_NumberDecimalSeparator = m_numberDecimalSeparator
'End Property
'Public Property Get NumFmt_NumberGroupSeparator() As String
'    NumFmt_NumberGroupSeparator = m_numberGroupSeparator
'End Property
'Public Property Get NumFmt_CurrencyPositivePattern() As Long
'    NumFmt_CurrencyPositivePattern = m_CurrencyPositivePattern
'End Property
'Public Property Get NumFmt_PositiveInfinitySymbol() As String
'    NumFmt_PositiveInfinitySymbol = m_positiveInfinitySymbol
'End Property
'Public Property Get NumFmt_PositiveSign() As String
'    NumFmt_PositiveSign = m_positiveSign
'End Property
'Public Property Get NumFmt_PercentDecimalDigits() As Long
'    NumFmt_PercentDecimalDigits = m_percentDecimalDigits
'End Property
'Public Property Get NumFmt_PercentDecimalSeparator() As String
'    NumFmt_PercentDecimalSeparator = m_percentDecimalSeparator
'End Property
'Public Property Get NumFmt_PercentGroupSeparator() As String
'    NumFmt_PercentGroupSeparator = m_percentGroupSeparator
'End Property
'Public Property Get NumFmt_PercentSymbol() As String
'    NumFmt_PercentSymbol = m_percentSymbol
'End Property
'Public Property Get NumFmt_PerMilleSymbol() As String
'    NumFmt_PerMilleSymbol = m_perMilleSymbol
'End Property
'Public Property Get NumFmt_NativeDigits() As String()
'    NumFmt_NativeDigits = m_nativeDigits '()
'End Property
''Public Property Get NumFmt_DigitSubstitution() As DigitShapes
''    '
''End Property
