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

Private Const LOCALE_NEUTRAL                As Long = &H0
Private Const LOCALE_ILANGUAGE              As Long = &H1
Private Const LOCALE_SLANGUAGE              As Long = &H2
Private Const LOCALE_SABBREVLANGNAME        As Long = &H3
Private Const LOCALE_SNATIVELANGNAME        As Long = &H4
Private Const LOCALE_ICOUNTRY               As Long = &H5
Private Const LOCALE_SCOUNTRY               As Long = &H6
Private Const LOCALE_SABBREVCTRYNAME        As Long = &H7
Private Const LOCALE_SNATIVECTRYNAME        As Long = &H8
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
Private Const LOCALE_S2359                  As Long = &H29
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

Private Const LOCALE_SPARENT                As Long = &H6D
'Private Const LOCALE_NEUTRAL                As Long = 0
Private Const LOCALE_INVARIANT              As Long = &H7F     '    127

Private Const LOCALE_USER_DEFAULT           As Long = &H400    '   1024
Private Const LOCALE_TRADITIONAL_SPANISH    As Long = &H40A    '   1034
Private Const LOCALE_SYSTEM_DEFAULT         As Long = &H800    '   2048
Private Const LOCALE_CUSTOM_DEFAULT         As Long = &HC00    '   3072
Private Const LOCALE_CUSTOM_UNSPECIFIED     As Long = &H1000   '   4096
Private Const LOCALE_SORTID_MASK            As Long = &HF0000  ' 983040

'Private m_useUserOverride As Boolean

Private Const LOCALE_SENGLANGUAGE           As Long = &H1001
Private Const LOCALE_SENGCOUNTRY            As Long = &H1002
Private Const LOCALE_STIMEFORMAT            As Long = &H1003
Private Const LOCALE_IDEFAULTANSICODEPAGE   As Long = &H1004
Private Const LOCALE_ITIMEMARKPOSN          As Long = &H1005
Private Const LOCALE_SYEARMONTH             As Long = &H1006
Private Const LOCALE_SENGCURRNAME           As Long = &H1007
Private Const LOCALE_SNATIVECURRNAME        As Long = &H1008
Private Const LOCALE_ICALENDARTYPE          As Long = &H1009
Private Const LOCALE_IPAPERSIZE             As Long = &H100A
Private Const LOCALE_IOPTIONALCALENDAR      As Long = &H100B
Private Const LOCALE_IFIRSTDAYOFWEEK        As Long = &H100C
Private Const LOCALE_IFIRSTWEEKOFYEAR       As Long = &H100D
Private Const LOCALE_SMONTHNAME13           As Long = &H100E
Private Const LOCALE_SABBREVMONTHNAME13     As Long = &H100F
Private Const LOCALE_INEGNUMBER             As Long = &H1010
Private Const LOCALE_IDEFAULTMACCODEPAGE    As Long = &H1011
Private Const LOCALE_IDEFAULTEBCDICCODEPAGE As Long = &H1012
Private Const LOCALE_SSORTNAME              As Long = &H1013
Private Const LOCALE_IDIGITSUBSTITUTION     As Long = &H1014

Private Const LOCALE_RETURN_NUMBER          As Long = &H20000000
Private Const LOCALE_USE_CP_ACP             As Long = &H40000000
Private Const LOCALE_NOUSEROVERRIDE         As Long = &H80000000

Private Type NUMBERFMT
    NumDigits     As Long    ' number of decimal digits
    LeadingZero   As Long    ' if leading zero decimal fields
    Grouping      As Long    ' group size left of decimal
    lpDecimalSep  As String  ' ptr to decimal separator string
    lpThousandSep As String  ' ptr to thousand separator string
    NegativeOrder As Long    ' negative number ordering
End Type

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

Private Declare Function LcidToRfc1766W Lib "mlang" (ByVal LocaleID As Long, pSzRfc1766 As Any, ByVal nChar As Long) As Long
Private Declare Function Rfc1766ToLcidW Lib "mlang" (ByRef pLocaleID_out As Any, pSzRfc1766 As Any) As Long
Private Declare Function VerLanguageNameW Lib "kernel32" (ByVal LocaleID As Long, ByRef pSzLang_out As Any, ByVal nSize As Long) As Long

Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long

'https://learn.microsoft.com/en-us/windows/win32/api/winnls/nf-winnls-getlocaleinfow
Private Declare Function GetLocaleInfoW Lib "kernel32" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As LongPtr, ByVal cchData As Long) As Long

'https://learn.microsoft.com/en-us/windows/win32/api/winnls/nf-winnls-getlocaleinfoex
Private Declare Function GetLocaleInfoEx Lib "kernel32" (ByVal lpLocaleName As Long, ByVal LCType As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long

Private Declare Function GetNumberFormatW Lib "kernel32" (ByVal Locale As Long, ByVal dwFlags As Long, ByVal lpValue As LongPtr, lpFormat As LongPtr, ByVal lpNumberStr As LongPtr, ByVal cchNumber As Long) As Long
Private Declare Function GetCalendarInfoW Lib "kernel32" (ByVal Locale As Long, ByVal Calendar As Long, ByVal CalType As Long, ByVal lpCalData_out As LongPtr, ByVal cchData As Long, ByVal lpValue As LongPtr) As Long
Private Declare Function GetCurrencyFormatW Lib "kernel32" (ByVal Locale As Long, ByVal dwFlags As Long, ByVal lpValue As LongPtr, ByRef lpFormat As CURRENCYFMT, ByVal lpCurrencyStr As String, ByVal cchCurrency As Long) As Long

Private m_LCID   As Long   'RFC1766 LocaleID  437
Private m_Name   As String 'RFC1766 String    de-DE
Private m_LgName As String 'Sprachname        Deutsch (Deutschland)

Private m_parent                        As CultureInfo
Private m_userDefaultCulture            As CultureInfo
Private m_InvariantCultureInfo          As CultureInfo
Private m_userDefaultUICulture          As CultureInfo
Private m_InstalledUICultureInfo        As CultureInfo
Private m_DefaultThreadCurrentUICulture As CultureInfo
Private m_DefaultThreadCurrentCulture   As CultureInfo

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


Friend Sub New_(ByVal aLCID As Long)
    m_LCID = aLCID
    m_Name = LCID_ToRfc1766(m_LCID)
    m_LgName = LCID_ToLanguageName(m_LCID)
End Sub

Friend Sub NewN(ByVal aLcidName As String)
    m_Name = aLcidName
    m_LCID = Rfc1766String_ToLCID(aLcidName)
    m_LgName = LCID_ToLanguageName(m_LCID)
End Sub

Friend Sub NewC(other As CultureInfo)
    m_LCID = other.LCID
    m_Name = other.Name
    m_LgName = other.LanguageName
End Sub

Public Function Clone() As CultureInfo
    Set Clone = New CultureInfo: Clone.NewC Me
End Function

Private Function LocaleInfo(ByVal ID As Long) As String
    Dim Length As Long:   Length = GetLocaleInfoW(m_LCID, ID, 0, 0)
    Dim Buffer As String: Buffer = Space(Length + 1)
    Dim Result As Long:   Result = GetLocaleInfoW(LCID, ID, StrPtr(Buffer), Length)
    LocaleInfo = Trim0(Left$(Buffer, Length))
End Function

Private Function IsValid() As Boolean
    IsValid = Len(m_Name)
End Function

Public Property Get LanguageName() As String
    LanguageName = m_LgName
End Property

Private Function LCID_ToRfc1766(ByVal aLCID As Long) As String
    Dim s As String: s = Space$(6) 'String$(6, vbNullChar)
    Dim hr As Long: hr = LcidToRfc1766W(aLCID, ByVal StrPtr(s), 6)
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
    LCID_ToRfc1766 = Trim0(s)
End Function

Private Function LCID_ToLanguageName(ByVal aLCID As Long) As String
    Dim l As Long:   l = 256
    Dim s As String: s = String$(l, vbNullChar)
    l = VerLanguageNameW(aLCID, ByVal StrPtr(s), l)
    If l = 0 Then Exit Function
    LCID_ToLanguageName = Trim0(Left$(s, l))
End Function

Private Function Rfc1766String_ToLCID(ByVal sRfc1766 As String) As Long
    Dim hr As Long: hr = Rfc1766ToLcidW(Rfc1766String_ToLCID, ByVal StrPtr(sRfc1766))
End Function

Friend Function ToStr() As String
    'ToStr = LCID_ToHex & " = " & m_Name & "; Language: " & m_LgName
    ToStr = LCID_ToHex & " | " & MString.PadRight(m_Name, 5) & " | " & m_LgName
End Function

Friend Function ToInfoStr() As String
    Dim i As Long, s As String
    For i = 0 To 90
        s = s & PadRight(ConstToStr(i), 22) & ": " & LocaleInfo(i) & vbCrLf
    Next
    For i = 4097 To 4116
        s = s & PadRight(ConstToStr(i), 22) & ": " & LocaleInfo(i) & vbCrLf
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
    '
End Property

Public Property Get CurrentUICulture() As CultureInfo
    '
End Property

Public Property Get InstalledUICulture() As CultureInfo
    '
End Property

Public Property Get DefaultThreadCurrentCulture() As CultureInfo
    '
End Property

Public Property Get DefaultThreadCurrentUICulture() As CultureInfo
    '
End Property

Public Property Get InvariantCulture() As CultureInfo
    If m_InvariantCultureInfo Is Nothing Then Set m_InvariantCultureInfo = New CultureInfo: m_InvariantCultureInfo.New_ LOCALE_INVARIANT
    Set InvariantCulture = m_InvariantCultureInfo
End Property

Public Property Get Parent() As CultureInfo
    If m_parent Is Nothing Then
        Set m_parent = New CultureInfo: m_parent.New_ &H7F
    End If
    Set Parent = m_parent
End Property

Public Property Get LCID() As Long
    LCID = m_LCID
End Property
Public Function LCID_ToHex() As String
    Dim sH As String: sH = Hex$(LCID)
    'sH = String$(4 - Len(sH), "0") & sH
    LCID_ToHex = "&H" & String$(4 - Len(sH), "0") & sH
End Function

Public Property Get KeyboardLayoutId() As Long
    
End Property
'=> m_cultureData.IINPUTLANGUAGEHANDLE;

Public Property Get Name() As String
    Name = m_Name ' en-EN ' en-US ' de-DE ' de-AT ' de-CH
End Property

Public Property Get Calendar_ToStr() As String
    'GregorianCalendar
End Property

Private Function ConstToStr(ByVal ID As Long) As String
    Dim s As String
    Select Case ID
    Case LOCALE_NEUTRAL:               s = "NEUTRAL"               '&H0
    Case LOCALE_ILANGUAGE:             s = "ILANGUAGE"             '&H1
    Case LOCALE_SLANGUAGE:             s = "SLANGUAGE"             '&H2
    Case LOCALE_SABBREVLANGNAME:       s = "SABBREVLANGNAME"       '&H3
    Case LOCALE_SNATIVELANGNAME:       s = "SNATIVELANGNAME"       '&H4
    Case LOCALE_ICOUNTRY:              s = "ICOUNTRY"              '&H5
    Case LOCALE_SCOUNTRY:              s = "SCOUNTRY"              '&H6
    Case LOCALE_SABBREVCTRYNAME:       s = "SABBREVCTRYNAME"       '&H7
    Case LOCALE_SNATIVECTRYNAME:       s = "SNATIVECTRYNAME"       '&H8
    Case LOCALE_IDEFAULTLANGUAGE:      s = "IDEFAULTLANGUAGE"      '&H9
    Case LOCALE_IDEFAULTCOUNTRY:       s = "IDEFAULTCOUNTRY"       '&HA
    Case LOCALE_IDEFAULTCODEPAGE:      s = "IDEFAULTCODEPAGE"      '&HB
    Case LOCALE_SLIST:                 s = "SLIST"                 '&HC
    Case LOCALE_IMEASURE:              s = "IMEASURE"              '&HD
    Case LOCALE_SDECIMAL:              s = "SDECIMAL"              '&HE
    Case LOCALE_STHOUSAND:             s = "STHOUSAND"             '&HF
    Case LOCALE_SGROUPING:             s = "SGROUPING"             '&H10
    Case LOCALE_IDIGITS:               s = "IDIGITS"               '&H11
    Case LOCALE_ILZERO:                s = "ILZERO"                '&H12
    Case LOCALE_SNATIVEDIGITS:         s = "SNATIVEDIGITS"         '&H13
    Case LOCALE_SCURRENCY:             s = "SCURRENCY"             '&H14
    Case LOCALE_SINTLSYMBOL:           s = "SINTLSYMBOL"           '&H15
    Case LOCALE_SMONDECIMALSEP:        s = "SMONDECIMALSEP"        '&H16
    Case LOCALE_SMONTHOUSANDSEP:       s = "SMONTHOUSANDSEP"       '&H17
    Case LOCALE_SMONGROUPING:          s = "SMONGROUPING"          '&H18
    Case LOCALE_ICURRDIGITS:           s = "ICURRDIGITS"           '&H19
    Case LOCALE_IINTLCURRDIGITS:       s = "IINTLCURRDIGITS"       '&H1A
    Case LOCALE_ICURRENCY:             s = "ICURRENCY"             '&H1B
    Case LOCALE_INEGCURR:              s = "INEGCURR"              '&H1C
    Case LOCALE_SDATE:                 s = "SDATE"                 '&H1D
    Case LOCALE_STIME:                 s = "STIME"                 '&H1E
    Case LOCALE_SSHORTDATE:            s = "SSHORTDATE"            '&H1F
    Case LOCALE_SLONGDATE:             s = "SLONGDATE"             '&H20
    Case LOCALE_IDATE:                 s = "IDATE"                 '&H21
    Case LOCALE_ILDATE:                s = "ILDATE"                '&H22
    Case LOCALE_ITIME:                 s = "ITIME"                 '&H23
    Case LOCALE_ICENTURY:              s = "ICENTURY"              '&H24
    Case LOCALE_ITLZERO:               s = "ITLZERO"               '&H25
    Case LOCALE_IDAYLZERO:             s = "IDAYLZERO"             '&H26
    Case LOCALE_IMONLZERO:             s = "IMONLZERO"             '&H27
    Case LOCALE_S1159:                 s = "S1159"                 '&H28
    Case LOCALE_S2359:                 s = "S2359"                 '&H29
    Case LOCALE_SDAYNAME1:             s = "SDAYNAME1"             '&H2A
    Case LOCALE_SDAYNAME2:             s = "SDAYNAME2"             '&H2B
    Case LOCALE_SDAYNAME3:             s = "SDAYNAME3"             '&H2C
    Case LOCALE_SDAYNAME4:             s = "SDAYNAME4"             '&H2D
    Case LOCALE_SDAYNAME5:             s = "SDAYNAME5"             '&H2E
    Case LOCALE_SDAYNAME6:             s = "SDAYNAME6"             '&H2F
    Case LOCALE_SDAYNAME7:             s = "SDAYNAME7"             '&H30
    Case LOCALE_SABBREVDAYNAME1:       s = "SABBREVDAYNAME1"       '&H31
    Case LOCALE_SABBREVDAYNAME2:       s = "SABBREVDAYNAME2"       '&H32
    Case LOCALE_SABBREVDAYNAME3:       s = "SABBREVDAYNAME3"       '&H33
    Case LOCALE_SABBREVDAYNAME4:       s = "SABBREVDAYNAME4"       '&H34
    Case LOCALE_SABBREVDAYNAME5:       s = "SABBREVDAYNAME5"       '&H35
    Case LOCALE_SABBREVDAYNAME6:       s = "SABBREVDAYNAME6"       '&H36
    Case LOCALE_SABBREVDAYNAME7:       s = "SABBREVDAYNAME7"       '&H37
    Case LOCALE_SMONTHNAME1:           s = "SMONTHNAME1"           '&H38
    Case LOCALE_SMONTHNAME2:           s = "SMONTHNAME2"           '&H39
    Case LOCALE_SMONTHNAME3:           s = "SMONTHNAME3"           '&H3A
    Case LOCALE_SMONTHNAME4:           s = "SMONTHNAME4"           '&H3B
    Case LOCALE_SMONTHNAME5:           s = "SMONTHNAME5"           '&H3C
    Case LOCALE_SMONTHNAME6:           s = "SMONTHNAME6"           '&H3D
    Case LOCALE_SMONTHNAME7:           s = "SMONTHNAME7"           '&H3E
    Case LOCALE_SMONTHNAME8:           s = "SMONTHNAME8"           '&H3F
    Case LOCALE_SMONTHNAME9:           s = "SMONTHNAME9"           '&H40
    Case LOCALE_SMONTHNAME10:          s = "SMONTHNAME10"          '&H41
    Case LOCALE_SMONTHNAME11:          s = "SMONTHNAME11"          '&H42
    Case LOCALE_SMONTHNAME12:          s = "SMONTHNAME12"          '&H43
    Case LOCALE_SABBREVMONTHNAME1:     s = "SABBREVMONTHNAME1"     '&H44
    Case LOCALE_SABBREVMONTHNAME2:     s = "SABBREVMONTHNAME2"     '&H45
    Case LOCALE_SABBREVMONTHNAME3:     s = "SABBREVMONTHNAME3"     '&H46
    Case LOCALE_SABBREVMONTHNAME4:     s = "SABBREVMONTHNAME4"     '&H47
    Case LOCALE_SABBREVMONTHNAME5:     s = "SABBREVMONTHNAME5"     '&H48
    Case LOCALE_SABBREVMONTHNAME6:     s = "SABBREVMONTHNAME6"     '&H49
    Case LOCALE_SABBREVMONTHNAME7:     s = "SABBREVMONTHNAME7"     '&H4A
    Case LOCALE_SABBREVMONTHNAME8:     s = "SABBREVMONTHNAME8"     '&H4B
    Case LOCALE_SABBREVMONTHNAME9:     s = "SABBREVMONTHNAME9"     '&H4C
    Case LOCALE_SABBREVMONTHNAME10:    s = "SABBREVMONTHNAME10"    '&H4D
    Case LOCALE_SABBREVMONTHNAME11:    s = "SABBREVMONTHNAME11"    '&H4E
    Case LOCALE_SABBREVMONTHNAME12:    s = "SABBREVMONTHNAME12"    '&H4F
    Case LOCALE_SPOSITIVESIGN:         s = "SPOSITIVESIGN"         '&H50
    Case LOCALE_SNEGATIVESIGN:         s = "SNEGATIVESIGN"         '&H51
    Case LOCALE_IPOSSIGNPOSN:          s = "IPOSSIGNPOSN"          '&H52
    Case LOCALE_INEGSIGNPOSN:          s = "INEGSIGNPOSN"          '&H53
    Case LOCALE_IPOSSYMPRECEDES:       s = "IPOSSYMPRECEDES"       '&H54
    Case LOCALE_IPOSSEPBYSPACE:        s = "IPOSSEPBYSPACE"        '&H55
    Case LOCALE_INEGSYMPRECEDES:       s = "INEGSYMPRECEDES"       '&H56
    Case LOCALE_INEGSEPBYSPACE:        s = "INEGSEPBYSPACE"        '&H57
    Case LOCALE_FONTSIGNATURE:         s = "FONTSIGNATURE"         '&H58
    Case LOCALE_SISO639LANGNAME:       s = "SISO639LANGNAME"       '&H59
    Case LOCALE_SISO3166CTRYNAME:      s = "SISO3166CTRYNAME"      '&H5A
                                  
    Case LOCALE_SENGLANGUAGE:          s = "SENGLANGUAGE"          '&H1001  ' 4097
    Case LOCALE_SENGCOUNTRY:           s = "SENGCOUNTRY"           '&H1002
    Case LOCALE_STIMEFORMAT:           s = "STIMEFORMAT"           '&H1003
    Case LOCALE_IDEFAULTANSICODEPAGE:  s = "IDEFAULTANSICODEPAGE"  '&H1004
    Case LOCALE_ITIMEMARKPOSN:         s = "ITIMEMARKPOSN"         '&H1005
    Case LOCALE_SYEARMONTH:            s = "SYEARMONTH"            '&H1006
    Case LOCALE_SENGCURRNAME:          s = "SENGCURRNAME"          '&H1007
    Case LOCALE_SNATIVECURRNAME:       s = "SNATIVECURRNAME"       '&H1008
    Case LOCALE_ICALENDARTYPE:         s = "ICALENDARTYPE"         '&H1009
    Case LOCALE_IPAPERSIZE:            s = "IPAPERSIZE"            '&H100A
    Case LOCALE_IOPTIONALCALENDAR:     s = "IOPTIONALCALENDAR"     '&H100B
    Case LOCALE_IFIRSTDAYOFWEEK:       s = "IFIRSTDAYOFWEEK"       '&H100C
    Case LOCALE_IFIRSTWEEKOFYEAR:      s = "IFIRSTWEEKOFYEAR"      '&H100D
    Case LOCALE_SMONTHNAME13:          s = "SMONTHNAME13"          '&H100E
    Case LOCALE_SABBREVMONTHNAME13:    s = "SABBREVMONTHNAME13"    '&H100F
    Case LOCALE_INEGNUMBER:            s = "INEGNUMBER"            '&H1010
    Case LOCALE_IDEFAULTMACCODEPAGE:   s = "IDEFAULTMACCODEPAGE"   '&H1011
    Case LOCALE_IDEFAULTEBCDICCODEPAGE: s = "IDEFAULTEBCDICCODEPAGE" '&H1012
    Case LOCALE_SSORTNAME:             s = "SSORTNAME"             '&H1013
    Case LOCALE_IDIGITSUBSTITUTION:    s = "IDIGITSUBSTITUTION"    '&H1014  '4116
    End Select
    ConstToStr = s
End Function


'Calendar_ToStr As String                                    'System.Globalization.GregorianCalendar
'DateTimeFormat_ToStr As DateTimeFormatInfo                  '
'DisplayName As String                                       'Deutsch (Schweiz)
'EnglishName As String                                       'German (Switzerland)
'IsNeutralCulture As Boolean                                 'False
'KeyboardLayoutId As Integer                                 '2055
'LCID As Integer                                             '2055
'Name As String                                              'de-CH
'NativeName As String                                        'Deutsch (Schweiz)
'NumberFormat As NumberFormatInfo                            '
'NumberFormat_ToStr As String                                'System.Globalization.NumberFormatInfo
'OptionalCalendars As Calendar()                             '
'OptionalCalendars_Length As Integer                         '1
'Parent As CultureInfo                                       '
'Parent.Name As String                                       'de
'TextInfo As TextInfo
'TextInfo.CultureName As String                              'de-CH
'TwoLetterISOLanguageName As String                          'de
'ThreeLetterISOLanguageName As String                        'deu
'ThreeLetterWindowsLanguageName As String                    'DES
'
'        s = s & ci.Calendar.ToString & vbCrLf
'        's = s & ci.Calendar. & vbCrLf
'        s = s & ci.DateTimeFormat.MonthNames.Count
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
