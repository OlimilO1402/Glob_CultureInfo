VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "www.activevb.de"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command2 
      Caption         =   "Refesh"
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   0
      Width           =   1335
   End
   Begin VB.ListBox List2 
      Height          =   3375
      Left            =   5160
      TabIndex        =   2
      Top             =   360
      Width           =   4935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refesh"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dieser Source stammt von http://www.activevb.de
'und kann frei verwendet werden. F�r eventuelle Sch�den
'wird nicht gehaftet.

'Um Fehler oder Fragen zu kl�ren, nutzen Sie bitte unser Forum.
'Ansonsten viel Spa� und Erfolg mit diesem Source !

Option Explicit

Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long

'https://learn.microsoft.com/en-us/windows/win32/api/winnls/nf-winnls-getlocaleinfoa
Private Declare Function GetLocaleInfoA Lib "kernel32" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

'https://learn.microsoft.com/en-us/windows/win32/api/winnls/nf-winnls-getlocaleinfow
Private Declare Function GetLocaleInfoW Lib "kernel32" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long

'https://learn.microsoft.com/en-us/windows/win32/api/winnls/nf-winnls-getlocaleinfoex
'int GetLocaleInfoEx(
'  [in, optional]  LPCWSTR lpLocaleName,
'  [in]            LCTYPE  LCType,
'  [out, optional] LPWSTR  lpLCData,
'  [in]            int     cchData
');
Private Declare Function GetLocaleInfoEx Lib "kernel32" (ByVal lpLocaleName As Long, ByVal LCType As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long

Const LOCALE_ILANGUAGE = &H1
Const LOCALE_SLANGUAGE = &H2
Const LOCALE_SENGLANGUAGE = &H1001
Const LOCALE_SABBREVLANGNAME = &H3
Const LOCALE_SNATIVELANGNAME = &H4
Const LOCALE_ICOUNTRY = &H5
Const LOCALE_SCOUNTRY = &H6
Const LOCALE_SENGCOUNTRY = &H1002
Const LOCALE_SABBREVCTRYNAME = &H7
Const LOCALE_SNATIVECTRYNAME = &H8
Const LOCALE_IDEFAULTLANGUAGE = &H9
Const LOCALE_IDEFAULTCOUNTRY = &HA
Const LOCALE_IDEFAULTCODEPAGE = &HB
Const LOCALE_SLIST = &HC
Const LOCALE_IMEASURE = &HD
Const LOCALE_SDECIMAL = &HE
Const LOCALE_STHOUSAND = &HF
Const LOCALE_SGROUPING = &H10
Const LOCALE_IDIGITS = &H11
Const LOCALE_ILZERO = &H12
Const LOCALE_SNATIVEDIGITS = &H13
Const LOCALE_SCURRENCY = &H14
Const LOCALE_SINTLSYMBOL = &H15
Const LOCALE_SMONDECIMALSEP = &H16
Const LOCALE_SMONTHOUSANDSEP = &H17
Const LOCALE_SMONGROUPING = &H18
Const LOCALE_ICURRDIGITS = &H19
Const LOCALE_IINTLCURRDIGITS = &H1A
Const LOCALE_ICURRENCY = &H1B
Const LOCALE_INEGCURR = &H1C
Const LOCALE_SDATE = &H1D
Const LOCALE_STIME = &H1E
Const LOCALE_SSHORTDATE = &H1F
Const LOCALE_SLONGDATE = &H20
Const LOCALE_STIMEFORMAT = &H1003
Const LOCALE_IDATE = &H21
Const LOCALE_ILDATE = &H22
Const LOCALE_ITIME = &H23
Const LOCALE_ICENTURY = &H24
Const LOCALE_ITLZERO = &H25
Const LOCALE_IDAYLZERO = &H26
Const LOCALE_IMONLZERO = &H27
Const LOCALE_S1159 = &H28
Const LOCALE_S2359 = &H29
Const LOCALE_SDAYNAME1 = &H2A
Const LOCALE_SDAYNAME2 = &H2B
Const LOCALE_SDAYNAME3 = &H2C
Const LOCALE_SDAYNAME4 = &H2D
Const LOCALE_SDAYNAME5 = &H2E
Const LOCALE_SDAYNAME6 = &H2F
Const LOCALE_SDAYNAME7 = &H30
Const LOCALE_SABBREVDAYNAME1 = &H31
Const LOCALE_SABBREVDAYNAME2 = &H32
Const LOCALE_SABBREVDAYNAME3 = &H33
Const LOCALE_SABBREVDAYNAME4 = &H34
Const LOCALE_SABBREVDAYNAME5 = &H35
Const LOCALE_SABBREVDAYNAME6 = &H36
Const LOCALE_SABBREVDAYNAME7 = &H37
Const LOCALE_SMONTHNAME1 = &H38
Const LOCALE_SMONTHNAME2 = &H39
Const LOCALE_SMONTHNAME3 = &H3A
Const LOCALE_SMONTHNAME4 = &H3B
Const LOCALE_SMONTHNAME5 = &H3C
Const LOCALE_SMONTHNAME6 = &H3D
Const LOCALE_SMONTHNAME7 = &H3E
Const LOCALE_SMONTHNAME8 = &H3F
Const LOCALE_SMONTHNAME9 = &H40
Const LOCALE_SMONTHNAME10 = &H41
Const LOCALE_SMONTHNAME11 = &H42
Const LOCALE_SMONTHNAME12 = &H43
Const LOCALE_SABBREVMONTHNAME1 = &H44
Const LOCALE_SABBREVMONTHNAME2 = &H45
Const LOCALE_SABBREVMONTHNAME3 = &H46
Const LOCALE_SABBREVMONTHNAME4 = &H47
Const LOCALE_SABBREVMONTHNAME5 = &H48
Const LOCALE_SABBREVMONTHNAME6 = &H49
Const LOCALE_SABBREVMONTHNAME7 = &H4A
Const LOCALE_SABBREVMONTHNAME8 = &H4B
Const LOCALE_SABBREVMONTHNAME9 = &H4C
Const LOCALE_SABBREVMONTHNAME10 = &H4D
Const LOCALE_SABBREVMONTHNAME11 = &H4E
Const LOCALE_SABBREVMONTHNAME12 = &H4F
Const LOCALE_SABBREVMONTHNAME13 = &H100F
Const LOCALE_SPOSITIVESIGN = &H50
Const LOCALE_SNEGATIVESIGN = &H51
Const LOCALE_IPOSSIGNPOSN = &H52
Const LOCALE_INEGSIGNPOSN = &H53
Const LOCALE_IPOSSYMPRECEDES = &H54
Const LOCALE_IPOSSEPBYSPACE = &H55
Const LOCALE_INEGSYMPRECEDES = &H56
Const LOCALE_INEGSEPBYSPACE = &H57

Const LOCALE_USER_DEFAULT = &H400
Const LOCALE_SYSTEM_DEFAULT As Long = &H400

Dim LCID As Long

Private Sub Form_Load()
    LCID = GetSystemDefaultLCID
    Command1_Click
    Command2_Click
End Sub

Private Sub Command1_Click()
    List1.Clear
    
    GLI LOCALE_SLIST, "Listentrennzeichen"
    GLI LOCALE_IMEASURE, "0=metrisch, 1=US"
    GLI LOCALE_SDECIMAL, "Dezimaltrennzeichen"
    GLI LOCALE_STHOUSAND, "Tausendertrennzeichen"
    GLI LOCALE_SGROUPING, "Gruppierung links vom Komma"
    GLI LOCALE_IDIGITS, "Zahlen hinter dem Komma"
    GLI LOCALE_ILZERO, "f�hrende Nullen"
    GLI LOCALE_SCURRENCY, "W�hrungsymbol"
    GLI LOCALE_SINTLSYMBOL, "W�hrung nach ISO 4217"
    GLI LOCALE_SMONDECIMALSEP, "W�hrungstrennzeichen"
    GLI LOCALE_SMONTHOUSANDSEP, "W�hrungstausendertrennzeichen"
    GLI LOCALE_SMONGROUPING, "W�hrungsgruppierung"
    GLI LOCALE_ICURRDIGITS, "Zahlen hinter dem Komma (Pf)"
    GLI LOCALE_ICURRENCY, "Anzeige des W�hrungssymbols"
    GLI LOCALE_INEGCURR, "Negatives W�hrungsvorzeichen"
    GLI LOCALE_SDATE, "Datumstrennzeichen"
    GLI LOCALE_STIME, "Zeittrennzeichen"
    GLI LOCALE_SSHORTDATE, "Kurzes Datumsformat"
    GLI LOCALE_SLONGDATE, "Langes Datumsformat"
    GLI LOCALE_STIMEFORMAT, "Zeitformat"
    GLI LOCALE_ITIME, "12/24 Stunden"
    GLI LOCALE_S1159, "AM-Zeichen"
    GLI LOCALE_S2359, "PM-Zeichen"
    GLI LOCALE_SPOSITIVESIGN, "Positives Vorz."
    GLI LOCALE_SNEGATIVESIGN, "Negatives Vorz."
    GLI LOCALE_ILANGUAGE, "Sprach ID"
    GLI LOCALE_SLANGUAGE, "Lokalisierter Sprachname"
    GLI LOCALE_SENGLANGUAGE, "Engl. �quivalent"
    GLI LOCALE_SABBREVLANGNAME, "Abgek�rzt"
    GLI LOCALE_SNATIVELANGNAME, "Sprache in Landessprache"
    GLI LOCALE_ICOUNTRY, "L�ndercode"
    GLI LOCALE_SCOUNTRY, "L�ndername"
    GLI LOCALE_SENGCOUNTRY, "L�ndername in Engl."
    GLI LOCALE_SABBREVCTRYNAME, "Abgek�rzt"
    GLI LOCALE_SNATIVECTRYNAME, "Land in Landessprache"
    GLI LOCALE_IDEFAULTLANGUAGE, "Standard Sprach-ID"
    GLI LOCALE_IDEFAULTCOUNTRY, "Standard Landes-ID"
    GLI LOCALE_IDEFAULTCODEPAGE, "Standard Codeseite"
    GLI LOCALE_SNATIVEDIGITS, "gebr�uchliche Zahlen"
    GLI LOCALE_IINTLCURRDIGITS, "Zahlen hinter Komma nach ISO"
    GLI LOCALE_IDATE, "Datums Gruppierung"
    GLI LOCALE_ILDATE, "Reihenfolge langes Datumsformat"
    GLI LOCALE_ICENTURY, "Jahr in 2/4 Ziffern"
    GLI LOCALE_ITLZERO, "f�hrende Null f�r Zeiten"
    GLI LOCALE_IDAYLZERO, "f�hrende Null f�r Tage"
    GLI LOCALE_IMONLZERO, "f�hrende Null f�r Monate"
    GLI LOCALE_SDAYNAME1, "Langer Name f�r Mo"
    GLI LOCALE_SDAYNAME2, "Langer Name f�r Di"
    GLI LOCALE_SDAYNAME3, "Langer Name f�r Mi"
    GLI LOCALE_SDAYNAME4, "Langer Name f�r Do"
    GLI LOCALE_SDAYNAME5, "Langer Name f�r Fr"
    GLI LOCALE_SDAYNAME6, "Langer Name f�r Sa"
    GLI LOCALE_SDAYNAME7, "Langer Name f�r So"
    GLI LOCALE_SABBREVDAYNAME1, "Abk�rzung f�r Mo"
    GLI LOCALE_SABBREVDAYNAME2, "Abk�rzung f�r Di"
    GLI LOCALE_SABBREVDAYNAME3, "Abk�rzung f�r Mi"
    GLI LOCALE_SABBREVDAYNAME4, "Abk�rzung f�r Do"
    GLI LOCALE_SABBREVDAYNAME5, "Abk�rzung f�r Fr"
    GLI LOCALE_SABBREVDAYNAME6, "Abk�rzung f�r Sa"
    GLI LOCALE_SABBREVDAYNAME7, "Abk�rzung f�r So"
    GLI LOCALE_SMONTHNAME1, "Langer Name f�r Jan"
    GLI LOCALE_SMONTHNAME2, "Langer Name f�r Feb"
    GLI LOCALE_SMONTHNAME3, "Langer Name f�r Mae"
    GLI LOCALE_SMONTHNAME4, "Langer Name f�r Mai"
    GLI LOCALE_SMONTHNAME5, "Langer Name f�r Apr"
    GLI LOCALE_SMONTHNAME6, "Langer Name f�r Jun"
    GLI LOCALE_SMONTHNAME7, "Langer Name f�r Jul"
    GLI LOCALE_SMONTHNAME8, "Langer Name f�r Aug"
    GLI LOCALE_SMONTHNAME9, "Langer Name f�r Sep"
    GLI LOCALE_SMONTHNAME10, "Langer Name f�r Okt"
    GLI LOCALE_SMONTHNAME11, "Langer Name f�r Nov"
    GLI LOCALE_SMONTHNAME12, "Langer Name f�r Dez"
    GLI LOCALE_SABBREVMONTHNAME1, "Abk�rzung f�r Jan"
    GLI LOCALE_SABBREVMONTHNAME2, "Abk�rzung f�r Feb"
    GLI LOCALE_SABBREVMONTHNAME3, "Abk�rzung f�r Mae"
    GLI LOCALE_SABBREVMONTHNAME4, "Abk�rzung f�r Apr"
    GLI LOCALE_SABBREVMONTHNAME5, "Abk�rzung f�r Mai"
    GLI LOCALE_SABBREVMONTHNAME6, "Abk�rzung f�r Jun"
    GLI LOCALE_SABBREVMONTHNAME7, "Abk�rzung f�r Jul"
    GLI LOCALE_SABBREVMONTHNAME8, "Abk�rzung f�r Aug"
    GLI LOCALE_SABBREVMONTHNAME9, "Abk�rzung f�r Sep"
    GLI LOCALE_SABBREVMONTHNAME10, "Abk�rzung f�r Okt"
    GLI LOCALE_SABBREVMONTHNAME11, "Abk�rzung f�r Nov"
    GLI LOCALE_SABBREVMONTHNAME12, "Abk�rzung f�r Dez"
    GLI LOCALE_IPOSSIGNPOSN, "Format f�r pos. W�hrung"
    GLI LOCALE_INEGSIGNPOSN, "Format f�r neg. W�hrung"
    GLI LOCALE_IPOSSYMPRECEDES, "Pr�fix f�r pos. W�hrungsvorzeichen"
    GLI LOCALE_IPOSSEPBYSPACE, "Trennz bei pos. W�hrungsbetrag"
    GLI LOCALE_INEGSYMPRECEDES, "Pr�fix f�r neg. W�hrungsvorzeichen"
    GLI LOCALE_INEGSEPBYSPACE, "Trennz. bei neg. W�hrungsbetrag"
End Sub

Private Sub Command2_Click()
    List2.Clear
    
    GLIW LOCALE_SLIST, "Listentrennzeichen"
    GLIW LOCALE_IMEASURE, "0=metrisch, 1=US"
    GLIW LOCALE_SDECIMAL, "Dezimaltrennzeichen"
    GLIW LOCALE_STHOUSAND, "Tausendertrennzeichen"
    GLIW LOCALE_SGROUPING, "Gruppierung links vom Komma"
    GLIW LOCALE_IDIGITS, "Zahlen hinter dem Komma"
    GLIW LOCALE_ILZERO, "f�hrende Nullen"
    GLIW LOCALE_SCURRENCY, "W�hrungsymbol"
    GLIW LOCALE_SINTLSYMBOL, "W�hrung nach ISO 4217"
    GLIW LOCALE_SMONDECIMALSEP, "W�hrungstrennzeichen"
    GLIW LOCALE_SMONTHOUSANDSEP, "W�hrungstausendertrennzeichen"
    GLIW LOCALE_SMONGROUPING, "W�hrungsgruppierung"
    GLIW LOCALE_ICURRDIGITS, "Zahlen hinter dem Komma (Pf)"
    GLIW LOCALE_ICURRENCY, "Anzeige des W�hrungssymbols"
    GLIW LOCALE_INEGCURR, "Negatives W�hrungsvorzeichen"
    GLIW LOCALE_SDATE, "Datumstrennzeichen"
    GLIW LOCALE_STIME, "Zeittrennzeichen"
    GLIW LOCALE_SSHORTDATE, "Kurzes Datumsformat"
    GLIW LOCALE_SLONGDATE, "Langes Datumsformat"
    GLIW LOCALE_STIMEFORMAT, "Zeitformat"
    GLIW LOCALE_ITIME, "12/24 Stunden"
    GLIW LOCALE_S1159, "AM-Zeichen"
    GLIW LOCALE_S2359, "PM-Zeichen"
    GLIW LOCALE_SPOSITIVESIGN, "Positives Vorz."
    GLIW LOCALE_SNEGATIVESIGN, "Negatives Vorz."
    GLIW LOCALE_ILANGUAGE, "Sprach ID"
    GLIW LOCALE_SLANGUAGE, "Lokalisierter Sprachname"
    GLIW LOCALE_SENGLANGUAGE, "Engl. �quivalent"
    GLIW LOCALE_SABBREVLANGNAME, "Abgek�rzt"
    GLIW LOCALE_SNATIVELANGNAME, "Sprache in Landessprache"
    GLIW LOCALE_ICOUNTRY, "L�ndercode"
    GLIW LOCALE_SCOUNTRY, "L�ndername"
    GLIW LOCALE_SENGCOUNTRY, "L�ndername in Engl."
    GLIW LOCALE_SABBREVCTRYNAME, "Abgek�rzt"
    GLIW LOCALE_SNATIVECTRYNAME, "Land in Landessprache"
    GLIW LOCALE_IDEFAULTLANGUAGE, "Standard Sprach-ID"
    GLIW LOCALE_IDEFAULTCOUNTRY, "Standard Landes-ID"
    GLIW LOCALE_IDEFAULTCODEPAGE, "Standard Codeseite"
    GLIW LOCALE_SNATIVEDIGITS, "gebr�uchliche Zahlen"
    GLIW LOCALE_IINTLCURRDIGITS, "Zahlen hinter Komma nach ISO"
    GLIW LOCALE_IDATE, "Datums Gruppierung"
    GLIW LOCALE_ILDATE, "Reihenfolge langes Datumsformat"
    GLIW LOCALE_ICENTURY, "Jahr in 2/4 Ziffern"
    GLIW LOCALE_ITLZERO, "f�hrende Null f�r Zeiten"
    GLIW LOCALE_IDAYLZERO, "f�hrende Null f�r Tage"
    GLIW LOCALE_IMONLZERO, "f�hrende Null f�r Monate"
    GLIW LOCALE_SDAYNAME1, "Langer Name f�r Mo"
    GLIW LOCALE_SDAYNAME2, "Langer Name f�r Di"
    GLIW LOCALE_SDAYNAME3, "Langer Name f�r Mi"
    GLIW LOCALE_SDAYNAME4, "Langer Name f�r Do"
    GLIW LOCALE_SDAYNAME5, "Langer Name f�r Fr"
    GLIW LOCALE_SDAYNAME6, "Langer Name f�r Sa"
    GLIW LOCALE_SDAYNAME7, "Langer Name f�r So"
    GLIW LOCALE_SABBREVDAYNAME1, "Abk�rzung f�r Mo"
    GLIW LOCALE_SABBREVDAYNAME2, "Abk�rzung f�r Di"
    GLIW LOCALE_SABBREVDAYNAME3, "Abk�rzung f�r Mi"
    GLIW LOCALE_SABBREVDAYNAME4, "Abk�rzung f�r Do"
    GLIW LOCALE_SABBREVDAYNAME5, "Abk�rzung f�r Fr"
    GLIW LOCALE_SABBREVDAYNAME6, "Abk�rzung f�r Sa"
    GLIW LOCALE_SABBREVDAYNAME7, "Abk�rzung f�r So"
    GLIW LOCALE_SMONTHNAME1, "Langer Name f�r Jan"
    GLIW LOCALE_SMONTHNAME2, "Langer Name f�r Feb"
    GLIW LOCALE_SMONTHNAME3, "Langer Name f�r Mae"
    GLIW LOCALE_SMONTHNAME4, "Langer Name f�r Mai"
    GLIW LOCALE_SMONTHNAME5, "Langer Name f�r Apr"
    GLIW LOCALE_SMONTHNAME6, "Langer Name f�r Jun"
    GLIW LOCALE_SMONTHNAME7, "Langer Name f�r Jul"
    GLIW LOCALE_SMONTHNAME8, "Langer Name f�r Aug"
    GLIW LOCALE_SMONTHNAME9, "Langer Name f�r Sep"
    GLIW LOCALE_SMONTHNAME10, "Langer Name f�r Okt"
    GLIW LOCALE_SMONTHNAME11, "Langer Name f�r Nov"
    GLIW LOCALE_SMONTHNAME12, "Langer Name f�r Dez"
    GLIW LOCALE_SABBREVMONTHNAME1, "Abk�rzung f�r Jan"
    GLIW LOCALE_SABBREVMONTHNAME2, "Abk�rzung f�r Feb"
    GLIW LOCALE_SABBREVMONTHNAME3, "Abk�rzung f�r Mae"
    GLIW LOCALE_SABBREVMONTHNAME4, "Abk�rzung f�r Apr"
    GLIW LOCALE_SABBREVMONTHNAME5, "Abk�rzung f�r Mai"
    GLIW LOCALE_SABBREVMONTHNAME6, "Abk�rzung f�r Jun"
    GLIW LOCALE_SABBREVMONTHNAME7, "Abk�rzung f�r Jul"
    GLIW LOCALE_SABBREVMONTHNAME8, "Abk�rzung f�r Aug"
    GLIW LOCALE_SABBREVMONTHNAME9, "Abk�rzung f�r Sep"
    GLIW LOCALE_SABBREVMONTHNAME10, "Abk�rzung f�r Okt"
    GLIW LOCALE_SABBREVMONTHNAME11, "Abk�rzung f�r Nov"
    GLIW LOCALE_SABBREVMONTHNAME12, "Abk�rzung f�r Dez"
    GLIW LOCALE_IPOSSIGNPOSN, "Format f�r pos. W�hrung"
    GLIW LOCALE_INEGSIGNPOSN, "Format f�r neg. W�hrung"
    GLIW LOCALE_IPOSSYMPRECEDES, "Pr�fix f�r pos. W�hrungsvorzeichen"
    GLIW LOCALE_IPOSSEPBYSPACE, "Trennz bei pos. W�hrungsbetrag"
    GLIW LOCALE_INEGSYMPRECEDES, "Pr�fix f�r neg. W�hrungsvorzeichen"
    GLIW LOCALE_INEGSEPBYSPACE, "Trennz. bei neg. W�hrungsbetrag"
End Sub

Private Sub GLI(ByVal ID As Long, Text As String)
    List1.AddItem Text & ":  " & GetEntry(ID)
    List1.ItemData(List1.NewIndex) = ID
End Sub

Private Sub GLIW(ByVal ID As Long, Text As String)
    List2.AddItem Text & ":  " & GetEntryW(ID)
    List2.ItemData(List2.NewIndex) = ID
End Sub

Private Function GetEntry(ByVal ID As Long) As String
    Dim Buffer As String
    Dim Length As Long:   Length = GetLocaleInfoA(LCID, ID, Buffer, 0) - 1
    Buffer = Space(Length + 1)
    Dim Result As Long:   Result = GetLocaleInfoA(LCID, ID, Buffer, Length)
    GetEntry = Left$(Buffer, Length)
End Function

'Private Function GetEntryW(ByVal ID As Long) As String
'    Dim Buffer As String
'    Dim Length As Long:   Length = GetLocaleInfoW(LCID, ID, StrPtr(Buffer), 0)
'    Buffer = Space(Length + 1)
'    Dim Result As Long:   Result = GetLocaleInfoW(LCID, ID, StrPtr(Buffer), Length)
'    GetEntryW = Left$(Buffer, Length)
'End Function
'
Private Function GetEntryW(ByVal ID As Long) As String
    Dim Length As Long:   Length = GetLocaleInfoW(LCID, ID, 0, 0)
    Dim Buffer As String: Buffer = Space(Length + 1)
    Dim Result As Long:   Result = GetLocaleInfoW(LCID, ID, StrPtr(Buffer), Length)
    GetEntryW = Left$(Buffer, Length)
End Function

