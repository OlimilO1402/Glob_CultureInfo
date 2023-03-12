Attribute VB_Name = "Module1"
'Dieser Source stammt von http://www.activevb.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.

'Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
'Ansonsten viel Spaß und Erfolg mit diesem Source !
Option Explicit

Private Declare Function EnumSystemLocalesA Lib "kernel32" (ByVal lpLocaleEnumProc As Long, ByVal dwFlags As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetLocaleInfoA Lib "kernel32" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Const LCID_INSTALLED As Long = &H1
Const LCID_SUPPORTED As Long = &H2
Const LCID_ALTERNATE_SORTS As Long = &H4

Const LOCALE_SNATIVELANGNAME As Long = &H4
Const LOCALE_SNATIVECTRYNAME As Long = &H8

Dim LCID() As Long

Public Sub EnumLocales(ByVal Mode As Integer, LB As ListBox)
    Dim Flag As Long
    Select Case Mode
    Case 0:    Flag = LCID_INSTALLED
    Case 1:    Flag = LCID_SUPPORTED
    Case 2:    Flag = LCID_ALTERNATE_SORTS
    Case Else: Flag = 0
    End Select
    If Flag Then
        ReDim LCID(0 To 0)
        LB.Clear
        EnumSystemLocalesA AddressOf LocaleEnumProc, Flag
        Dim x As Long, aa As String
        For x = 0 To UBound(LCID) - 1
            aa = CStr(LCID(x)) & " "
            aa = aa & GetEntry(LCID(x), LOCALE_SNATIVECTRYNAME) & " ["
            aa = aa & GetEntry(LCID(x), LOCALE_SNATIVELANGNAME) & "]"
            LB.AddItem aa
        Next
    Else
        MsgBox "Diese Option wird nicht unterstützt!"
    End If
End Sub

Private Function LocaleEnumProc(LCID_Pointer As Long) As Long
    Dim Buffer As String: Buffer = Space$(255)
    RtlMoveMemory ByVal Buffer, LCID_Pointer, Len(Buffer)
    Buffer = "&H" & Left$(Buffer, InStr(Buffer, Chr$(0)) - 1)
    LCID(UBound(LCID)) = CLng(Buffer)
    ReDim Preserve LCID(0 To UBound(LCID) + 1)
    LocaleEnumProc = 1&
End Function

Private Function GetEntry(LCID As Long, ID As Long) As String
    Dim Result As Long, Buffer As String, Length As Long
    Length = GetLocaleInfoA(LCID, ID, Buffer, 0) - 1
    If Length > 0 Then
        Buffer = Space(Length + 1)
        Result = GetLocaleInfoA(LCID, ID, Buffer, Length)
        GetEntry = Left$(Buffer, Length)
    End If
End Function
