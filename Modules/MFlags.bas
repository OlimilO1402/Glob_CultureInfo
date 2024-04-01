Attribute VB_Name = "MFlags"
Option Explicit
Private m_Path  As String
Private m_dPFN  As String
Private m_data  As Collection
Private m_Flags As Collection

Public Function ReadFlagPics() As Boolean
    m_Path = App.Path & "\Resources\Flagsbmp\"
    Set m_Flags = New Collection
    Dim pfn As String, PFNs() As String: PFNs = ReadDataFile
    Dim sISO() As String, key As String
    Dim i As Long, spic As StdPicture
    For i = 0 To UBound(PFNs)
        pfn = PFNs(i)
        sISO = Split(pfn, ", ")
        key = sISO(1)
        Set spic = LoadPicture(m_Path & pfn)
        If Not Col_Contains(m_Flags, key) Then
            m_Flags.Add spic, key
        End If
    Next
End Function

Private Function ReadDataFile() As String()
Try: On Error GoTo Catch
    m_dPFN = "_data.txt"
    Dim FNm As String:  FNm = m_Path & m_dPFN
    Dim FNr As Integer: FNr = FreeFile
    Open FNm For Binary Access Read As FNr
    Dim sCont As String: sCont = Space$(LOF(FNr))
    Get FNr, , sCont
    ReadDataFile = Split(sCont, vbCrLf)
    Exit Function
Catch:
End Function

Public Property Get Flag(key_ISO3 As String) As StdPicture
    If Len(key_ISO3) = 0 Then Exit Property
    If Col_Contains(m_Flags, key_ISO3) Then
        Set Flag = m_Flags(key_ISO3)
    End If
End Property
'
'Public Function Col_Contains(col As Collection, key As String) As Boolean
'    'for this Function all credits go to the incredible www.vb-tec.de alias Jost Schwider
'    'you can find the original version of this function here: https://vb-tec.de/collctns.htm
'    On Error Resume Next
''  '"Extras->Optionen->Allgemein->Unterbrechen bei Fehlern->Bei nicht verarbeiteten Fehlern"
'    If IsEmpty(col(key)) Then: 'DoNothing
'    Col_Contains = (Err.Number = 0)
'    On Error GoTo 0
'End Function
'
