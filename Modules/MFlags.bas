Attribute VB_Name = "MFlags"
Option Explicit
Private m_Path  As String
Private m_dPFN  As String
Private m_data  As Collection
Private m_Flags As Collection

Public Function ReadFlagPics() As Boolean
    m_Path = App.Path & "\Resources\Flagsbmp\"
    Set m_data = New Collection 'contains the filenames
    Set m_Flags = New Collection
    Dim PFN As String, PFNs() As String: PFNs = ReadDataFile
    Dim sa() As String, Key As String
    Dim i As Long
    For i = 0 To UBound(PFNs)
        PFN = PFNs(i)
        If Len(PFN) Then
            sa = Split(PFN, ", ")
            Key = sa(0)
            If Col_Contains(m_data, Key) Then
                'MsgBox "bereits enthalten: " & vbCrLf & m_data(key) & vbCrLf & pfn
            Else
                m_data.Add PFN, Key
            End If
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
    Else
        Dim spic As StdPicture
        Dim PFN As String
        If Col_Contains(m_data, key_ISO3) Then
            PFN = m_data(key_ISO3)
            Set spic = LoadPicture(m_Path & PFN)
            If Not Col_Contains(m_Flags, key_ISO3) Then
                m_Flags.Add spic, key_ISO3
                Set Flag = spic
            End If
        End If
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
