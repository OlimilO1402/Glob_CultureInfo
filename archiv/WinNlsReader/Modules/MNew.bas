Attribute VB_Name = "MNew"
Option Explicit

Public Function PathFileName(ByVal aPathFileName As String, _
                     Optional ByVal aFileName As String, _
                     Optional ByVal aExt As String) As PathFileName
    Set PathFileName = New PathFileName: PathFileName.New_ aPathFileName, aFileName, aExt
End Function

Public Function List(Of_T As EDataType, _
                     Optional ArrColStrTypList, _
                     Optional ByVal IsHashed As Boolean = False, _
                     Optional ByVal Capacity As Long = 32, _
                     Optional ByVal GrowRate As Single = 2, _
                     Optional ByVal GrowSize As Long = 0) As List
    Set List = New List: List.New_ Of_T, ArrColStrTypList, IsHashed, Capacity, GrowRate, GrowSize
End Function

Public Function CConstant(ByVal aName As String, ByVal aValue As String, ByVal aComment As String) As CConstant
    Set CConstant = New CConstant: CConstant.New_ aName, aValue, aComment
End Function

Public Function ConstsWinNlsh(aPFN As PathFileName) As ConstsWinNlsh
    Set ConstsWinNlsh = New ConstsWinNlsh: ConstsWinNlsh.New_ aPFN
End Function
