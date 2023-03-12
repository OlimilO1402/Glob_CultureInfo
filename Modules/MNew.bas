Attribute VB_Name = "MNew"
Option Explicit

Public Function CultureInfo(ByVal aLCID As Long) As CultureInfo
    Set CultureInfo = New CultureInfo: CultureInfo.New_ aLCID
End Function

Public Function CultureInfoN(ByVal aLcidName As String) As CultureInfo
    Set CultureInfoN = New CultureInfo: CultureInfoN.New_ aLcidName
End Function


