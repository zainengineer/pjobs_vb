Attribute VB_Name = "MRegistry"
Option Explicit
Dim dicSettings As Dictionary
Dim strAppName As String
Dim strSection As String
Public Sub Initialize(strGivenApp As String, _
                strGivenSection As String, _
                dicDefualtValues As Dictionary)
  Set dicSettings = New Dictionary
  
  strAppName = strGivenApp
  strSection = strGivenSection
  Set dicSettings = dicDefualtValues
  
End Sub
Private Function GetRegistryString(ByVal vsItem As String, ByVal vsDefault As String) As String
  'It gets it from HKEY_CURRENT_USER\Software\VB and VBA Program Settings
  GetRegistryString = GetSetting(strAppName, strSection, vsItem, vsDefault)
End Function
Public Function GetRegistryValue(ByVal strKey As String) As String
Dim strResult As String
  If dicSettings.Exists(strKey) Then
    strResult = GetRegistryString(strKey, dicSettings.Item(strKey))
  Else
    Err.Raise 1000, "MRegistry/GetRegistryValue", "No Default Value Found for the " & strKey
  End If
GetRegistryValue = strResult
End Function
Public Sub PutRegistryValue(strKey As String, strValue As String)
'It puts it into HKEY_CURRENT_USER\Software\VB and VBA Program Settings
  If dicSettings.Exists(strKey) Then
    SaveSetting strAppName, strSection, strKey, strValue
  Else
    Err.Raise 1000, "MRegistry/PutRegistryValue", "No Default Value Found for the " & strKey
  End If
  
End Sub

