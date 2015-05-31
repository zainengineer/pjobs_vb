Attribute VB_Name = "modCommon"
Option Explicit
Public mconnSettingsDB As Connection
Private mconnContentDB As Connection
Private mdicSettings As Dictionary
Private mstrHTMLOrignalSource As String
Private mstrHTMLProcessedSource As String
Private mdicProcessedSettings As Dictionary
Private mlngAppID As Long
Private mdicAppFields As Dictionary
Private mrsContent As Recordset
Private mrsAdds As Recordset
Private mstrProducedHTML As String
Private mstrTemplate As String
Private mlngAddCounter As Long
Private mlngAddInterval As Long
Private mstrFlashSource As String
Private mstrImageSource As String
Private mstrAddType As String
Private mstrAddTemplateWithTags As String
Private mstrAddTemplate As String
Private mlngTotalRecords As Long
Private mlngCurrentRecord As Long
Dim mlngMaxEntries As Long
Dim mstrNavigationTemplate As String
Dim mlngCurrentPageNumber As Long
Dim mstrDestination As String
Private mstrBreakField As String
Private mstrBreak As String
Private mbolLinked As Boolean
Public Sub Init()
Dim strDBPath As String
  
  frmMain.Caption = "Initializing DB ......."
  strDBPath = App.Path & "\..\Files\GeneralHTML.mdb"
  Set mconnSettingsDB = ADOFunct.Access2000Conn(strDBPath, False)
  
End Sub
Private Sub GetAppSettings()
Dim rs As Recordset
Dim fldLoop As Field
Dim strQuery As String

  
  strQuery = "SELECT * FROM tblApp WHERE ID=" & mlngAppID
  Set rs = GetADORecordSet(mconnSettingsDB, strQuery)
  Set mdicSettings = New Dictionary
  For Each fldLoop In rs.Fields
    mdicSettings.Add fldLoop.Name, getFieldValue(fldLoop.Value)
  Next fldLoop
  
 mlngAddInterval = GetSingleSetting("AddRepeatInterval")
 mlngMaxEntries = GetSingleSetting("MaxEntries")
 mstrBreakField = GetSingleSetting("PageSeperatorField")
End Sub
Public Sub ExecuteApp(strAppName As String)
Dim dicBreaks As Dictionary
Dim varBreak As Variant
  
  AppInit strAppName
  
  
  If mstrBreakField = "" Then
    StartMainProcessing
  Else
    Set dicBreaks = GetBreaks
    For Each varBreak In dicBreaks.Items
      mstrBreak = getFieldValue((varBreak))
      If mstrBreak = "" Then
        mstrBreak = "Null"
      End If
      ExecuteBreak
    Next varBreak
  End If
  
  RunAfterApp
End Sub
Private Sub ReadOrignalHTML()
Dim strHtmlSource As String
Dim fs As FileSystemObject
Dim stmHTML As TextStream
Dim bolRelative As Boolean
  
  frmMain.Caption = "Reading HTML Source ......"
  Set fs = New FileSystemObject
  
  strHtmlSource = GetSingleSetting("HTMLSource")
  bolRelative = GetSingleSetting("HTMLSourceRelative")
  If bolRelative Then
    strHtmlSource = App.Path & "\" & strHtmlSource
  End If
  Set stmHTML = fs.OpenTextFile(strHtmlSource)
  mstrHTMLOrignalSource = stmHTML.ReadAll
  
End Sub
Private Function GetSingleSetting(strField As String) As Variant
Dim varResult As Variant
   
  varResult = modGeneral.SecurelyGetDicValue(mdicSettings, strField)
  
GetSingleSetting = varResult
End Function

Private Sub AppInit(strAppName As String)
Dim strCritaria As String

  mstrBreakField = ""
  mstrBreak = ""
  mbolLinked = False
  
  
  
  
  strCritaria = ADOFunct.KnownTypeSQLAssign("Name", strAppName)
  mlngAppID = ADOFunct.GetDefault(mconnSettingsDB, "ID", "tblApp", strCritaria)
  
  frmMain.Caption = "Getting '" & strAppName & "' Settings ....."
  GetAppSettings
  StartRegistry
  RunBeforeApp
  GetAppFields
  ReadOrignalHTML
  ProcessHTMLSource
  GetProcessSettings
  PutFlashSource
  PutImageSource
  ConnectContentDB
  GetAddRecordSet
  
  
End Sub
Private Sub GetProcessSettings()
  
  frmMain.Caption = "Geting Processed Settings ....."
  Set mdicProcessedSettings = New Dictionary
  
  AddSingleProcessSetting "Template", "TemplateStartTag", "TemplateEndTag"
  AddSingleProcessSetting "Add", "AddStartTag", "AddEndTag"
  AddSingleProcessSetting "HTMLStart", "", "TemplateStartTag", False
  AddSingleProcessSetting "FlashTemplate", "FlashTemplateStartTag", "FlashTemplateEndTag"
  AddSingleProcessSetting "ImageTemplate", "ImageTemplateStartTag", "ImageTemplateEndTag"
  If mlngMaxEntries > 0 Then
    AddSingleProcessSetting "HTMLEnd", "NavigationEndTag", "", False
  Else
    AddSingleProcessSetting "HTMLEnd", "TemplateEndTag", "", False
  End If
  
  AddSingleProcessSetting "AddType", "AddTypeStart", "AddTypeEnd"
  
  AddSingleProcessSetting "Navigation", "NavigationStartTag", "NavigationEndTag"
  AddSingleProcessSetting "PreviousEnable", "PreviousEnableStart", "PreviousEnableEnd"
  AddSingleProcessSetting "PreviousDisable", "PreviousDisableStart", "PreviousDisableEnd"
  AddSingleProcessSetting "NextEnable", "NextEnableStart", "NextEnableEnd"
  AddSingleProcessSetting "NextDisable", "NextDisableStart", "NextDisableEnd"
  AddSingleProcessSetting "Enumirated", "EnumiratedStart", "EnumiratedEnd"
  AddSingleProcessSetting "EnumiratedNumber", "EnumiratedNumberStart", "EnumiratedNumberEnd"
  AddSingleProcessSetting "EnumiratedCurrent", "EnumiratedCurrentStart", "EnumiratedCurrentEnd"
  
  
  mstrTemplate = GetSingleProcessSetting("Template")
  mstrAddTemplateWithTags = GetSingleSetting("AddStartTag") & GetSingleProcessSetting("Add") & GetSingleSetting("AddEndTag")
  mstrAddTemplate = GetSingleProcessSetting("Add")
  mstrAddType = GetSingleSetting("AddTypeStart") & GetSingleProcessSetting("AddType") & GetSingleSetting("AddTypeEnd")
  mstrNavigationTemplate = GetSingleProcessSetting("Navigation")
  PutDestination
  
End Sub
Sub AddSingleProcessSetting(ByVal strKeyName As String, _
            ByVal strStartSetting As String, _
            ByVal strEndSetting As String, _
            Optional bolEmptyIfNotFound As Boolean = True)
Dim strValue As String
  
    
    If strStartSetting <> "" Then
      strStartSetting = GetSingleSetting(strStartSetting)
    Else
      strStartSetting = ""
    End If
    
    If strEndSetting <> "" Then
      strEndSetting = GetSingleSetting(strEndSetting)
    Else
      strEndSetting = ""
    End If
    If (strStartSetting = "" Or strEndSetting = "") And bolEmptyIfNotFound Then
      strValue = ""
    Else
      strValue = modGeneral.GetBetweenText(mstrHTMLProcessedSource, strStartSetting, strEndSetting)
    End If
    mdicProcessedSettings.Add strKeyName, strValue
End Sub
Sub ProcessHTMLSource()
Dim strHTMLStartTag As String
Dim strHTMLEndTag As String
  strHTMLStartTag = GetSingleSetting("HTMLStartTag")
  strHTMLEndTag = GetSingleSetting("HTMLEndTag")
  mstrHTMLProcessedSource = modGeneral.GetBetweenText(mstrHTMLOrignalSource, strHTMLStartTag, strHTMLEndTag)
End Sub
Sub GetAppFields()
Dim rs As Recordset
Dim strQuery As String
Dim strKey As String
Dim strItem As String
  
  frmMain.Caption = "Getting App Fields ..."
  strQuery = "SELECT Name,FindText FROM tblFieldLookUp WHERE AppID=" & mlngAppID
  Set rs = ADOFunct.GetADORecordSet(mconnSettingsDB, strQuery)
  Set mdicAppFields = New Dictionary
  
  Do While Not (rs.EOF Or rs.BOF)
    strKey = rs.Fields("Name")
    strItem = rs.Fields("FindText")
    mdicAppFields.Add strKey, strItem
    rs.MoveNext
  Loop
  
End Sub

Sub StartMainProcessing()
    
  If Not mbolLinked Then
    mstrProducedHTML = GetSingleProcessSetting("HTMLStart")
  End If
  GetContentRecordSet
  mrsContent.MoveLast
  mrsContent.MoveFirst
  Do While Not (mrsContent.EOF Or mrsContent.BOF)
    DoEvents
    mlngCurrentRecord = mlngCurrentRecord + 1
    AddContent
    mrsContent.MoveNext
    BreakAndWritePage
  Loop
  
  
End Sub
Sub ConnectContentDB()
Dim strDBPath As String
Dim bolRelative As Boolean

  frmMain.Caption = "Connecting Content DB ....."
  strDBPath = GetSingleSetting("DBPath")
  bolRelative = GetSingleSetting("RelativeDBPath")
  If bolRelative Then
    strDBPath = App.Path & "\" & strDBPath
  End If
  
  Set mconnContentDB = ADOFunct.Access2000Conn(strDBPath)
  
End Sub

Sub GetContentRecordSet()
Dim strQuery As String
Dim strFields As String
Dim strTable As String
Dim strOrderField As String
Dim bolAscending As Boolean
Dim strOrderFlow As String
Dim strCritaria As String
  

  If mstrBreak = "" Then
    frmMain.Caption = "Getting Content Values ......"
  Else
    If mstrBreak = "Null" Then
      strCritaria = KnownTypeSQLAssign(mstrBreakField, Null)
    Else
      strCritaria = KnownTypeSQLAssign(mstrBreakField, mstrBreak)
    End If
    strCritaria = PutWhere(strCritaria)
  End If
  
  strOrderField = GetSingleSetting("OrderBy")
  bolAscending = GetSingleSetting("Ascending")
  strFields = Join(mdicAppFields.Keys, ",")
  If Not mdicAppFields.Exists(strOrderField) Then
    strFields = strOrderField & "," & strFields
  End If
  If bolAscending Then
    strOrderFlow = "asc"
  Else
    strOrderFlow = "desc"
  End If
  strTable = GetSingleSetting("ContentTable")
  strQuery = "SELECT " & strFields & " FROM " & strTable & strCritaria & " ORDER BY " & strOrderField & " " & strOrderFlow
  Set mrsContent = ADOFunct.GetADORecordSet(mconnContentDB, strQuery)
  mrsContent.MoveLast
  mlngTotalRecords = mrsContent.RecordCount
  mlngCurrentRecord = 0
End Sub
Private Function GetSingleProcessSetting(strField As String) As Variant
Dim varResult As Variant
   
  varResult = modGeneral.SecurelyGetDicValue(mdicProcessedSettings, strField)
  
GetSingleProcessSetting = varResult
End Function
Private Sub AddContent()
Dim lngCounter As Long
Dim strToFind As String
Dim strField As String
Dim strToReplace As String
Dim strTemplate As String
  
  strTemplate = mstrTemplate
  For lngCounter = 0 To mdicAppFields.Count - 1
  
    strField = mdicAppFields.Keys(lngCounter)
    strToFind = mdicAppFields.Items(lngCounter)
    strToReplace = getFieldValue(mrsContent.Fields(strField).Value)
    strToReplace = Replace(strToReplace, vbCrLf, "<BR>")
    strToReplace = Replace(strToReplace, "  ", "&nbsp;&nbsp;")
    strTemplate = Replace(strTemplate, strToFind, strToReplace)
    
  Next lngCounter
  AddAdd strTemplate
  mstrProducedHTML = mstrProducedHTML & strTemplate
  
End Sub
Private Sub AddAdd(ByRef strTemplate As String)
Dim strAdd As String
  
  
  strAdd = GetAdd
  strTemplate = Replace(strTemplate, mstrAddTemplateWithTags, strAdd)
  
End Sub
Sub WriteAndRunContent()
Dim textFile As TextStream
Dim strDestination As String
Dim lngNextApp As Long
Dim strNextApp As String
  
  
  If mstrBreak = "" Then
    frmMain.Caption = "Writing Start ......."
  End If
  PutBreakHeading
  lngNextApp = GetSingleSetting("NextLink")
  If lngNextApp > 0 Then
    strNextApp = GetAppName(lngNextApp)
    mstrProducedHTML = mstrProducedHTML & GetSingleProcessSetting("HTMLEnd")
    RemoveNavigation
    mbolLinked = True
    ExecuteApp strNextApp
    RemoveNavigation
    Exit Sub
  End If
  AddNavigation
  mstrProducedHTML = mstrProducedHTML & GetSingleProcessSetting("HTMLEnd")
  RemoveNavigation
  Dim fs As FileSystemObject
  
  
  Set fs = New FileSystemObject
  strDestination = GetFileLink(mlngCurrentPageNumber)
  modGeneral.ConfirmFolder fs.GetParentFolderName(strDestination)
  PutSelfPath strDestination
  Set textFile = fs.CreateTextFile(strDestination)
  textFile.Write mstrProducedHTML
  textFile.Close
  Set fs = Nothing
  If mlngCurrentPageNumber = 1 Then
    frmMain.Caption = "Opening File ......."
    modGeneral.StartFile strDestination
  End If
End Sub
Sub GetAddRecordSet()
Dim strQuery As String
Dim strTable As String

  frmMain.Caption = "Getting Adds Values ......"
  
  strTable = GetSingleSetting("AddsTable")
  strQuery = "SELECT * FROM " & strTable & " ORDER BY ID"
  Set mrsAdds = ADOFunct.GetADORecordSet(mconnContentDB, strQuery)

End Sub
Function GetAdd() As String
Dim strResult As String
Dim strSource As String
Dim strText As String
Dim strHREF As String
Dim strAddType As String
Dim strTarget As String
Dim strAddSuffex As String
  
  mlngAddCounter = mlngAddCounter + 1
  If (mlngAddCounter Mod mlngAddInterval) <> 0 Then
    Exit Function
  End If
  If (mrsAdds.EOF Or mrsAdds.BOF) Then
    mrsAdds.MoveFirst
  End If
  
  strSource = getFieldValue(mrsAdds.Fields("Source"))
  strText = getFieldValue(mrsAdds.Fields("Text"))
  strHREF = getFieldValue(mrsAdds.Fields("HREF"))
  strTarget = getFieldValue(mrsAdds.Fields("Target"))
  strAddSuffex = GetSingleSetting("AddSuffex")
  
  If strSource <> "" Then
    strSource = strAddSuffex & strSource
  End If
  strSource = FixSource(strSource)
  If strResult = "" Then
    strResult = getFlashAdd(strSource)
  End If
  
  If strResult = "" Then
    strResult = getImageAdd(strSource, strHREF, strTarget)
  End If
  
  If strResult = "" Then
    strResult = strText
  End If
    
  
  strResult = Replace(mstrAddTemplate, mstrAddType, strResult)
    
  mrsAdds.MoveNext
    
    
  
GetAdd = strResult
End Function
Sub PutFlashSource()
Dim strBefore As String
Dim strAfter As String
Dim strFlashTemplate As String
    
      

  strFlashTemplate = GetSingleProcessSetting("FlashTemplate")
  If strFlashTemplate <> "" Then
    strBefore = "<param name=""movie"" value="""
    strAfter = """>"
    mstrFlashSource = modGeneral.GetBetweenText(strFlashTemplate, strBefore, strAfter)
  End If
  

End Sub
Sub PutImageSource()
Dim strBefore As String
Dim strAfter As String
Dim strImageTemplate As String
    
      

  strImageTemplate = GetSingleProcessSetting("ImageTemplate")
  
  If InStr(strImageTemplate, "src=""") <> 0 Then
    strBefore = "src="""
    strAfter = """"
    mstrImageSource = modGeneral.GetBetweenText(strImageTemplate, strBefore, strAfter)
  End If

End Sub
Function getFlashAdd(strSource As String) As String
Dim strResult As String
Dim strFlashTemplate As String
Dim fs As FileSystemObject
Dim strExtention As String
  
  Set fs = New FileSystemObject
  strExtention = fs.GetExtensionName(strSource)
  If LCase(strExtention) = "swf" Then
    strFlashTemplate = GetSingleProcessSetting("FlashTemplate")
    strResult = Replace(strFlashTemplate, mstrFlashSource, strSource)
  Else
    strResult = ""
  End If
getFlashAdd = strResult
End Function
Function getImageAdd(strSource As String, _
          strHREF As String, _
          ByVal strTarget As String) As String
Dim strResult As String
Dim strImageTemplate As String

  If strSource = "" Then
    Exit Function
  End If
  strImageTemplate = GetSingleProcessSetting("ImageTemplate")
  strResult = Replace(strImageTemplate, mstrImageSource, strSource)
  If Trim(strHREF) <> "" Then
    If Trim(strTarget) <> "" Then
      strTarget = " Target=""" & strTarget & """"
    End If
    strResult = "<a href=""" & strHREF & """" & strTarget & ">" & strResult & "</a>"
    
  End If
  
getImageAdd = strResult
End Function
Sub AddNavigation()

  If mlngMaxEntries = 0 Then
    mlngCurrentPageNumber = 1
    RemoveNavigation
    Exit Sub
  End If
  If mlngTotalRecords <= mlngMaxEntries Then Exit Sub
  mlngCurrentPageNumber = mlngCurrentRecord \ mlngMaxEntries
  If (mlngCurrentRecord Mod mlngMaxEntries) <> 0 Then
    mlngCurrentPageNumber = mlngCurrentPageNumber + 1
  End If
  
  AddPreviousLink
  AddEnumirated
  AddNextLink
    
  mstrProducedHTML = mstrProducedHTML & mstrNavigationTemplate
    
End Sub
Sub BreakAndWritePage()
Dim lngMaxEntries As Long

  If mlngCurrentRecord = mlngTotalRecords Then
    WriteAndRunContent
  ElseIf mlngMaxEntries > 0 Then
    If (mlngCurrentRecord Mod mlngMaxEntries) = 0 Then
      WriteAndRunContent
      mstrProducedHTML = GetSingleProcessSetting("HTMLStart")
      mstrNavigationTemplate = GetSingleProcessSetting("Navigation")
    End If
  End If

End Sub
Sub AddPreviousLink()
Dim strPreviousEnable As String
Dim strPreviousDisable As String
Dim strPreviousEnableWithTags As String
Dim strPreviousDisableWithTags As String
Dim strPreviousHREF As String
Dim strFile As String
Dim fs As FileSystemObject
  
  Set fs = New FileSystemObject
  strPreviousEnable = GetSingleProcessSetting("PreviousEnable")
  strPreviousEnableWithTags = GetSingleSetting("PreviousEnableStart") & strPreviousEnable & GetSingleSetting("PreviousEnableEnd")
  strPreviousDisable = GetSingleProcessSetting("PreviousDisable")
  strPreviousDisableWithTags = GetSingleSetting("PreviousDisableStart") & strPreviousDisable & GetSingleSetting("PreviousDisableEnd")
  strPreviousHREF = GetSingleSetting("PreviousHREF")
  
  If mlngCurrentRecord <= mlngMaxEntries Then 'Disable Previous
    mstrNavigationTemplate = Replace(mstrNavigationTemplate, strPreviousEnableWithTags, "")
    mstrNavigationTemplate = Replace(mstrNavigationTemplate, strPreviousDisableWithTags, strPreviousDisable)
  Else 'Enable
    mstrNavigationTemplate = Replace(mstrNavigationTemplate, strPreviousDisableWithTags, "")
    strFile = GetFileLink(mlngCurrentPageNumber - 1)
    strFile = fs.GetFileName(strFile)
    strPreviousEnable = Replace(strPreviousEnable, strPreviousHREF, strFile)
    mstrNavigationTemplate = Replace(mstrNavigationTemplate, strPreviousEnableWithTags, strPreviousEnable)
  End If
  
End Sub
Sub AddEnumirated()
Dim strEnumiratedWithTags As String
Dim strEnumirated As String
Dim strEnumiratedNumberWithTags As String
Dim strEnumiratedNumber As String
Dim strEnumiratedTags As String
Dim strEnumiratedNumberCaption As String
Dim lngCounter As Long
Dim strEnumiratedList As String
Dim lngTotalPages As Long
Dim strLoopEnum As String
Dim strEnumiratedFile As String
Dim strSeperator As String
Dim strEnumiratedCurrent As String
Dim strEnumiratedCurrentWithTags As String
Dim strEnumiratedHREF As String
Dim fs As FileSystemObject

  Set fs = New FileSystemObject
  
  strEnumirated = GetSingleProcessSetting("Enumirated")
  strEnumiratedWithTags = GetSingleSetting("EnumiratedStart") & strEnumirated & GetSingleSetting("EnumiratedEnd")
  strEnumiratedNumber = GetSingleProcessSetting("EnumiratedNumber")
  strEnumiratedNumberWithTags = GetSingleSetting("EnumiratedNumberStart") & strEnumiratedNumber & GetSingleSetting("EnumiratedNumberEnd")
  strEnumiratedNumberCaption = GetSingleSetting("EnumiratedNumberCaption")
  strEnumiratedFile = GetSingleSetting("EnumiratedFile")
  strEnumiratedCurrent = GetSingleProcessSetting("EnumiratedCurrent")
  strEnumiratedCurrentWithTags = GetSingleSetting("EnumiratedCurrentStart") & strEnumiratedCurrent & GetSingleSetting("EnumiratedCurrentEnd")
  
  lngTotalPages = mlngTotalRecords \ mlngMaxEntries
  If (mlngTotalRecords Mod mlngMaxEntries) <> 0 Then
    lngTotalPages = lngTotalPages + 1
  End If
  
  For lngCounter = 1 To lngTotalPages
    If lngCounter <> mlngCurrentPageNumber Then
      strLoopEnum = Replace(strEnumiratedNumber, strEnumiratedNumberCaption, lngCounter)
      strEnumiratedHREF = GetFileLink(lngCounter)
      strEnumiratedHREF = fs.GetFileName(strEnumiratedHREF)
      strLoopEnum = Replace(strLoopEnum, strEnumiratedFile, strEnumiratedHREF)
      strEnumiratedList = strEnumiratedList & strSeperator & strLoopEnum
      strSeperator = GetSingleSetting("EnumiratedSeperator")
    Else
      strLoopEnum = Replace(strEnumiratedCurrent, strEnumiratedNumberCaption, lngCounter)
      strEnumiratedHREF = GetFileLink(lngCounter)
      strEnumiratedHREF = fs.GetFileName(strEnumiratedHREF)
      strLoopEnum = Replace(strLoopEnum, strEnumiratedFile, strEnumiratedHREF)
      strEnumiratedList = strEnumiratedList & strSeperator & strLoopEnum
      strSeperator = GetSingleSetting("EnumiratedSeperator")
    End If
  Next lngCounter
  
  
  strEnumirated = Replace(strEnumirated, strEnumiratedNumberWithTags, strEnumiratedList)
  mstrNavigationTemplate = Replace(mstrNavigationTemplate, strEnumiratedWithTags, strEnumirated)
  mstrNavigationTemplate = Replace(mstrNavigationTemplate, strEnumiratedCurrentWithTags, "")
  
End Sub
Sub AddNextLink()
Dim strNextEnable As String
Dim strNextDisable As String
Dim strNextEnableWithTags As String
Dim strNextDisableWithTags As String
Dim strNextHREF As String
Dim strFile As String
Dim fs As FileSystemObject
  
  Set fs = New FileSystemObject
  strNextEnable = GetSingleProcessSetting("NextEnable")
  strNextEnableWithTags = GetSingleSetting("NextEnableStart") & strNextEnable & GetSingleSetting("NextEnableEnd")
  strNextDisable = GetSingleProcessSetting("NextDisable")
  strNextDisableWithTags = GetSingleSetting("NextDisableStart") & strNextDisable & GetSingleSetting("NextDisableEnd")
  strNextHREF = GetSingleSetting("NextHREF")
  
  If mlngCurrentRecord = mlngTotalRecords Then  'Disable
    mstrNavigationTemplate = Replace(mstrNavigationTemplate, strNextEnableWithTags, "")
    mstrNavigationTemplate = Replace(mstrNavigationTemplate, strNextDisableWithTags, strNextDisable)
  Else 'Enable
    mstrNavigationTemplate = Replace(mstrNavigationTemplate, strNextDisableWithTags, "")
    strFile = GetFileLink(mlngCurrentPageNumber + 1)
    strFile = fs.GetFileName(strFile)
    strNextEnable = Replace(strNextEnable, strNextHREF, strFile)
    mstrNavigationTemplate = Replace(mstrNavigationTemplate, strNextEnableWithTags, strNextEnable)
  End If

End Sub

Sub PutDestination()
Dim bolRelative As Boolean

  mstrDestination = GetSingleSetting("ProducePath")
  bolRelative = GetSingleSetting("ProducePathRelative")
  
  If bolRelative Then
    mstrDestination = App.Path & "\" & mstrDestination
  End If
End Sub
Function GetFileLink(ByVal lngNumber As Long) As String
Dim strResult As String
Dim strExtention As String
Dim fs As FileSystemObject
Dim strBaseName As String
Dim strOrignalName As String

  If lngNumber = 0 Then
    lngNumber = 1
    mlngCurrentPageNumber = 1
  End If
  If lngNumber <> 1 Then
    Set fs = New FileSystemObject
    strExtention = fs.GetExtensionName(mstrDestination)
    strBaseName = fs.GetBaseName(mstrDestination)
    strResult = strBaseName & lngNumber & "." & strExtention
    strOrignalName = strBaseName & "." & strExtention
    strResult = Replace(mstrDestination, strOrignalName, strResult)
  Else
    strResult = mstrDestination
  End If
  strResult = CreateBreakDestination(strResult)
'  If lngNumber = 0 Then
'    RemoveOldFiles strResult
'  End If
GetFileLink = strResult
End Function
Function GetBreaks() As Dictionary
Dim dicResult As Dictionary
Dim strTable As String
Dim strBreakField As String
   strTable = GetSingleSetting("ContentTable")
   strBreakField = GetSingleSetting("PageSeperatorField")
   Set dicResult = ADOFunct.getFieldValueList(mconnContentDB, strTable, strBreakField, 0)
Set GetBreaks = dicResult
End Function
Sub ExecuteBreak()
  
  mlngCurrentPageNumber = 1
  frmMain.Caption = "Executing " & mstrBreak & " ......."
  mstrProducedHTML = GetSingleProcessSetting("HTMLStart")
  GetContentRecordSet
  mrsContent.MoveLast
  mrsContent.MoveFirst
  Do While Not (mrsContent.EOF Or mrsContent.BOF)
    DoEvents
    mlngCurrentRecord = mlngCurrentRecord + 1
    AddContent
    mrsContent.MoveNext
    BreakAndWritePage
  Loop
  
  
End Sub

Function CreateBreakDestination(strDestination As String)
Dim strResult As String
Dim strParent As String
Dim strExtention As String
Dim fs As FileSystemObject
Dim strBaseName As String
Dim strOrignalName As String
  
  If mstrBreak <> "" Then
    Set fs = New FileSystemObject
    strExtention = fs.GetExtensionName(mstrDestination)
    strBaseName = fs.GetBaseName(mstrDestination)
    strOrignalName = strBaseName & "." & strExtention
    strResult = Replace(mstrDestination, strOrignalName, "")
    strResult = strResult & LCase(mstrBreak)
    modGeneral.ConfirmFolder strResult
    
    strExtention = fs.GetExtensionName(strDestination)
    strBaseName = fs.GetBaseName(strDestination)
    strOrignalName = strBaseName & "." & strExtention
    strResult = strResult & "\" & strOrignalName
  Else
    strResult = strDestination
  End If
  
CreateBreakDestination = strResult
End Function
Function FixSource(strSource As String) As String
Dim strResult As String

  
  
  If mstrBreak <> "" Then
    If strSource <> "" Then
     strResult = "../" & strSource
    End If
  Else
    strResult = strSource
  End If
  
FixSource = strResult
End Function
Function GetAppName(lngID As Long) As String
Dim strResult As String
Dim strCritaria As String
  
  strCritaria = "ID=" & lngID
  strResult = ADOFunct.GetDefault(mconnSettingsDB, "Name", "tblApp", strCritaria)
  
GetAppName = strResult
End Function
Sub RemoveNavigation()
Dim strNavigationStartTag As String
Dim strNavigationEndTag As String
Dim strNavigationTemplate As String
  
  strNavigationStartTag = GetSingleSetting("NavigationStartTag")
  strNavigationEndTag = GetSingleSetting("NavigationEndTag")
  strNavigationTemplate = strNavigationStartTag & mstrNavigationTemplate & strNavigationEndTag
  mstrProducedHTML = Replace(mstrProducedHTML, strNavigationTemplate, "")
  
End Sub
Sub PutBreakHeading()
Dim strHeading As String
Dim strHeadingTag As String

  If mstrBreak = "" Then Exit Sub
  If mstrBreak = "Null" Then
    strHeading = GetSingleSetting("PageSeperatorNullName")
  Else
    strHeading = mstrBreak
  End If
  strHeadingTag = GetSingleSetting("PageSeperatorHeading")
  mstrProducedHTML = Replace(mstrProducedHTML, strHeadingTag, strHeading)
  
End Sub
Sub RemoveOldFiles(strSingleFile)
Dim fs As FileSystemObject
Dim flParent As Folder
Dim fsCurrent As File
Dim filList As File
Dim strItsName As String
Dim strExtention As String
Dim strGivenName As String
Dim strLoopExtention As String
Dim strLoopName As String
Dim strRemaining As String
Dim strToDelete As String

  
  Set fs = New FileSystemObject
  Set fsCurrent = fs.GetFile(strSingleFile)
  Set flParent = fsCurrent.ParentFolder
  strExtention = fs.GetExtensionName(strSingleFile)
  strGivenName = fs.GetBaseName(strSingleFile)
  
  
  For Each filList In flParent.Files
      strLoopExtention = fs.GetExtensionName(filList.Name)
      strLoopName = fs.GetBaseName(filList.Name)
      If InStr(strLoopName, strGivenName) = 1 Then
        strRemaining = Replace(strLoopName, strGivenName, "")
        If IsNumeric(strRemaining) Then
          strToDelete = strLoopName & "." & strLoopExtention
          strToDelete = Replace(strSingleFile, strGivenName, strLoopName)
          Debug.Print "Deleting ...... " & strToDelete
          fs.DeleteFile strToDelete
        Else
          Debug.Print "         " & strSingleFile
        End If
      Else
        Debug.Print "         " & strSingleFile
      End If
    
  Next filList
  
End Sub

Sub RunBeforeApp()
Dim strAppBefore As String
Dim bolExist As Boolean
Dim fs As FileSystemObject
Dim bolAppRunning As Boolean
Dim bolRelative As Boolean
  
  frmMain.Caption = "Running Before App ....."
  strAppBefore = GetSingleSetting("AppBefore")
  If Trim(strAppBefore) <> "" Then
    Set fs = New FileSystemObject
    bolRelative = IsRelativePath(strAppBefore)
    If bolRelative Then
      strAppBefore = App.Path & "\" & strAppBefore
    End If
    
    bolExist = fs.FileExists(strAppBefore)
    If bolExist Then
      MRegistry.PutRegistryValue "PreAppRunning", True
      Shell strAppBefore
      
      Do
        bolAppRunning = (MRegistry.GetRegistryValue("PreAppRunning") = True)
        DoEvents
      Loop While bolAppRunning
      
    Else
       MsgBox "'" & strAppBefore & "' Doesn't Exist", vbCritical
       Debug.Assert False
    End If
  End If

  
  
  
End Sub
Sub RunAfterApp()
Dim strAppAfter As String
Dim bolExist As Boolean
Dim fs As FileSystemObject
  
  strAppAfter = GetSingleSetting("AppAfter")
  If Trim(strAppAfter) <> "" Then
    Set fs = New FileSystemObject
    bolExist = fs.FileExists(strAppAfter)
    If bolExist Then
      Shell strAppAfter
    Else
       MsgBox "'" & strAppAfter & "' Doesn't Exist", vbCritical
    End If
  End If
  
End Sub
Sub StartRegistry()
Dim dicSettings As Dictionary
Dim strApp As String
Dim strSection As String
Dim strDBJobs As String
Dim bolRelative As Boolean
Dim strProducePath As String
Dim strTemplate As String

  strApp = "HTMLGEN"
  strSection = "LinkedExes"
  Set dicSettings = New Dictionary
  
  dicSettings.Add "PreAppRunning", "Unknown"
  dicSettings.Add "DBJobs", "Unknown"
  dicSettings.Add "ProducedPath", "Unknown"
  dicSettings.Add "Template", "Unknown"
  MRegistry.Initialize strApp, strSection, dicSettings
  
  strDBJobs = GetSingleSetting("DBPath")
  strTemplate = GetSingleSetting("HTMLSource")
  strProducePath = GetSingleSetting("ProducePath")
  bolRelative = modGeneral.IsRelativePath(strDBJobs)
  If bolRelative Then
    strDBJobs = App.Path & "\" & strDBJobs
  End If
  bolRelative = modGeneral.IsRelativePath(strTemplate)
  If bolRelative Then
    strTemplate = App.Path & "\" & strTemplate
  End If
  bolRelative = modGeneral.IsRelativePath(strProducePath)
  
  If bolRelative Then
    strProducePath = App.Path & "\" & strProducePath
  End If
  
  MRegistry.PutRegistryValue "Template", strTemplate
  MRegistry.PutRegistryValue "DBJobs", strDBJobs
  MRegistry.PutRegistryValue "ProducedPath", strProducePath
  
End Sub
Sub PutSelfPath(ByVal strDestination As String)
Dim strSelfTag As String
Dim fs As FileSystemObject
  
  
  strSelfTag = GetSingleSetting("SelfTag")
  Set fs = New FileSystemObject
  strDestination = fs.GetFileName(strDestination)
  mstrProducedHTML = Replace(mstrProducedHTML, strSelfTag, strDestination)
  
End Sub
