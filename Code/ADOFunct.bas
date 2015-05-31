Attribute VB_Name = "ADOFunct"
Option Explicit


Function Access2000Conn(strDBPath As String, _
      Optional bolRelative As Boolean = False, _
      Optional bolExitIfDoesntExist As Boolean = True) As ADODB.Connection
Dim connResult As ADODB.Connection
Dim strConnString As String
Dim fs As FileSystemObject
Dim bolExists As Boolean
  
  Set connResult = New ADODB.Connection
  If bolRelative Then
    strDBPath = App.Path & "\" & strDBPath
  End If
  If bolExitIfDoesntExist Then
    Set fs = New FileSystemObject
     bolExists = fs.FileExists(strDBPath)
     If Not bolExists Then
      MsgBox "Path '" & strDBPath & "' Doesn't Exist so Existing Program", vbCritical
      End
     End If
  End If
  
  strConnString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
  "Data Source=" & strDBPath & ";Persist Security Info=False"
  
  connResult.Open strConnString
Set Access2000Conn = connResult
End Function
Public Sub FillCombo(conn As Connection, _
        strTable As String, _
        strField As String, _
        cmbGiven As ComboBox, _
        Optional ByVal strCondition As String, _
        Optional bolUnique As Boolean = False)
Dim rs As ADODB.Recordset
Dim lngCounter As Long
Dim strQuery As String
Dim strUnique As String
    
    If strCondition <> "" Then
        strCondition = " WHERE " & strCondition
    End If
    
    If bolUnique Then
        strUnique = " DISTINCT "
    End If
    
    strQuery = "Select " & strUnique & strField & _
            " From " & strTable & strCondition
    
    'Set rs = db.OpenRecordset(strQuery)
    Set rs = New ADODB.Recordset
    rs.Open strQuery, conn, adOpenStatic
    cmbGiven.Clear
    If cmbGiven.Style <> 2 Then
        cmbGiven.Text = ""
    End If
    If rs.EOF Or rs.BOF Then Exit Sub
    rs.MoveLast
    rs.MoveFirst
    
    
    For lngCounter = 0 To rs.RecordCount - 1
        If Not IsNull(rs.Fields(0)) Then
            cmbGiven.AddItem rs.Fields(0)
        End If
        rs.MoveNext
    Next lngCounter
    
End Sub

Public Function GetDefault(connGiven As Connection, _
      strField As String, _
        strTable As String, _
        strExtraCritaria As String) As Variant
        
Dim strWhereCritaria As String
Dim strQuery As String
Dim rs As Recordset
Dim varResult As Variant
On Error GoTo IgnoreError

  If strExtraCritaria <> "" Then
     strWhereCritaria = " WHERE " & strExtraCritaria
  End If
  strQuery = "SELECT TOP 1 " & strField & " FROM " & strTable _
            & strWhereCritaria
  Set rs = GetADORecordSet(connGiven, strQuery)
  If (rs.RecordCount = 0) Or IsNull(rs.Fields(strField)) Then
    varResult = Empty
  Else
    rs.MoveFirst
    varResult = CStr(rs.Fields(strField))
  End If
  GetDefault = varResult
  Exit Function
IgnoreError:
  UnExpectedEror Err, "clsDB", "GetDefault"
End Function
Function GetADORecordSet(connGiven As Connection, _
    strQuery As String, _
    Optional curGiven As CursorTypeEnum = adOpenStatic) As Recordset
On Error GoTo QueryError
Dim rsResult As Recordset
   
  Set rsResult = New ADODB.Recordset
  rsResult.Open strQuery, connGiven, curGiven
  Set GetADORecordSet = rsResult
Exit Function
QueryError:
  Clipboard.Clear
  Clipboard.SetText strQuery
  UnExpectedEror Err, "ADOFunct", "GetADORecordSet"
End Function

Public Function SQLFreindly(ByVal varGiven As Variant, _
       Optional typeGiven As DataType = 255, _
       Optional NoNull As Boolean = False, _
       Optional strFirstChar As String, _
       Optional strLastChar As String) As String
Dim strResult As String
Dim strFormatedDate As String
Dim dtTime As Date
  If typeGiven = 255 Then
    typeGiven = GetType(varGiven)
  End If
  If typeGiven = typestring Then
    varGiven = JetSQLFixup(CStr(varGiven))
    strResult = "'" & strFirstChar & varGiven & strLastChar & "'"
  ElseIf typeGiven = TypeDate Then
    strFormatedDate = Format(varGiven, "m-d-yy")
    If CDate(Int(varGiven)) <> varGiven Then 'Date with time
       dtTime = CDate((varGiven - Int(varGiven)))
       strFormatedDate = strFormatedDate & " " & dtTime
    End If
    strResult = "#" & strFormatedDate & "#"
  ElseIf IsNull(varGiven) Then
    Debug.Assert Not NoNull
    strResult = "NULL"
  ElseIf (IsEmpty(varGiven) Or IsMissing(varGiven)) And Not NoNull Then
     strResult = "NULL"
  Else
    strResult = varGiven
  End If
  SQLFreindly = strResult
End Function
Function JetSQLFixup(strTextIn As String)
Dim strResult As String
   
     strResult = Replace(strTextIn, "'", "''")
     strResult = Replace(strResult, "|", "' & chr(124) & '")
     strResult = Replace(strResult, "[", "' & chr(123) & '")
     JetSQLFixup = strResult
End Function


Public Function GetType(varGiven As Variant, _
      Optional TreatAllNumberSame As Boolean = False) As DataType
Dim typeResult As DataType
  Select Case TypeName(varGiven)
  Case "String"
   typeResult = typestring
  Case "Integer"
    typeResult = TypeInteger
  Case "Double"
    typeResult = TypeDecimal
  Case "Boolean"
    typeResult = TypeBoolean
  Case "Date"
    typeResult = TypeDate
  Case "Currency"
    typeResult = TypeCurrency
  End Select
  If TreatAllNumberSame Then
    If (typeResult = TypeCurrency) Or _
       (typeResult = TypeDecimal) Or _
       (typeResult = TypeInteger) Or _
       (typeResult = typelong) Then
       typeResult = TypeInteger
     End If
  End If
  GetType = typeResult
End Function
Function getFieldValue(varValue As Variant) As Variant
Dim varResult As Variant
    
    If IsNull(varValue) Then
        varResult = Empty
    Else
        If IsObject(varValue) Then
            Set varResult = varValue
        Else
            varResult = varValue
        End If
    End If
    
    If IsObject(varResult) Then
      Set getFieldValue = varResult
    Else
      getFieldValue = varResult
    End If
End Function



Public Function KnownTypeUpdateQuery(strTable As String, dicItemValue As Dictionary, strCritaria As String) As String
Dim strResult As String
Dim strItemValue As String
Dim strEffectiveCritaria As String
Dim intCounter As Integer
Dim strField As String
Dim varValue As Variant
Dim EmptyisNull As Boolean
Dim typData As DataType

  'Valiation (Through execeptions instead of just ignoring)
  If (dicItemValue Is Nothing) Then Exit Function
  If (dicItemValue.Count = 0) Then Exit Function
  'If strCritaria = "" Then Exit Function

  
  'For the First use Set
  strField = dicItemValue.Keys(0)
  varValue = dicItemValue.Items(0)
  
  strItemValue = " Set " & SQLAssign(strField, varValue, EmptyisNull, typData)
  
  If dicItemValue.Count > 1 Then 'Don't execute if Only One Value
    For intCounter = 1 To dicItemValue.Count - 1
      strField = dicItemValue.Keys(intCounter)
      varValue = dicItemValue.Items(intCounter)
      EmptyisNull = True
      strItemValue = strItemValue & "," & SQLAssign(strField, varValue, EmptyisNull, typData)
    Next intCounter
  End If
  If strCritaria <> "" Then
    strEffectiveCritaria = " WHERE " & strCritaria
  End If
  strResult = " Update " & strTable & strItemValue & strEffectiveCritaria
  KnownTypeUpdateQuery = strResult
  
End Function

Public Function SQLAssign(strField As String, varValue As Variant, EmptyisNull As Boolean, GivenType As DataType)
Dim strResult As String
Dim strTransformedField As String
Dim varFriendly As Variant
   strTransformedField = FixLongName(strField)
   If Not EmptyisNull Then ValidateValue varValue, GivenType
   varFriendly = SQLFreindly(varValue, , Not EmptyisNull)
   strResult = strTransformedField & " = " & varFriendly
SQLAssign = strResult
End Function

Public Function FixLongName(strGivenLongName) As String
'This Function is used to Fix the Problem of spaces
'Which are allowed in Access if the Name has Spaces
'It puts [] around them i.e Long Name is converted
'into [Long Name] Note [] around the Result
'Note:- I Haven't Checked on the function on actual databases
'Because I don't have Fields with long Names


Dim strTransformedName As String

    If InStr(strGivenLongName, " ") = 0 Then 'Meaning No Spaces
      strTransformedName = strGivenLongName
    Else 'Space Found
       strTransformedName = "[" & strGivenLongName & "]"
    End If

FixLongName = strTransformedName

End Function

Public Function InsertQuery(strTable As String, dicFieldValue As Dictionary) As String
Dim strResult As String
Dim strFields As String
Dim strValues As String
Dim strQueryHead As String
Dim intCounter As Integer
Dim varTempValue As Variant

  'Valiation (Through execeptions instead of just ignoring)
  If (dicFieldValue Is Nothing) Or (dicFieldValue.Count = 0) Then Exit Function
  If strTable = "" Then Exit Function
  strQueryHead = "INSERT INTO " & strTable
  
  strFields = " ( "
  strValues = " ( "
  
'  If dicFieldValue.Count > 0 Then 'Don't execute if Only One Value
'    For intCounter = 1 To dicFieldValue.Count - 1
'      strFields = strFields & "," & dicFieldValue.Keys(intCounter)
'      varTempValue = dicFieldValue.Items(intCounter)
'      varTempValue = SQLFreindly(varTempValue)
'      strValues = strValues & "," & varTempValue
'    Next intCounter
'  End If

  strFields = strFields & Join(dicFieldValue.Keys, " , ")
  strValues = strValues & MyJoin(dicFieldValue.Items)

  strFields = strFields & ")"
  strValues = strValues & ")"
  strResult = strQueryHead & strFields & " Values " & strValues
  InsertQuery = strResult
  
End Function

Public Function MyJoin(varGiven As Variant, Optional strGivenSeperator As String = " , ") As String
'Similar to vb 's join statement but Gives Special Treatment to Null and Strings
'For Dictionary Use dicList.Items like MyJoin(dicList.items)
Dim strResult As String
Dim intCounter As Integer
Dim strSeperator As String
Dim strEffectiveValue As String


  strSeperator = "" 'To avoid First Comma
  
  For intCounter = LBound(varGiven) To UBound(varGiven)
    strEffectiveValue = SQLFreindly(varGiven(intCounter))
    strResult = strResult & strSeperator & strEffectiveValue
    strSeperator = strGivenSeperator
  Next intCounter
  
  MyJoin = strResult
  
End Function

Function GetDataField(ctrGiven As Control) As String
Dim strResult As String
On Error Resume Next
  
    strResult = ctrGiven.DataField
GetDataField = strResult
End Function
Function GetDataFieldDic(frmGiven As Form, _
        bolIgnoreIDField As Boolean, _
        Optional strIDField As String = "ID") As Dictionary
Dim dicResult As Dictionary
Dim ctrLoop As Control
Dim strControlName As String
Dim strDataField As String
  
  Set dicResult = New Dictionary
  For Each ctrLoop In frmGiven.Controls
    strDataField = GetDataField(ctrLoop)
    If strDataField <> "" Then
      If bolIgnoreIDField And strDataField <> strIDField Then
        strControlName = modGeneral.UniqueControlName(ctrLoop)
        dicResult.Add strControlName, strDataField
      End If
    End If
  Next ctrLoop
  
Set GetDataFieldDic = dicResult
End Function
Function GetControlValuesDic(frmGiven As Form, _
        bolIgnoreIDField As Boolean, _
        Optional strIDField As String = "ID") As Dictionary
Dim dicResult As Dictionary
Dim ctrLoop As Control
Dim strControlName As String
Dim strDataField As String
Dim varControlValue As Variant
  
  Set dicResult = New Dictionary
  For Each ctrLoop In frmGiven.Controls
    strDataField = GetDataField(ctrLoop)
    If strDataField <> "" Then
      If bolIgnoreIDField And strDataField <> strIDField Then
        strControlName = modGeneral.UniqueControlName(ctrLoop)
        varControlValue = FormatedControlValue(ctrLoop)
        dicResult.Add strControlName, varControlValue
      End If
    End If
  Next ctrLoop
  
Set GetControlValuesDic = dicResult
End Function
Function GetFieldValueDic(dicControlValues As Dictionary, _
          dicDataFields As Dictionary, _
          bolIgnoreIDField As Boolean, _
          Optional strIDField = "ID") As Dictionary
Dim dicResult As Dictionary
Dim strFieldName As String
Dim varValue As Variant
Dim intCounter As Integer
Dim strControlName As String
  
    Set dicResult = New Dictionary
    For intCounter = 0 To dicDataFields.Count - 1
      strControlName = dicDataFields.Keys(intCounter)
      If dicControlValues.Exists(strControlName) Then
        strFieldName = dicDataFields.Item(strControlName)
        If bolIgnoreIDField And (strFieldName <> strIDField) Then
          varValue = dicControlValues.Item(strControlName)
          dicResult.Add strFieldName, varValue
        End If
      Else
        modGeneral.ExpectedError "Control Name Not Found in dicControlValues", , "GetFieldValueDic"
      End If
    Next intCounter
    
Set GetFieldValueDic = dicResult
End Function
Function GetInsertQuery(frmGiven As Form, _
        strTable As String, _
        bolUseIDField As Boolean, _
        Optional strIDField As String = "ID") As String
Dim strResult As String
Dim dicDataFields As Dictionary
Dim dicValues As Dictionary
Dim dicFieldValues As Dictionary

  Set dicDataFields = GetDataFieldDic(frmGiven, True)
  Set dicValues = GetControlValuesDic(frmGiven, True)
  Set dicFieldValues = GetFieldValueDic(dicValues, dicDataFields, bolUseIDField, strIDField)
  strResult = InsertQuery(strTable, dicFieldValues)
GetInsertQuery = strResult
End Function

Function FormatedControlValue(ctrGiven As Control) As Variant
'Dim varResult As Variant
'Dim varValue As Variant
'Dim fmtControl As IStdDataFormatDisp
'Dim strDataFormat As String
'Dim strControlName As String
'
'    varValue = GetControlValue(ctrGiven)
'    Set fmtControl = ctrGiven.DataFormat
'    strDataFormat = fmtControl.Format
'    If strDataFormat <> "" Then 'Some format is provided
'      If strDataFormat = "0" Then 'Number
'        varResult = Val(varValue)
'      Else
'        strControlName = UniqueControlName(ctrGiven)
'        ExpectedError "Unknown Data format of Control" & strControlName, "Unknown Format", "FormatedControlValue"
'      End If
'    Else 'No format
'      varResult = varValue
'    End If
'
'FormatedControlValue = varResult
End Function

Sub PrintTypes(frmGiven As Form)
Dim dicResult As Dictionary
Dim ctrLoop As Control
Dim strControlName As String
Dim strDataField As String
Dim varControlValue As Variant
Dim varType As Variant
  
  Set dicResult = New Dictionary
  For Each ctrLoop In frmGiven.Controls
    strDataField = GetDataField(ctrLoop)
    If strDataField <> "" Then
      strControlName = modGeneral.UniqueControlName(ctrLoop)
      varControlValue = GetControlValue(ctrLoop)
      varType = ctrLoop.DataFormat.Format
      Debug.Print varControlValue, varType
    End If
  Next ctrLoop
  

End Sub
Function GetSelectQuery(frmGiven As Form, _
        strTable As String, _
        Optional strIDField As String = "ID") As String
Dim strResult As String
Dim strCritaria As String
Dim varIDValue As Variant
  
  varIDValue = GetIDValue(frmGiven, strIDField)
  strCritaria = strIDField & "=" & SQLFreindly(varIDValue)
  strResult = GetSelectQueryByCritaria(frmGiven, strTable, strCritaria, True, strIDField)
  
GetSelectQuery = strResult
End Function
Function GetSelectQueryByCritaria(frmGiven As Form, _
        strTable As String, _
        strCritaria As String, _
        bolIgnoreIDField As Boolean, _
        Optional strIDField As String = "ID") As String
Dim strResult As String
Dim dicDataFields As Dictionary
Dim varIDValue As Variant

  Set dicDataFields = GetDataFieldDic(frmGiven, bolIgnoreIDField, strIDField)
  strResult = SelectQuery(strTable, dicDataFields, strCritaria)

GetSelectQueryByCritaria = strResult
End Function
Function SelectQuery(strTable As String, _
              dicFields As Dictionary, _
              strCritaria As String, _
              Optional bolUseKeys As Boolean = False) As String
Dim strResult As String
Dim strFields As String
Dim intCounter As Integer
Dim varTempValue As Variant

  'Valiation (Through execeptions instead of just ignoring)
  If (dicFields Is Nothing) Or (dicFields.Count = 0) Then Exit Function
  If strTable = "" Then Exit Function
  
  If strCritaria <> "" Then
    strCritaria = " WHERE " & strCritaria
  End If
  If bolUseKeys Then
    strFields = strFields & Join(dicFields.Keys, " , ")
  Else
    strFields = strFields & Join(dicFields.Items, " , ")
  End If
  strResult = "SELECT " & strFields & " FROM  " & strTable & strCritaria
  
SelectQuery = strResult
End Function
Function GetIDValue(frmGiven As Form, _
        Optional strIDField As String = "ID") As Variant
Dim varResult As Variant
Dim ctrLoop As Control
Dim strDataField As String
  
  
  For Each ctrLoop In frmGiven.Controls
    strDataField = GetDataField(ctrLoop)
    If strDataField = strIDField Then
      varResult = FormatedControlValue(ctrLoop)
      Exit For
    End If
  Next ctrLoop
  
GetIDValue = varResult
End Function



Function PrintRecordset(rst As Recordset, _
      Optional bolCopytoClipboard As Boolean = False) As Boolean
      
Dim fld As Field
Dim strClip As String
Dim strLinePrint As String
  On Error GoTo Err_PrintRecordset
  rst.MoveFirst
  With rst
    Do Until .EOF
      For Each fld In .Fields
        strLinePrint = fld.Name & " >> " & fld.Value
        If bolCopytoClipboard Then
          strClip = strClip & strLinePrint & vbCrLf
        Else
          Debug.Print fld.Name & " >> " & fld.Value
        End If
      Next fld
      If bolCopytoClipboard Then
        strClip = strClip & vbCrLf
      Else
        Debug.Print
      End If
      .MoveNext
    Loop
  End With
  PrintRecordset = True
  If bolCopytoClipboard Then
    Clipboard.Clear
    Clipboard.SetText strClip
  End If
Exit_PrintRecordset:
  Exit Function

Err_PrintRecordset:
  MsgBox "Error: " & Err & vbCrLf & Err.Description
  PrintRecordset = False
  Resume Exit_PrintRecordset
End Function

Public Function AndTwo(strFirstString As String, _
      strSecondString As String, _
      Optional UseBraces As Boolean) As String
Dim strResult As String
  If (strFirstString <> "") And (strSecondString <> "") Then
     strResult = strFirstString & " AND " & strSecondString
  ElseIf (strFirstString = "") And (strSecondString = "") Then
     strResult = ""
  ElseIf (strFirstString = "") Then
     strResult = strSecondString
  ElseIf (strSecondString = "") Then
     strResult = strFirstString
  Else
    ExpectedError "This Error was not suppose to Come Contact your software provider to solve it", , "andtwo"
  End If
  If UseBraces And strResult <> "" Then
    strResult = "(" & strResult & ")"
  End If
  AndTwo = strResult
End Function


Public Function PutWhere(strCritaria As String) As String
  If strCritaria = "" Then Exit Function
  PutWhere = "  WHERE " & strCritaria
End Function

Public Function UpdateQuery(strTable As String, dicItemValue As Dictionary, strCritaria As String, arrEmptyisNull() As Boolean, arrType() As DataType) As String
Dim strResult As String
Dim strItemValue As String
Dim strEffectiveCritaria As String
Dim intCounter As Integer
Dim strField As String
Dim varValue As Variant
Dim EmptyisNull As Boolean
Dim typData As DataType

  'Valiation (Through execeptions instead of just ignoring)
  If (dicItemValue Is Nothing) Or (dicItemValue.Count = 0) Then Exit Function
  If strCritaria = "" Then Exit Function
  If strTable = "" Then Exit Function
  
  'For the First use Set
  strField = dicItemValue.Keys(0)
  varValue = dicItemValue.Items(0)
  EmptyisNull = arrEmptyisNull(0)
  typData = arrType(0)
  
  strItemValue = " Set " & SQLAssign(strField, varValue, EmptyisNull, typData)
  
  If dicItemValue.Count > 1 Then 'Don't execute if Only One Value
    For intCounter = 1 To dicItemValue.Count - 1
      strField = dicItemValue.Keys(intCounter)
      varValue = dicItemValue.Items(intCounter)
      EmptyisNull = arrEmptyisNull(intCounter)
      typData = arrType(intCounter)
      strItemValue = strItemValue & "," & SQLAssign(strField, varValue, EmptyisNull, typData)
    Next intCounter
  End If
  
  strEffectiveCritaria = " WHERE " & strCritaria
  strResult = " Update " & strTable & strItemValue & strEffectiveCritaria
  UpdateQuery = strResult
  
End Function

Public Function GetQueryValue(conn As Connection, _
            strQuery As String, _
            Optional strField As String = "", _
            Optional bolConvertNullToEmpty As Boolean = False) As Variant
Dim rs As Recordset
Dim varResult As Variant

 Set rs = GetADORecordSet(conn, strQuery)
 If rs.EOF Or rs.BOF Then
   varResult = Empty
 Else
   If strField = "" Then
     varResult = rs.Fields(0)
   Else
     varResult = rs.Fields(strField)
   End If
   If bolConvertNullToEmpty Then
    If IsNull(varResult) Then
      varResult = Empty
    End If
   End If
 End If
GetQueryValue = varResult
End Function
Public Function KnownTypeSQLAssign(strField As String, varValue As Variant)
Dim strResult As String
Dim strTransformedField As String
Dim varFriendly As Variant
   strTransformedField = FixLongName(strField)
   If IsNull(varValue) Then
     strResult = strTransformedField & " IS Null"
   Else
    varFriendly = SQLFreindly(varValue)
    strResult = strTransformedField & " = " & varFriendly
   End If
KnownTypeSQLAssign = strResult
End Function
Function GetNextID(conn As Connection, _
      strTable As String, _
      Optional strIDField As String = "ID") As Long
Dim lngResult As Long
Dim strQuery As String

  
  strQuery = "SELECT MAX(" & strIDField & ") FROM " & strTable
  lngResult = GetQueryValue(conn, strQuery, , True)
  lngResult = lngResult + 1
  
GetNextID = lngResult
End Function
      

Public Function DeleteQuery(strTableName As String, _
                    strCritaria As String) As String
Dim strResult As String
  If Not (Trim(strCritaria) = "") Then
    strResult = "DELETE FROM " & strTableName & " WHERE " & strCritaria
  Else
    Exit Function
  End If
DeleteQuery = strResult
End Function
Public Function rsToDic(rs As Recordset, _
    strField As String, _
    Optional strOtherField As String = "") As Dictionary
Dim dicResult As Dictionary
Dim lngCounter As Long
Dim lngTotalRecors As Long
Dim varKey As Variant
  Set dicResult = New Dictionary
  If (Not rs.EOF) Or (Not rs.BOF) Then
    rs.MoveLast
    lngTotalRecors = rs.RecordCount
    rs.MoveFirst
  End If

  For lngCounter = 0 To lngTotalRecors - 1
    If strOtherField = "" Then
      varKey = lngCounter
    Else
      varKey = rs.Fields(strOtherField).Value
    End If
    dicResult.Add varKey, rs.Fields(strField).Value
    rs.MoveNext
  Next lngCounter
Set rsToDic = dicResult
End Function



Sub AutoCompleteDB(cmbGiven As ComboBox, _
          conn As Connection, _
          intLastKeyDown As Integer, _
          intResults As Integer, _
          strTable As String, _
          strField As String, _
          strExtraCritaria As String, _
          bolIgnoreEmptyConnection As Boolean)
          
Dim strGiven As String
Dim dicResult As Dictionary
Dim strTop As String
Dim strQuery As String
Dim rs As Recordset
Static bolChangedByProgram As Boolean
Dim strCritaria As String


  If bolChangedByProgram Then
    Exit Sub
  End If
  If bolIgnoreEmptyConnection Then
    If conn Is Nothing Then
      Exit Sub
    End If
  End If
  strGiven = cmbGiven.Text
  If intLastKeyDown = vbKeyDelete Or _
        intLastKeyDown = vbKeyBack Then
    If strGiven <> "" Then
      Exit Sub
    End If
  End If


  
  
  strTop = Top(intResults)
  strExtraCritaria = strExtraCritaria
  strCritaria = strField & " like '" & strGiven & "%'"
  strCritaria = AndTwo(strCritaria, strExtraCritaria)
  strCritaria = PutWhere(strCritaria)
  strQuery = "SELECT " & strTop & strField & " FROM " & strTable & strCritaria
  Set rs = GetADORecordSet(conn, strQuery)
  Set dicResult = rsToDic(rs, strField)
  bolChangedByProgram = True
  AutoCompleteDic cmbGiven, dicResult, True
  bolChangedByProgram = False
  
End Sub
Public Function Top(ByVal lngNumber As Long) As String
   If lngNumber < 1 Then Exit Function
   Top = "TOP " & CStr(lngNumber) & " "
End Function

Public Function getFieldValueList(mConn As Connection, _
            strTable As String, _
            strField As String, _
            lngMaxRecords As Long, _
            Optional bolUnique As Boolean = True, _
            Optional ByVal strCondition As String) As Dictionary
Dim dicResult As Dictionary
Dim strQuery As String
Dim rs As Recordset
Dim strUnique As String

  If bolUnique Then
    strUnique = " DISTINCT "
  End If
  
 
  strCondition = PutWhere(strCondition)
  strQuery = "SELECT " & strUnique & Top(lngMaxRecords) & strField & " FROM " & strTable & strCondition
  Set rs = GetADORecordSet(mConn, strQuery)
  Set dicResult = rsToDic(rs, strField)
  
Set getFieldValueList = dicResult
End Function
