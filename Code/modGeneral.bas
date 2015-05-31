Attribute VB_Name = "modGeneral"
Option Explicit
Enum DataType
   TypeInteger = 0
   typestring = 1
   TypeBoolean = 2
   TypeDecimal = 3
   TypeDate = 4
   TypeEmpty = 5
   typelong = 6
   typenumber = 7
   TypeCurrency = 8
End Enum
Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const CB_SHOWDROPDOWN = &H14F
Public Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Sub UnExpectedEror(errGivenError As ErrObject, strModuleName As String, _
  strLocation As String, _
  Optional bolEnd As Boolean = False)
  If errGivenError Is Nothing Then
    MsgBox "UnExpected Error in " & strModuleName & "/" & strLocation & "'"
  Else
    MsgBox "UnExpected Error in " & strModuleName & "/" & strLocation & vbCrLf & _
           "Error Number is " & errGivenError.Number & vbCrLf & _
          "Description is >> " & errGivenError.Description & vbCrLf, vbCritical
  End If
  Debug.Assert False
  If bolEnd Then
    End
  End If
End Sub

Public Sub SetControlValue(ctrGiven As Control, _
      varGivenValue As Variant, _
      SelectIt As Boolean, _
      Optional dicComboValue As Dictionary = Nothing, _
      Optional bolUseKey As Boolean = False)
Dim intValue As Integer
Dim strNeeded As String

   Select Case TypeName(ctrGiven)
    Case "TextBox"
       ctrGiven.Text = varGivenValue
       If SelectIt Then
          ctrGiven.SelStart = 0
          ctrGiven.SelLength = Len(ctrGiven.Text)
      End If
    Case "ComboBox"
       If (dicComboValue Is Nothing) Then
         If IsNumeric(varGivenValue) Then
           ctrGiven.ListIndex = varGivenValue - 1
          End If
        Else
          If bolUseKey Then
            strNeeded = GetKeyOfItem(dicComboValue, varGivenValue)
            SetComboValue ctrGiven, strNeeded
          Else
            strNeeded = varGivenValue
            If varGivenValue = "" Then
              ctrGiven.ListIndex = -1
            Else
             ctrGiven.ListIndex = varGivenValue - 1
            End If
          End If
          
        End If
    Case "CheckBox"
     intValue = PutCheckBoxValue((varGivenValue))
     ctrGiven.Value = intValue
    Case "DTPicker"
       If IsEmpty(varGivenValue) Or (varGivenValue = "") Then
          varGivenValue = Date
       End If
       
      ctrGiven.Value = Round(CDate(varGivenValue))
    Case Else
       MsgBox "Not Implemented Yet " & vbCrLf & _
             "SetValue Conrols other than textbox and ComboBox", vbCritical
      Debug.Assert False
   End Select
End Sub
Public Function GetKeyOfItem(dicGiven As Dictionary, _
                        varGivenItem As Variant, _
                        Optional ByVal bolReturnFoundStatus As Boolean = False, _
                        Optional ByRef bolFoundStatus As Boolean) As Variant
Dim intCounter As Integer
Dim varItem As Variant

  bolFoundStatus = True
  For intCounter = 0 To dicGiven.Count - 1
    varItem = dicGiven.Items(intCounter)
    If IsNumeric(varItem) Then 'A Possible bug
      If Val(varItem) = varGivenItem Then
         GetKeyOfItem = dicGiven.Keys(intCounter)
         Exit Function
      End If
    Else
      If varItem = varGivenItem Then
         GetKeyOfItem = dicGiven.Keys(intCounter)
         Exit Function
      End If
    End If
  Next intCounter
  If bolReturnFoundStatus Then
    bolFoundStatus = False
  Else
    If varGivenItem <> "" Then
      Err.Raise 1000, , "Item not Found"
    End If
  End If
End Function
Public Sub SetComboValue(cmbGiven As ComboBox, _
             strGiven As String, _
             Optional bolRemoveIfNotFound As Boolean = True)
Dim intIndex As Integer
Dim intCounter As Integer
Dim strCurrent As String
  For intCounter = 0 To cmbGiven.ListCount - 1
    strCurrent = cmbGiven.List(intCounter)
    If strCurrent = strGiven Then Exit For
  Next intCounter
  intIndex = intCounter
  If intIndex < cmbGiven.ListCount Then
    cmbGiven.ListIndex = intIndex
  Else
    If bolRemoveIfNotFound Then
      cmbGiven.ListIndex = -1
    End If
  End If
  
End Sub

Public Function MyCbool(varGiven As Variant) As Boolean
  If varGiven = True Then
    MyCbool = True
  Else
    MyCbool = False
  End If
End Function

Public Function PutCheckBoxValue(bolGiven As Boolean) As Integer
Dim valResult As Integer
  If bolGiven Then
    valResult = 1
  Else
    valResult = 0
  End If
  PutCheckBoxValue = valResult
End Function
Public Sub ValidateValue(ByRef varValue As Variant, typGiven As DataType)
Dim varResult As Variant
  If IsEmpty(varValue) Then
    varResult = GetDefault(typGiven, False)
  Else
    varResult = varValue
  End If
  varValue = varResult
End Sub

Public Function GetDefault(typeGiven As DataType, NumberIsEmpty As Boolean) As Variant
Dim varResult As Variant
  Select Case typeGiven
    Case TypeBoolean
      varResult = False
    Case typestring
      varResult = ""
    Case TypeBoolean
      varResult = False
    Case TypeDate
      varResult = Date
    Case Else
      If NumberIsEmpty Then
        varResult = Empty
      Else
       varResult = 0 'It will work for long integer float etc
      End If
  End Select
  GetDefault = varResult
End Function

Function SimpleDic(varKey, varItem) As Dictionary
Dim dicResult As Dictionary
  Set dicResult = New Dictionary
  dicResult.Add varKey, varItem
Set SimpleDic = dicResult
End Function

Public Sub SelectAll(txtGiven As TextBox, _
    Optional bolSetFocus As Boolean = False)
  txtGiven.SelStart = 0
  txtGiven.SelLength = Len(txtGiven.Text)
  If bolSetFocus Then
    SecureSetFocus txtGiven
  End If
End Sub


Public Sub SecureSetFocus(ctrGiven As Control)

  If ctrGiven.Visible And ctrGiven.Enabled Then
      ctrGiven.SetFocus
        ctrGiven.SetFocus
  End If
End Sub

Public Function UniqueName(strName As String, intIndex As Integer)
    UniqueName = (strName & "(" & intIndex & ")")
End Function

Public Function UniqueControlName(ctrGiven As Control) As String
   If ctrGiven Is Nothing Then Exit Function
    UniqueControlName = UniqueName(ctrGiven.Name, ForcedIndex(ctrGiven))
End Function
Public Function ForcedIndex(ctrGiven As Control) As Integer
   On Error GoTo ErrorHandler
   ForcedIndex = ctrGiven.Index
   Exit Function
ErrorHandler:
  ForcedIndex = 0
End Function

Public Sub PrintDic(dicGiven As Dictionary, _
         Optional strName As String, _
         Optional UpperLine As Boolean = True, _
         Optional LowerLine As Boolean = True)
Dim intCounter As Integer
Dim strLine As String
  If dicGiven Is Nothing Then
    If strName = "" Then
      Debug.Print "Dictionary is Nothing"
     Else
       Debug.Print strName & " is Nothing "
     End If
     Exit Sub
  End If
  strLine = "********************"
  If UpperLine Then
     Debug.Print strLine & strName & strLine
  End If
  For intCounter = 0 To dicGiven.Count - 1
     Debug.Print intCounter, dicGiven.Keys(intCounter), dicGiven.Items(intCounter)
  Next intCounter
  If LowerLine Then
     Debug.Print strLine & MultiplyString("*", Len(strName)) & strLine
  End If
End Sub
Public Function MultiplyString(strGiven As String, intTimes As Integer) As String
Dim strResult As String
Dim intCounter As Integer
  If intTimes < 1 Then Exit Function
  strResult = strGiven
  For intCounter = 1 To intTimes - 1
    strResult = strResult & strGiven
  Next intCounter
MultiplyString = strResult
End Function

Public Function GetControlValue(ctrGiven As Control, _
      Optional TrimText As Boolean = True, _
      Optional bolIgnoreIfDisable As Boolean = True, _
      Optional dicCombo As Dictionary = Nothing) As Variant
Dim varResult As Variant
Dim intValue As Integer

  Select Case TypeName(ctrGiven)
    Case "TextBox"
      varResult = ctrGiven.Text
      If TrimText Then
        varResult = Trim(varResult)
      End If
    Case "ComboBox"
      If dicCombo Is Nothing Then
       If ctrGiven.Style = 2 Then
        varResult = ctrGiven.ListIndex
       Else
        varResult = ctrGiven.Text
       End If
      Else
        varResult = ComboValue(dicCombo, ctrGiven)
      End If
    Case "CheckBox"
       varResult = vbBool(ctrGiven.Value)
     Case "DTPicker"
       varResult = Round(CDate(ctrGiven.Value))
    Case Else
       MsgBox "Not Implemented Yet " & vbCrLf & _
             "SetValue Conrols other than textbox,Date Time Picker and ComboBox", vbCritical
      Debug.Assert False
  End Select
  If bolIgnoreIfDisable And (ctrGiven.Enabled = False) Then
    varResult = Empty
  End If
  
GetControlValue = varResult
End Function

Public Sub ExpectedError(strMessage As String, _
          Optional strTitle As String = "Unexpected Error", _
          Optional strLocaltion As String, _
          Optional bolEnd As Boolean = False)
    MsgBox strMessage, vbCritical, strTitle
     Debug.Assert False
     If bolEnd Then End
End Sub


Public Function ComboValue(dicGiven As Dictionary, cmbGiven As ComboBox) As Variant
Dim varValue As Variant
Dim varResult As Variant
   If cmbGiven.Text = "" Then
     varResult = Empty
   Else
    varValue = cmbGiven.Text
    varResult = dicGiven.Item(varValue)
  End If
   ComboValue = varResult
End Function

Public Function vbBool(varGiven As Variant) As Boolean
Dim bolResult As Boolean

  If varGiven = 0 Then
    bolResult = False
  Else
    bolResult = True
  End If
  
vbBool = bolResult
End Function

Public Sub FillDicCombo(cmbGiven As ComboBox, _
          dicFields As Dictionary, _
          bolUseKeys As Boolean)
Dim lngCounter As Long
    
    cmbGiven.Clear
    If cmbGiven.Style <> 2 Then
        cmbGiven.Text = ""
    End If
    If dicFields Is Nothing Or dicFields.Count = 0 Then Exit Sub
    
    
    For lngCounter = 0 To dicFields.Count - 1
        If bolUseKeys Then
            cmbGiven.AddItem dicFields.Keys(lngCounter)
        Else
          cmbGiven.AddItem dicFields.Items(lngCounter)
        End If
    Next lngCounter
    
End Sub

Public Sub AdjustPicture(pboxGiven As PictureBox)
Dim lngPictureHeight As Long
Dim lngPictureWidth As Long
Dim sngPictureRatio As Single
Dim bolLongPicture As Boolean
Dim lngAdjustHeight As Long
Dim lngAdjustWidth As Long
Dim bololdAutoDraw As Boolean
Dim sngBoxRatio As Single
Dim lngBoxWidth As Long
Dim lngBoxHeight As Long
Dim sngDesiredRatio As Single
   lngPictureHeight = pboxGiven.Picture.Height
   lngPictureWidth = pboxGiven.Picture.Width
   lngBoxHeight = pboxGiven.Height
   lngBoxWidth = pboxGiven.Width
   
   If (lngPictureWidth = 0) Or (lngPictureHeight = 0) Then
     Exit Sub
   End If
   sngPictureRatio = lngPictureHeight / lngPictureWidth
   sngBoxRatio = lngBoxHeight / lngBoxWidth
   'bolLongPicture = lngPictureHeight > lngPictureWidth
   sngDesiredRatio = sngPictureRatio / sngBoxRatio
   bolLongPicture = sngDesiredRatio > 1
   If bolLongPicture Then
     lngAdjustHeight = pboxGiven.Height
     lngAdjustWidth = Round(lngAdjustHeight / sngPictureRatio)
   Else 'Wide Picture
     lngAdjustWidth = pboxGiven.Width
     lngAdjustHeight = Round(lngAdjustWidth * sngPictureRatio)
   End If
   
   bololdAutoDraw = pboxGiven.AutoRedraw
   pboxGiven.AutoRedraw = False
   pboxGiven.Cls
   pboxGiven.Line (0, 0)-(lngBoxWidth, lngBoxHeight), pboxGiven.BackColor, BF
   pboxGiven.PaintPicture pboxGiven.Picture, 0, 0, lngAdjustWidth, lngAdjustHeight
   pboxGiven.AutoRedraw = bololdAutoDraw
'   If bolLongPicture Then
'     pboxGiven.Line (lngAdjustWidth, 0)-(lngPictureWidth, lngPictureHeight), pboxGiven.BackColor, BF
'   Else
'     pboxGiven.Line (0, lngAdjustHeight)-(lngPictureWidth, lngPictureHeight), pboxGiven.BackColor, BF
'   End If
   
End Sub
Sub EnterKeyPressed(intKeyAscii As Integer)
  If intKeyAscii = 13 Then
    intKeyAscii = 0
    SendKeys "{TAB}"
  End If
End Sub

Sub FormActivateOnce(frmGiven As Form, _
      bolUnLoaded As Boolean)
Static dicForms As Dictionary
'Dim bolAgain As Boolean
Dim bolFirstTime As Boolean

    If bolUnLoaded Then 'Unloaded so if function is called again Call It Again
      dicForms.Remove frmGiven.Name
    Else 'Activated
      If dicForms Is Nothing Then
        Set dicForms = New Dictionary
      End If
       
      If dicForms.Exists(frmGiven.Name) Then
        bolFirstTime = False
      Else
        bolFirstTime = True
        dicForms.Add frmGiven.Name, "Dummy"
      End If
      'bolFirstTime = Not bolAgain
      
      If bolFirstTime Then
        frmGiven.FormStart
      End If
    End If

End Sub
Function RemoveEnters(ByVal strGiven As String, _
      Optional bolTrim As Boolean = True) As String
Dim strResult As String
Dim lngCounter As Long
Dim bolLoop As Boolean

  If bolTrim Then
    strResult = Trim(strGiven)
  End If
  If strResult = "" Then
    RemoveEnters = strResult
    Exit Function
  End If
  
  bolLoop = True
  Do While bolLoop
    If (Asc(Left(strResult, 1)) = 13) Or (Asc(Left(strResult, 1)) = 10) Then
      strResult = Right(strResult, Len(strResult) - 1)
    Else
      bolLoop = False
    End If
  Loop
  
  bolLoop = True
  Do While bolLoop
    If (Asc(Right(strResult, 1)) = 13) Or (Asc(Right(strResult, 1)) = 10) Then
      strResult = Left(strResult, Len(strResult) - 1)
    Else
      bolLoop = False
    End If
  Loop
  
  'strResult = Replace(strGiven, vbCrLf, "")
RemoveEnters = strResult
End Function
Function ComboValueExists(cmbGiven As ComboBox, _
              strValue As String) As Boolean
              
Dim bolResult As Boolean
Dim intCounter As Integer
Dim strItem As String

    bolResult = False
    For intCounter = 0 To cmbGiven.ListCount - 1
      strItem = cmbGiven.List(intCounter)
      If strItem = strValue Then
        bolResult = True
        Exit For
      End If
    Next intCounter


ComboValueExists = bolResult
End Function
Function ListValueExists(lstGiven As ListBox, _
              strValue As String) As Boolean
              
Dim bolResult As Boolean
Dim intCounter As Integer
Dim strItem As String

    bolResult = False
    For intCounter = 0 To lstGiven.ListCount - 1
      strItem = lstGiven.List(intCounter)
      If strItem = strValue Then
        bolResult = True
        Exit For
      End If
    Next intCounter


ListValueExists = bolResult
End Function

Function ControlsHaveValue(dicControls As Dictionary) As Boolean
Dim lngCounter As Long
Dim strMessege As String
Dim ctrLoop As Control
Dim strTitle As String
Dim bolEmpty As Boolean

  
  For lngCounter = 0 To dicControls.Count - 1
    Set ctrLoop = dicControls.Items(lngCounter)
    strTitle = dicControls.Keys(lngCounter)
    bolEmpty = ControlEmpty(ctrLoop)
    
    If bolEmpty Then
        strMessege = "Enter the Value Of '" & strTitle & "'"
        MsgBox strMessege, vbCritical, "Enter Value"
        ctrLoop.SetFocus
        ControlsHaveValue = False
        Exit Function
    End If
  Next lngCounter
  ControlsHaveValue = True
End Function
Function ControlEmpty(ctrGiven As Control, _
      Optional bolTrimText As Boolean = True, _
      Optional bolListIndexInCombos As Boolean = True, _
      Optional bolIgnoreIfDisable As Boolean = True) As Boolean
Dim bolResult As Boolean
Dim varResult As Variant
Dim intValue As Integer

  Select Case TypeName(ctrGiven)
    Case "TextBox"
      If bolTrimText Then
        bolResult = (Trim(ctrGiven.Text) = "")
      Else
        bolResult = (ctrGiven.Text = "")
      End If
    Case "ComboBox"
    
      If (ctrGiven.Style = 2) And bolListIndexInCombos Then
         bolResult = (ctrGiven.ListIndex = -1)
      Else
         If bolTrimText Then
           bolResult = (Trim(ctrGiven.Text) = "")
         Else
           bolResult = (ctrGiven.Text = "")
         End If
      End If
      
    Case "CheckBox"
       
     Case "DTPicker"
       bolResult = (TypeName(ctrGiven.Value) = "Null")
    Case Else
       MsgBox "Not Implemented Yet " & vbCrLf & _
             "SetValue Conrols other than textbox,Date Time Picker and ComboBox", vbCritical
      Debug.Assert False
  End Select
  If bolIgnoreIfDisable And (ctrGiven.Enabled = False) Then
    varResult = False
  End If
  
ControlEmpty = bolResult
End Function
Public Sub DropCombo(cmbGiven As ComboBox)
  SendMessage cmbGiven.hwnd, CB_SHOWDROPDOWN, 1, 0
End Sub

Public Function GetRandomUniqueKey() As String
Dim strResult As String

  strResult = Now & Rnd & Rnd
  
GetRandomUniqueKey = strResult
End Function
Sub AddDicItem(dicGiven As Dictionary, _
        varItem As Variant)
Dim lngID As Long
  
  If dicGiven Is Nothing Then
    Set dicGiven = New Dictionary
  End If
  
  lngID = dicGiven.Count
  dicGiven.Add lngID, varItem
  
End Sub


Sub AutoCompleteDic(cmbGiven As ComboBox, _
            dicGiven As Dictionary, _
            intLastKeyDown As Integer, _
            Optional bolIgnoreChangeByProgram As Boolean = False)
            
Static bolChangedByProgram As Boolean
Dim lngCounter As Long
Dim strCurrentText As String 'To save it incase not found
Dim intSelStart As Integer
Dim intSelLength As Integer

  If bolChangedByProgram And (Not bolIgnoreChangeByProgram) Then
    Exit Sub
  End If
  
  If intLastKeyDown = vbKeyDelete Or _
        intLastKeyDown = vbKeyBack Then
    Exit Sub
  End If
  
  strCurrentText = cmbGiven.Text

  If dicGiven.Count = 0 Then
    
    cmbGiven.Clear
    SendMessage cmbGiven.hwnd, CB_SHOWDROPDOWN, _
     0, 0
     cmbGiven.Text = strCurrentText
     cmbGiven.SelStart = Len(strCurrentText)
    Exit Sub
  End If
  SendMessage cmbGiven.hwnd, CB_SHOWDROPDOWN, _
     0, 0
  cmbGiven.Clear
  If strCurrentText <> "" Then
    cmbGiven.Text = dicGiven.Items(0)
  End If
  
  For lngCounter = 0 To dicGiven.Count - 1
    cmbGiven.AddItem dicGiven.Items(lngCounter)
  Next lngCounter
  
    If cmbGiven.ListCount <> 0 Then
    SendMessage cmbGiven.hwnd, CB_SHOWDROPDOWN, _
       1, 0
  End If
  
  intSelStart = Len(strCurrentText)
  intSelLength = Len(cmbGiven.Text)
  
  cmbGiven.SelStart = intSelStart
  cmbGiven.SelLength = intSelLength
  
  


  bolChangedByProgram = False
End Sub


Function GetMiliSeconds() As Long
Dim stuDT As SYSTEMTIME
  GetLocalTime stuDT
  GetMiliSeconds = stuDT.wMilliseconds
  
End Function
Function MakeFixedLength(strGiven As String, _
          intLength As Integer, _
          Optional bolTrim As Boolean = True, _
          Optional strCharacter As String = " ", _
          Optional bolSpaceOnLeft As Boolean = False)
Dim strResult As String
Dim intAdditional As Integer
Dim strSpaces As String

  If bolTrim Then
    strGiven = Trim(strGiven)
  End If
  intAdditional = intLength - Len(strGiven)
  If intAdditional > 0 Then
    strSpaces = MultiplyString(" ", intAdditional)
    If bolSpaceOnLeft Then
      strResult = strSpaces & strGiven
    Else
      strResult = strGiven & strSpaces
    End If
  Else
    strResult = strGiven
  End If
MakeFixedLength = strResult
End Function

Public Sub EnableConrol(ctrGiven As Control)
Dim txtTemp As TextBox
Dim cmbTemp As ComboBox
  If ctrGiven Is Nothing Then Exit Sub
  ctrGiven.Enabled = True
  If TypeName(ctrGiven) = "TextBox" Then
     Set txtTemp = ctrGiven
     txtTemp.BackColor = vbWhite
  End If
  If TypeName(ctrGiven) = "ComboBox" Then
     Set cmbTemp = ctrGiven
     cmbTemp.BackColor = vbWhite
  End If
End Sub

Public Sub DisableControl(ctrGiven As Control)
Dim txtTemp As TextBox
Dim cmbTemp As ComboBox
  If ctrGiven Is Nothing Then Exit Sub
  ctrGiven.Enabled = False
  If TypeName(ctrGiven) = "TextBox" Then
     Set txtTemp = ctrGiven
     txtTemp.BackColor = vbBlack
  End If
  If TypeName(ctrGiven) = "ComboBox" Then
     Set cmbTemp = ctrGiven
     cmbTemp.BackColor = vbBlack
  End If
End Sub

Public Sub SelectAllCombo(cmbGiven As ComboBox, _
    Optional bolSetFocus As Boolean = False)
  cmbGiven.SelStart = 0
  cmbGiven.SelLength = Len(cmbGiven.Text)
  If bolSetFocus Then
    SecureSetFocus cmbGiven
  End If
End Sub


Public Function ItemExists(dicGiven As Dictionary, _
                        varGivenItem As Variant) As Boolean
Dim bolResult As Boolean
Dim lngCounter As Long
Dim varItem As Variant

  bolResult = False
  
  For lngCounter = 0 To dicGiven.Count - 1
    varItem = dicGiven.Items(lngCounter)
    If IsNumeric(varItem) Then 'A Possible bug (May Be Inherited from "GetKeyOfItem")
      If Val(varItem) = varGivenItem Then
         bolResult = True
         GoTo OutSide
      End If
    Else
      If varItem = varGivenItem Then
         bolResult = True
         GoTo OutSide
      End If
    End If
  Next lngCounter
OutSide:
  ItemExists = bolResult
End Function


Public Sub DicToList(lstGiven As ListBox, dicGiven As Dictionary)
Dim intCounter As Integer
  If dicGiven Is Nothing Then Exit Sub
  lstGiven.Clear
  For intCounter = 0 To dicGiven.Count - 1
    lstGiven.AddItem dicGiven.Items(intCounter)
  Next intCounter
End Sub
Public Sub ListToDic(ByRef dicGiven As Dictionary, _
          lstGiven As ListBox)
Dim intCounter As Integer
  
  Set dicGiven = Nothing
  Set dicGiven = New Dictionary
  For intCounter = 0 To lstGiven.ListCount - 1
     dicGiven.Add intCounter, lstGiven.List(intCounter)
  Next intCounter

End Sub


Public Function GetListDic(lstGiven As ListBox) As Dictionary
Dim dicResult As Dictionary
  ListToDic dicResult, lstGiven
Set GetListDic = dicResult
End Function

Function SecurelyGetDicValue(dicGiven As Dictionary, _
          varKey As Variant) As Variant
Dim varResult As Variant
  
  If dicGiven.Exists(varKey) Then
    If IsObject(dicGiven.Item(varKey)) Then
      Set varResult = dicGiven.Item(varKey)
    Else
      varResult = dicGiven.Item(varKey)
    End If
  Else
    MsgBox "Item not found in collection", vbCritical, "Item Not Found"
    Debug.Assert False
  End If
  
  If IsObject(varResult) Then
    Set SecurelyGetDicValue = varResult
  Else
    SecurelyGetDicValue = varResult
  End If
  
End Function
Function GetBetweenText(strContent As String, _
          strStart As String, _
          strEnd As String) As String
Dim strResult As String
Dim lngStart As Long
Dim lngEnd As Long
Dim lngLength As Long
Dim lngSearchFrom As Long


  lngStart = InStr(strContent, strStart)
  If strEnd = "" Then
    lngEnd = Len(strContent) + 1
  Else
    lngSearchFrom = lngStart + Len(strStart)
    lngEnd = InStr(lngSearchFrom, strContent, strEnd)
  End If
  lngLength = Len(strContent) - (lngStart + Len(strStart)) - (Len(strContent) - lngEnd)
  strResult = Mid(strContent, lngStart + Len(strStart), lngLength)
  
GetBetweenText = strResult
End Function

Public Function StartFile(strPath As String, _
              Optional strParameters As String = "", _
              Optional strOperation As String = "open", _
              Optional lngHwnd As Long = 0, _
              Optional strDirectory = "", _
              Optional ShowStyle = 3, _
              Optional bolUseAppPathIfEmpty As Boolean = True) As Long
Dim lngResult As Long
   If bolUseAppPathIfEmpty Then
     If strDirectory = "" Then
       strDirectory = App.Path
     End If
   End If
   lngResult = ShellExecute(lngHwnd, strOperation, strPath, strParameters, strDirectory, ShowStyle)
StartFile = lngResult
End Function


Public Sub ConfirmFolder(strFolder As String)
Dim fs As FileSystemObject
Dim strParentFolder As String

  Set fs = New FileSystemObject
  If Not fs.FolderExists(strFolder) Then
   strParentFolder = fs.GetParentFolderName(strFolder)
   If fs.FolderExists(strParentFolder) Then
      fs.CreateFolder strFolder
    Else
      ConfirmFolder strParentFolder
      fs.CreateFolder strFolder
    End If
  End If
End Sub

Function IsRelativePath(strPath As String) As Boolean
Dim bolResult As Boolean

  bolResult = (InStr(strPath, ":") <= 0)

IsRelativePath = bolResult
End Function
