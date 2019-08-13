VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QRS_XlHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Module : QRS_XlHist
' Purpose: A value history
' Concept: Eases value input using the Dropdown list of the
'          Range.Validation object provided by Excel
'          Prevents the dialog box from popping up
'          if unlisted values are entered
'          The list values are kept in memory unless the
'          list top left cell is specified
'          In order to minimize memory operations,
'          The buffer is a wrapped (circular) list
'          Values are added only when they have changed
' Use A    use for convenience
'          1. Instantiate class
'          2. Specify if no drop-down list is required
'          3. Set list top cell
'             eventually initalize buffer from there
'          4. Set linked cell
'          5. Call Update method this class
'             upon the worksheet_Change event
'          6. Disable buffer before large calculations
'          Every operation is stored (caller in worksheet)
'          Undo recalls previous value at cell
'          Redo may undo the previous undo
' Use B    use as moving drop-down list container
'          1. Instantiate class
'          2. Initalize list from 1D string list
'          3. Set linked cell
'          4. Update or refresh list eventually
' By     : QRS, Roger Strebel
' Date   : 28.01.2019                  Validation mangement optimized
' --- The public interface
' Properties
'     Enabled   (let     )             Enable or disable updates
'     Force     (let     )             Force list values
'     Length    (let     )             History length
'     NoDropDn  (set     )             Prevent drop-down list
'     RangeLink (set     )             Set linked cell
'     RangeList (set     )             Set list cell
' Methods
'     Clear                            Clear history
'     Update                           Update value change
'     SetLstStr                        Set list from string list
' --- The private sphere
'     List2Hist                        Drop-down list string
'     ListCur1                         Initial current element index
'     ListIsNew                        Add only new elements to list
'     ListRangeRead                    Input history from list range
'     ListRangeWrite                   Output histroy to list range
'     ListUpdate                       Append a value to the circular list
'     ValidationClear                  Clean up validation object
'     ValidationSetUp                  Set up validation for dropdown list

Private Const MClMaxLen As Long = 64   ' --- Maximum legth

Private MlLen As Long                  ' --- Cycle length
Private MlCur As Long                  ' --- Current
Private MbDDn As Boolean               ' --- has drop-down list
Private MbAbl As Boolean               ' --- updating enabled
Private MbFor As Boolean               ' --- validation forced
Private MbSet As Boolean               ' --- Drop-down list set
Private MbClr As Boolean               ' --- Previous validation cleared

Private MaLink As Range                ' --- Linked cell
Private MaList As Range                ' --- History list top

Private MvVal()                        '     Values (2D)
Private MsLst As String                ' --- List string
Private MsSep As String                '     Separator string

Public Property Set RangeLink(aRng As Range)

' Specify the linked range
' If a previous linked range was set its dropdown list is removed
' but the buffer is not cleared

   If aRng Is Nothing Then Exit Property

   ValidationClear MaLink

   Set MaLink = aRng

   ValidationSetUp MaLink

End Property

Public Property Set RangeList(aRng As Range)

' Specify the history list display range
' if a previous list range was set, that range is left unaltered
'

   If aRng Is Nothing Then Exit Property

   Set MaList = aRng

   ListRangeRead aRng

End Property

Public Property Let Enable(bEnable As Boolean)

   MbAbl = bEnable

End Property

Public Property Let Force(bForce As Boolean)

   MbFor = bForce                      ' --- memorize new state
                                       '     update if validation set up
   If MbSet Then MaLink.Validation.ShowError = MbFor

End Property

Public Property Let Length(lDepth As Long)

   MlLen = QRS_Lib0.MinOfL(2, lDepth, MClMaxLen)
   If Not MlLen > 0 Then Exit Property

   QRS_LibArr.ArrAllocV MvVal(), MlLen, 1

End Property

Public Sub Update(aRng As Range)

   If Not MbAbl Then Exit Sub          ' --- Only if enabled

   If MaLink Is Nothing Then Exit Sub  ' --- Check for proper state

   If aRng Is Nothing Then Exit Sub

   If Not aRng.Address = MaLink.Address Then Exit Sub
   If Not aRng.Parent.Name = MaLink.Parent.Name Then Exit Sub

   ListUpdate MaLink.Value
   ListRangeWrite MaList
   ValidationSetUp MaLink

End Sub

Public Sub SetLstVAr(vArr(), lCol As Long)

' Sets the values list from a 2D variant array

   QRS_LibArr.GetSubVarVar vArr(), MvVal(), , , lCol, lCol
   QRS_LibArr.ArrBoundV MvVal(), , MlLen

End Sub

Public Sub SetLstStr(sLst() As String)

' Sets the values list from a 1D string list
   
   Dim lR1 As Long
   
   QRS_LibLst.LstBoundS sLst(), lR1, MlLen
   QRS_LibArr.ArrAllocV MvVal(), MlLen, 1
   QRS_LibA2L.PutColStrVar sLst(), MvVal()

End Sub

Private Function ListCur1() As Long

' Returns the initial "current" element in the history

   Const Cl01 As Long = 1

   Dim lEle As Long, lCnt As Long

   For lEle = Cl01 To MlLen
      If MvVal(lEle, Cl01) = "" Then Exit For
      lCnt = lCnt + Cl01
   Next lEle

   If lCnt = MlLen Then lCnt = 0

   ListCur1 = lCnt + Cl01

End Function

Private Function List2Hist() As String

' Assembles the string for the drop-down list
' The point is to assemble the values in reverse order
' to have the latest entries first

   Const Cl01 As Long = 1

   Dim lEle As Long, lCnt As Long
   Dim sVal As String, sLst As String

   lEle = MlCur
   Do
      lEle = lEle - 1
      If lEle < 1 Then lEle = MlLen
      sVal = MvVal(lEle, Cl01)
      If Not sVal = "" Then
         If lCnt > 0 Then sLst = sLst & MsSep
         sLst = sLst & sVal
         lCnt = lCnt + 1
      End If
   Loop Until lEle = MlCur
                   
   List2Hist = sLst                    ' --- Return string for drop-down list

End Function

Private Sub ListUpdate(v)

' Increments/wraps the pointer and stores the elements

   Const Cl01 As Long = 1

   If Not MlLen > 0 Then Exit Sub      ' --- Invalid buffer length: bail out
                                       ' --- Ignore ill-representable values
   If v = "" Or IsEmpty(v) Or IsNull(v) Then Exit Sub

   If QRS_LibFnd.ColFind1_V(v, MvVal(), 1) > 0 Then Exit Sub

   MvVal(MlCur, Cl01) = v              ' --- Store

   MlCur = MlCur + Cl01                ' --- increment
   If MlCur > MlLen Then MlCur = Cl01  ' --- wrap

End Sub

Private Sub ListRangeRead(aRng As Range)

' Reads the list table from the list range
' If the range is empty, and an internal list exists, nothing is changed

   If aRng Is Nothing Then Exit Sub    ' --- No range: can't do anything

   If aRng.Value = "" Then             ' --- No data, No list yet: create empty
      If Not QRS_LibLst.LstIsAllV(MvVal()) Then
         QRS_LibLst.LstAllocV MvVal(), MlLen
      End If                               ' otherwise don't clear
      Exit Sub
   End If
   If MlLen = 0 Then MlLen = -1        ' --- No length set: autosize

   QRS_LibXL.RngGetArrV aRng, MvVal(), MlLen

   MlCur = ListCur1                    ' --- Continue past end

End Sub

Private Sub ListRangeWrite(aRng As Range)

' Outputs the list table to the list range
' If not all elements are initialized, output empty values

   If aRng Is Nothing Then Exit Sub

   QRS_LibXL.RngPutArrV aRng, MvVal()

End Sub

Private Sub ValidationClear(aRng As Range)

' Delete any validation conditions of the
' validation object for the specified cell

   If Not aRng Is Nothing Then
      aRng.Validation.Delete
   End If

   MbSet = False

End Sub

Private Sub ValidationSetUp(aRng As Range)

' Set up the range validation object to
'  - delete any existing validation conditions
'  - display an In-Cell-Drop-Down list if cell selected
'    (requires list of valid values, operator "between")
'  - disable stop on error if no force required
' Alert styles
'    xlValidAlertInformation: Information, "OK" accepts
'    xlValidAlertWarning    : Warning, "Ja" accepts
'    xlValidAlertStop       : Stop, no way to accept

   Dim sLst As String

   sLst = List2Hist

   If Not aRng Is Nothing Then
      With aRng.Validation
         If Not MbSet Then             ' --- Not yet set
            If Not MbClr Then
               .Delete                 ' --- Delete previous validations
               MbClr = True
            End If
            If Not sLst = "" Then      ' --- Non-empty list available -> set
               .Add xlValidateList, xlValidAlertStop, xlBetween, sLst
               .InCellDropdown = MbDDn
               MbSet = True
               .IgnoreBlank = True
               .ShowError = MbFor
            End If
         Else                          ' --- Already set
            .Modify Formula1:=sLst     '     Modify list
         End If
      End With
   End If

End Sub

Private Sub Class_Initialize()
                                       ' --- Validation List separator
                                       '     Didn't prove to work OK
'   MsSep = Application.International(xlListSeparator)
   MsSep = ","
   Me.Length = 6
   MlCur = 1                           ' --- Current is set to 1
   MbFor = False
   MbAbl = True
   MbDDn = True

End Sub

Private Sub Class_Terminate()

   ValidationClear MaLink

   Erase MvVal()

End Sub
