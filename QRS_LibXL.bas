Attribute VB_Name = "QRS_LibXL"
Option Explicit

' Module : QRS_LibXL
' Project: any
' Purpose: VBA utility library for Excel object integration
'          The library processes objects but does not retain
'          any state. States are managed by classes
' By     : QRS, Roger Strebel
' Date   : 21.01.2018
'          18.02.2018                  Column ABC-123 conversion added
'          04.03.2018                  Range input to array functions added
'          06.03.2018                  Masterpiece RefXtr_Ele added
'          13.03.2018                  RngGetArrV improved, RefSetObj added
'          18.03.2018                  RefGetArr improved, RngPutArrV added
'          19.03.2018                  RngPutArr direct output for all types
'          22.03.2018                  Range clear, color, fill added
'          04.07.2018                  RefGetColV, RefGetRowV added
'          21.07.2018                  Offset reference functions added
'          25.07.2018                  ReFromShRg added
'          15.08.2018                  RefPutColV bug fixed
'          07.02.2019                  RngAreSame added
' --- The public interface
'     Chr1232ABC                       Value 1..26 to A..Z           18.02.2018
'     ChrABC2123                       Character A..Z to 1..26       18.02.2018
'     Col1232ABC                       Column ref numeric to alpha   21.07.2018
'     ColABC2123                       Column ref alpha to numeric   18.02.2018
'     GetXlLngID                       Obtain language ID for Excel
'     RefFromEle                       Ref string from elements      06.03.2018
'     ReFromShRg                       Ref from sheet and range      25.07.2018
'     RefGetArrD                       Ref input to array of dates   18.03.2018
'     RefGetArrF                       Ref input to array of real    18.03.2018
'     RefGetArrL                       Ref input to array of integer 18.03.2018
'     RefGetArrS                       Ref input to array of strings 18.03.2018
'     RefGetArrV                       Ref input to array of variant 18.03.2018
'     RefGetColV                       Ref col input to variables    04.07.2018
'     RefGetRowV                       Ref row input to variables    04.07.2018
'     RefGetTxt                        Get object names from objects 13.03.2018
'     RefPutArrF                       Output to Ref array of real   19.03.2018
'     RefPutArrL                       Output to Ref array of long   19.03.2018
'     RefPutArrS                       Output to Ref array of string 19.03.2018
'     RefPutArrV                       Output to Ref array of var    19.03.2018
'     RefPutColV                       Variable output to Ref column 04.07.2018
'     RefSetObj                        Set Excel objects from names  13.03.2018
'     RefStrOff                        Reference with offset         21.07.2018
'     RefXLGetRef                      Ref string from elements      07.03.2018
'     RefXLGetTxt                      Get object names in RefXL     21.01.2018
'     RefXLSetObj                      Set Excel objects in RefXL    13.03.2018
'     RefXLSetOff                      Set Reference with offset     20.07.2018
'     RefXLXtrEle                      Get RefXL elements from sRef  07.03.2018
'     RefXtr_Ele                       Extract sRef elements         06.03.2018
'     RngClrVal                        Clear values in a range       22.03.2018
'     RngAreSame                       Compare two ranges            07.02.2019
'     RngColorBG                       Background color of a range   22.03.2018
'     RngFill                          Fill range with one value     22.03.2018
'     RngGetArrF                       Range input to array of real  04.03.2018
'     RngGetArrL                       Range input to array of long  04.03.2018
'     RngGetArrS                       Range input to array of long  04.03.2018
'     RngGetArrV                       Range input to variant array  13.03.2018
'     RngGetRefXL                      Cell reference separation     21.01.2018
'     RngGetTxt                        Cell reference string parts   21.01.2018
'     RngPutArrF                       Output to range real array    19.03.2018
'     RngPutArrL                       Output to range long array    19.03.2018
'     RngPutArrS                       Output to range string array  19.03.2018
'     RngPutArrV                       Output to range variant array 19.03.2018
'     TxtRCRowCol                      Row and Col from RC string

                                       ' --- RngGetArr input size control
Public Const MCl_LibXL_Auto As Long = -1   ' Auto-detect by first empty cells
Public Const MCl_LibXL_ArSz As Long = -2   ' Use array size
Public Const MCl_LibXL_RgSz As Long = -3   ' Use range size
                                       ' --- RefSetObj Workbook default control
Public Const MCl_LibXL_WbThis As Long = -1 ' Use ThisWorkbook if not specified
Public Const MCl_LibXL_WbKeep As Long = -2 ' Keep present workbook if not spec

Public Const MCl_LibXL_ChA As Long = 64
Public Const MCl_LibXL_Chr As Long = 26

Public Type tRefXL
   sRef As String
   sCl As String
   sSh As String
   sWb As String
   sFP As String
   aCl As Range
   aSh As Worksheet
   aWb As Workbook
End Type

Public Sub RngGetArrD(aCl As Range, dLst() As Date, _
                      Optional lNEle As Long = -1, _
                      Optional bRow As Boolean = False)

' Inputs values from the specified Excel range to the list of dates
' provided in dLst(). Direct input from Excel to a table is possible
' with variant arrays only. This routine uses RngGetArrV for the input

'  o Separately handles row and columns
'  o Negative values are compared to public constant values
'    in this module:
'     -1    MCl_LibXL_Auto             Auto-detect by first empty cells
'     -2    MCl_LibXL_ArSz             Use array size (must be allocated)
'     -3    MCl_LibXL_RgSz             Use range size "TopLeft:BottomRight"
'  o Zero results in no input
'  o Positive values specify the input size

   Const Cl01 As Long = 1

   Dim vArr() As Variant
   Dim lNC As Long, lNR As Long

   If lNEle = MCl_LibXL_ArSz Then      ' --- Copy list size
      QRS_LibLst.LstAllocD dLst(), lNR '     Get list size
      If bRow Then                     ' --- Obtain row
         lNC = lNR                     '     input width
         lNR = Cl01                    '     height = 1 row
      Else                             ' --- Obtain column
         lNC = Cl01                    '     width = 1 column
      End If
      QRS_LibArr.ArrAllocV vArr(), lNR, lNC
   End If
   If bRow Then
      lNR = Cl01
      lNC = lNEle
   Else
      lNR = lNEle
      lNC = Cl01
   End If                              ' --- Other special size cases
   RngGetArrV aCl, vArr(), lNR, lNC    '     handled by RngGetArrV

   If bRow Then
      QRS_LibArr.GetRowVarDat vArr(), dLst()
   Else
      QRS_LibArr.GetColVarDat vArr(), dLst()
   End If

End Sub

Public Sub RngGetArrF(aCl As Range, fArr() As Double, _
                      Optional lNRow As Long = -1, _
                      Optional lNCol As Long = -1)

' Inputs values from the specified Excel range to the table of double
' provided in fArr(). Direct input from Excel to a table is possible
' with variant arrays only. This routine uses RngGetArrV for the input

'  o Separately handles row and columns
'  o Negative values are compared to public constant values
'    in this module:
'     -1    MCl_LibXL_Auto             Auto-detect by first empty cells
'     -2    MCl_LibXL_ArSz             Use array size (must be allocated)
'     -3    MCl_LibXL_RgSz             Use range size "TopLeft:BottomRight"
'  o Zero results in no input
'  o Positive values specify the input size

   Const Cl01 As Long = 1

   Dim vArr() As Variant
   Dim lNC As Long, lNR As Long
                                       ' --- Copy array size
   If lNRow = MCl_LibXL_ArSz Or lNCol = MCl_LibXL_ArSz Then
      QRS_LibArr.ArrAllocF fArr(), lNR, lNC
      QRS_LibArr.ArrAllocV vArr(), lNR, lNC
   End If
   RngGetArrV aCl, vArr(), lNRow, lNCol
   QRS_LibArr.GetSubVarDbl vArr(), fArr()

End Sub

Public Sub RngGetArrL(aCl As Range, lArr() As Long, _
                      Optional lNRow As Long = -1, _
                      Optional lNCol As Long = -1)

' Inputs values from the specified Excel range to the table of long integer
' provided in lArr(). Direct input from Excel to a table is possible
' with variant arrays only. This routine uses RngGetArrV for the input

'  o Separately handles row and columns
'  o Negative values are compared to public constant values
'    in this module:
'     -1    MCl_LibXL_Auto             Auto-detect by first empty cells
'     -2    MCl_LibXL_ArSz             Use array size (must be allocated)
'     -3    MCl_LibXL_RgSz             Use range size "TopLeft:BottomRight"
'  o Zero results in no input
'  o Positive values specify the input size

   Const Cl01 As Long = 1

   Dim vArr() As Variant
   Dim lNC As Long, lNR As Long
                                       ' --- Copy array size
   If lNRow = MCl_LibXL_ArSz Or lNCol = MCl_LibXL_ArSz Then
      QRS_LibArr.ArrAllocL lArr(), lNR, lNC
      QRS_LibArr.ArrAllocV vArr(), lNR, lNC
   End If
   RngGetArrV aCl, vArr(), lNRow, lNCol
   QRS_LibArr.GetSubVarLon vArr(), lArr()

End Sub

Public Sub RngGetArrS(aCl As Range, sArr() As String, _
                      Optional lNRow As Long = -1, _
                      Optional lNCol As Long = -1)

' Inputs values from the specified Excel range to the table of strings
' provided in sArr(). Direct input from Excel to a table is possible
' with variant arrays only. This routine uses RngGetArrV for the input

'  o Separately handles row and columns
'  o Negative values are compared to public constant values
'    in this module:
'     -1    MCl_LibXL_Auto             Auto-detect by first empty cells
'     -2    MCl_LibXL_ArSz             Use array size (must be allocated)
'     -3    MCl_LibXL_RgSz             Use range size "TopLeft:BottomRight"
'  o Zero results in no input
'  o Positive values specify the input size

   Const Cl01 As Long = 1

   Dim vArr() As Variant
   Dim lNC As Long, lNR As Long
                                       ' --- Copy array size
   If lNRow = MCl_LibXL_ArSz Or lNCol = MCl_LibXL_ArSz Then
      QRS_LibArr.ArrAllocS sArr(), lNR, lNC
      QRS_LibArr.ArrAllocV vArr(), lNR, lNC
   End If
   RngGetArrV aCl, vArr(), lNRow, lNCol
   QRS_LibArr.GetSubVarStr vArr(), sArr()

End Sub

Public Sub RngGetArrV(aCl As Range, vArr() As Variant, _
                      Optional lNRow As Long = -1, _
                      Optional lNCol As Long = -1)

' Inputs values from the specified Excel range to the variant table
' provided in vArr(). The array size returned is determined as follows:
'  o Separately handles row and columns
'  o Negative values are compared to public constant values
'    in this module:
'     -1    MCl_LibXL_Auto             Auto-detect by first empty cells
'     -2    MCl_LibXL_ArSz             Use array size (must be allocated)
'     -3    MCl_LibXL_RgSz             Use range size "TopLeft:BottomRight"
'  o Zero results in no input
'  o Positive values specify the input size

   Const Cl01 As Long = 1

   Dim aCl1 As Range, aClB As Range

   If aCl.Rows.Count > 1 Or aCl.Columns.Count > 1 Then
      Set aCl1 = aCl.Cells(1, 1)       ' --- Extract 1 by 1 range
   Else
      Set aCl1 = aCl
   End If

   If lNRow = MCl_LibXL_Auto Then      ' --- Auto-detect input row count
      If aCl1.Offset(Cl01, 0).Value = "" Then
         lNRow = 1                     '     one row only
      Else
         Set aClB = aCl1.End(xlDown)   '     before first empty row
         lNRow = aClB.Row + Cl01 - aCl1.Row
      End If
   End If
   If lNRow = MCl_LibXL_ArSz Then      ' --- Use array row count
      QRS_LibArr.ArrAllocV vArr(), lNRow
   End If
   If lNRow = MCl_LibXL_RgSz Then      ' --- Use range row count
      lNRow = aCl1.Rows.Count
   End If
   If lNCol = MCl_LibXL_Auto Then      ' --- Auto-detect input col count
      If aCl1.Offset(0, Cl01).Value = "" Then
         lNCol = 1                     '     one column only
      Else
         Set aClB = aCl1.End(xlToRight) '    before first empty column
         lNCol = aClB.Column + Cl01 - aCl1.Column
      End If
   End If
   If lNCol = MCl_LibXL_ArSz Then      ' --- Use array column count
      QRS_LibArr.ArrAllocV vArr(), , lNCol
   End If
   If lNCol = MCl_LibXL_RgSz Then      ' --- Use range column count
      lNCol = aCl.Columns.Count
   End If

   If lNRow > 0 Or lNCol > 0 Then
      If lNRow = 1 And lNCol = 1 Then
         QRS_LibArr.ArrAllocV vArr(), 1, 1
         vArr(1, 1) = aCl.Value
      Else
         If aCl.Rows.Count < lNRow Or aCl.Columns.Count < lNCol Then
            vArr() = Range(aCl1, aCl1.Offset(lNRow - Cl01, lNCol - Cl01)).Value
         Else                          ' --- Range size matches
            vArr() = aCl.Value
         End If
      End If
   End If

End Sub

Public Sub RefFromEle(sRef As String, _
                      Optional sFP As String = "", Optional sWb As String = "", _
                      Optional sSh As String = "", Optional sCl As String = "")

' Get full reference string with corresponding delimiters
' Cases:
'   'Path[File]Sheet'!Range
'   [File]Sheet!Range
'   Sheet!Range
'   Range

   Dim bWb As Boolean, bFP As Boolean

   bFP = Not sFP = ""                  ' --- Contains path
   bWb = Not sWb = ""                  ' --- Contains workbook
   If Not (sSh = "" And sCl = "") Then
      If bFP Then
         sRef = sSh & "'!" & sCl       ' --- Reference with workbook
      Else
         sRef = sSh & "!" & sCl        ' --- Reference on sheet
      End If
   Else
      sRef = sSh                       ' --- Refer to sheet only
      If bFP Then sRef = sRef & "'"
   End If
   If bWb Then                         ' --- Workbook
      sRef = "[" & sWb & "]" & sRef
   End If
   If bFP Then                         ' --- File path
      sRef = "'" & sFP & sRef
   End If

End Sub

Public Function ReFromShRg(sSh As String, sCl As String) As String

' Assembles local reference. Takes care of the exclamation mark

   Const CsSep = "!"

   Dim sShRg As String
   Dim lLen As Long
   Dim bIns As Boolean

   bIns = Not (sSh = "" Or sCl = "")   ' --- Only if both contain values

   If bIns Then
      lLen = Len(CsSep)
      bIns = Not (Right(sSh, lLen) = CsSep Or Left(sCl, lLen) = CsSep)
   End If
   If bIns Then
      sShRg = sSh & CsSep & sCl
   Else
      sShRg = sSh & sCl
   End If

   ReFromShRg = sShRg

End Function

Public Function RefGetArrD(sRef As String, dArr() As Date, _
                           Optional lNEle As Long = -1, _
                           Optional bRow As Boolean = False) As Boolean

' Obtains a list of dates from a reference string
' Returns true if the reference is not valid

   Dim aRefXL As tRefXL
   Dim bFail As Boolean

   aRefXL.sRef = sRef
   bFail = RefXLSetObj(aRefXL)
   If Not bFail Then RngGetArrD aRefXL.aCl, dArr(), lNEle, bRow

   RefGetArrD = bFail

End Function

Public Function RefGetArrF(sRef As String, fArr() As Double, _
                           Optional lNRow As Long = -1, _
                           Optional lNCol As Long = -1) As Boolean

' Obtains an array of real from a reference string
' Returns true if the reference is not valid

   Dim aRefXL As tRefXL
   Dim bFail As Boolean

   aRefXL.sRef = sRef
   bFail = RefXLSetObj(aRefXL)
   If Not bFail Then RngGetArrF aRefXL.aCl, fArr(), lNRow, lNCol

   RefGetArrF = bFail

End Function

Public Function RefGetArrL(sRef As String, lArr() As Long, _
                           Optional lNRow As Long = -1, _
                           Optional lNCol As Long = -1) As Boolean

' Obtains an array of long integers from a reference string
' Returns true if the reference is not valid

   Dim aRefXL As tRefXL
   Dim bFail As Boolean

   aRefXL.sRef = sRef
   bFail = RefXLSetObj(aRefXL)
   If Not bFail Then RngGetArrL aRefXL.aCl, lArr(), lNRow, lNCol

   RefGetArrL = bFail

End Function

Public Function RefGetArrS(sRef As String, sArr() As String, _
                           Optional lNRow As Long = -1, _
                           Optional lNCol As Long = -1) As Boolean

' Obtains an array of character strings from a reference string
' Returns true if the reference is not valid

   Dim aRefXL As tRefXL
   Dim bFail As Boolean

   aRefXL.sRef = sRef
   bFail = RefXLSetObj(aRefXL)
   If Not bFail Then RngGetArrS aRefXL.aCl, sArr(), lNRow, lNCol

   RefGetArrS = bFail

End Function

Public Function RefGetArrV(sRef As String, vArr() As Variant, _
                           Optional lNRow As Long = -1, _
                           Optional lNCol As Long = -1) As Boolean

' Obtains an array of real from a reference string^
' Returns true if the reference is not valid

   Dim aRefXL As tRefXL
   Dim bFail As Boolean

   aRefXL.sRef = sRef
   bFail = RefXLSetObj(aRefXL)
   If Not bFail Then RngGetArrV aRefXL.aCl, vArr(), lNRow, lNCol

   RefGetArrV = bFail

End Function

Public Function RefGetColV(sRef As String, _
                           Optional A1, Optional A2, Optional A3, _
                           Optional A4, Optional A5, Optional a6)

' Returns values at and right of sRef into the variables a1 to a6

   Dim aRefXL As tRefXL
   Dim bFail As Boolean

   aRefXL.sRef = sRef
   bFail = RefXLSetObj(aRefXL)
   If Not bFail Then
      If Not IsMissing(A1) Then A1 = aRefXL.aCl.Value
      If Not IsMissing(A2) Then A2 = aRefXL.aCl(2, 1).Value
      If Not IsMissing(A3) Then A3 = aRefXL.aCl(3, 1).Value
      If Not IsMissing(A4) Then A4 = aRefXL.aCl(4, 1).Value
      If Not IsMissing(A5) Then A5 = aRefXL.aCl(5, 1).Value
      If Not IsMissing(a6) Then a6 = aRefXL.aCl(6, 1).Value
   End If
   
   RefGetColV = bFail

End Function

Public Function RefGetRowV(sRef As String, _
                           Optional A1, Optional A2, Optional A3, _
                           Optional A4, Optional A5, Optional a6)

' Returns values at and right of sRef into the variables a1 to a6

   Dim aRefXL As tRefXL
   Dim bFail As Boolean

   aRefXL.sRef = sRef
   bFail = RefXLSetObj(aRefXL)
   If Not bFail Then
      If Not IsMissing(A1) Then A1 = aRefXL.aCl.Value
      If Not IsMissing(A2) Then A2 = aRefXL.aCl(1, 2).Value
      If Not IsMissing(A3) Then A3 = aRefXL.aCl(1, 3).Value
      If Not IsMissing(A4) Then A4 = aRefXL.aCl(1, 4).Value
      If Not IsMissing(A5) Then A5 = aRefXL.aCl(1, 5).Value
      If Not IsMissing(a6) Then a6 = aRefXL.aCl(1, 6).Value
   End If
   
   RefGetRowV = bFail

End Function

Public Function RefPutArrF(sRef As String, fArr() As Double) As Boolean

' Outputs an array of real to a reference in sRef
' Returns true if the reference is not valid

   Dim aRefXL As tRefXL
   Dim bFail As Boolean

   aRefXL.sRef = sRef
   bFail = RefXLSetObj(aRefXL)
   If Not bFail Then
      RngPutArrF aRefXL.aCl, fArr()
   End If

   RefPutArrF = bFail

End Function

Public Function RefPutArrL(sRef As String, lArr() As Long) As Boolean

' Outputs an array of real to a reference in sRef
' Returns true if the reference is not valid

   Dim aRefXL As tRefXL
   Dim bFail As Boolean

   aRefXL.sRef = sRef
   bFail = RefXLSetObj(aRefXL)
   If Not bFail Then
      RngPutArrL aRefXL.aCl, lArr()
   End If

   RefPutArrL = bFail

End Function

Public Function RefPutArrS(sRef As String, sArr() As String) As Boolean

' Outputs an array of strinfs to a reference in sRef
' Returns true if the reference is not valid

   Dim aRefXL As tRefXL
   Dim bFail As Boolean

   aRefXL.sRef = sRef
   bFail = RefXLSetObj(aRefXL)
   If Not bFail Then
      RngPutArrS aRefXL.aCl, sArr()
   End If

   RefPutArrS = bFail

End Function

Public Function RefPutArrV(sRef As String, vArr() As Variant) As Boolean

' Outputs an array of real to a reference in sRef
' Returns true if the reference is not valid

   Dim aRefXL As tRefXL
   Dim bFail As Boolean

   aRefXL.sRef = sRef
   bFail = RefXLSetObj(aRefXL)
   If Not bFail Then RngPutArrV aRefXL.aCl, vArr()
   
End Function

Public Function RefPutColV(sRef As String, _
                           Optional A1, Optional A2, Optional A3, _
                           Optional A4, Optional A5, Optional a6) As Boolean

' Outputs scalar values in cells at and below sRef

   Dim aRefXL As tRefXL
   Dim bFail As Boolean

   aRefXL.sRef = sRef
   bFail = RefXLSetObj(aRefXL)
   If Not bFail Then
      If Not IsMissing(A1) Then aRefXL.aCl.Value = A1
      If Not IsMissing(A2) Then aRefXL.aCl(2, 1).Value = A2
      If Not IsMissing(A3) Then aRefXL.aCl(3, 1).Value = A3
      If Not IsMissing(A4) Then aRefXL.aCl(4, 1).Value = A4
      If Not IsMissing(A5) Then aRefXL.aCl(5, 1).Value = A5
      If Not IsMissing(a6) Then aRefXL.aCl(6, 1).Value = a6
   End If

End Function

Public Function RefSetObj(sFP As String, sWb As String, _
                          sSh As String, sCl As String, _
                          aWb As Workbook, aSh As Worksheet, aCl As Range, _
                          Optional lWbDef As Long = MCl_LibXL_WbThis) As Boolean

' Sets Excel objects to names specified
' 1. If sWB is not specified
'    a) if aWb is nothing, "Thisworkbook" is used
'    b) else, the present workbook object is used
'    else
'    a) if aWb is nothing, the workbook is searched among the open workbooks
'       if a workbook with equal name is open, it is set, then c)
'    b) else, the aWb.Name and aWb.Path are compared to arguments
'       i) if aWb.Name is equal, the paths are compared and
'          if Name and Path match, the object is re-used
'       ii) if aWb.Name is different, the objectc is released
' 2. If sSh is not specified
'    a) is aSh is nothing, returns true (failure)
'    b) else, the present worksheet object is used
'    if sSh is specified
'    a) if aSh is allocated, compare name to argument
'       i)  if not aSh.Name = sSh, release aSh
'    b) if aSh is not allocated, search in workbook sheets
'       if not found, return true (failure)
' 3. If sCl is not specified
'    a) if aCl is nothing, remains nothing
'    b) else, the present range object is released
'       if the worksheet has changed

   Dim sFull As String
   Dim bFail As Boolean, bShNew As Boolean

   If sWb = "" Then                    ' --- 1. No workbook name given
      If aWb Is Nothing Then           '     Default ThisWorkbook
         If lWbDef = MCl_LibXL_WbThis Then Set aWb = ThisWorkbook
      Else                             '     Don't keep present workbook
         If lWbDef = MCl_LibXL_WbThis Then
            If Not StrComp(aWb.Name, ThisWorkbook.Name, vbTextCompare) = 0 Then
               Set aWb = Nothing
            End If
         Else
            If Not lWbDef = MCl_LibXL_WbKeep Then Set aWb = Nothing
         End If
      End If
   Else
      If Not aWb Is Nothing Then       ' --- Present workbook -> check
         If StrComp(aWb.Name, sWb, vbTextCompare) = 0 Then
            If Not sFP = "" Then       '     b) Path specified -> check
               If Not StrComp(aWb.Path, sFP, vbTextCompare) = 0 Then
                  Set aWb = Nothing    '     a) no path match, release
               End If
            End If
         Else
            Set aWb = Nothing          '     a) no name match, release
         End If
      End If
      If aWb Is Nothing Then           ' --- Workbook to be found
         For Each aWb In Application.Workbooks
            If StrComp(aWb.Name, sWb, vbTextCompare) = 0 Then Exit For
         Next aWb
         If Not aWb Is Nothing Then    ' --- Open workbook found
            If Not sFP = "" Then       '     Path aspecified -> check
               If Not StrComp(aWb.Path, sFP, vbTextCompare) = 0 Then
                  aWb.Close False      '     path mismatch, close
                  Set aWb = Nothing    '     and release
               End If
            End If
         End If
      End If
      If aWb Is Nothing Then           ' --- No present workbook yet
         sFull = QRS_LibDOS.PathFile(sFP, sWb)
         bFail = Not QRS_LibDOS.FileExists(sFull)
         If Not bFail Then
            Set aWb = Application.Workbooks.Open(sFull)
         End If
      End If
   End If
   If bFail Then GoTo RefSetObj_Ende

   If sSh = "" Then                    ' --- 2. sSh not specified
      bFail = aSh Is Nothing           '     a) no present sheet -> failure
   Else                                '        sSh specified
      If Not aSh Is Nothing Then       '     b) present sheet, check name
         If StrComp(aSh.Name, sSh, vbTextCompare) Then Set aSh = Nothing
      End If                           '        release if no match
      If aSh Is Nothing Then           '     b) no present sheet -> Set
         For Each aSh In aWb.Worksheets
            If StrComp(aSh.Name, sSh, vbTextCompare) = 0 Then Exit For
         Next aSh
         bFail = aSh Is Nothing
         bShNew = Not bFail
      End If
   End If
   If bFail Then GoTo RefSetObj_Ende

   If sCl = "" Then                    ' --- New sheet, present range->release
      If bShNew And Not aCl Is Nothing Then Set aCl = Nothing
   Else
      If Not aCl Is Nothing Then
         If Not StrComp(RngGetAddR(aCl), sCl, vbTextCompare) = 0 Then
            Set aCl = Nothing
         End If
      End If
      If aCl Is Nothing Then Set aCl = aSh.Range(sCl)
   End If

RefSetObj_Ende:

   RefSetObj = bFail

End Function

Public Function RefStrOff(sRef As String, _
                          Optional lOffRow As Long = 0, _
                          Optional lOffCol As Long = 0) As String

' Returns a reference string specified by offsets to sRef
' The reference must refer to an existing range

   Dim aRefXL As tRefXL
   Dim lRow As Long, lCol As Long
   
   With aRefXL
      .sRef = sRef
      RefXtr_Ele .sRef, .sFP, .sWb, .sSh, .sCl
      TxtRCRowCol .sCl, lRow, lCol
      lRow = lRow + lOffRow
      lCol = lCol + lOffCol
      .sCl = Col1232ABC(lCol) & CStr(lRow)
      RefFromEle .sRef, .sFP, .sWb, .sSh, .sCl
      RefStrOff = .sRef
   End With

End Function

Public Sub RefGetTxt(aWb As Workbook, aSh As Worksheet, aCl As Range, _
                     sFP As String, sWb As String, _
                     sSh As String, sCl As String)

   If Not aCl Is Nothing Then sCl = aCl.Address
   If Not aSh Is Nothing Then sSh = aSh.Name
   If Not aWb Is Nothing Then
      sWb = aWb.Name
      sFP = aWb.Path
   End If

End Sub

Public Sub RefXLGetRef(aRefXL As tRefXL)

' Get RefXL full reference string from Excel reference elements
' in a tRefXL data structure

   With aRefXL
      RefFromEle .sRef, .sFP, .sWb, .sSh, .sCl
   End With

End Sub

Public Sub RefXLGetTxt(aRefXL As tRefXL)

' Get RefXL object names from RefXL objects

   With aRefXL
      RefGetTxt .aWb, .aSh, .aCl, .sFP, .sWb, .sSh, .sCl
   End With

End Sub

Public Sub RefXLXtrEle(aRefXL As tRefXL)

' Get Excel Reference elements from reference string
' in a tRefXL data structure

   With aRefXL
      RefXtr_Ele .sRef, .sFP, .sWb, .sSh, .sCl
   End With

End Sub

Public Function RefXLSetObj(aRefXL As tRefXL) As Boolean

' Sets the objects in the RefXL data structure
' If sRef is not empty, extracts the elements

   With aRefXL
      If Not .sRef = "" Then           ' --- sRef given -> Extract elements
         RefXtr_Ele .sRef, .sFP, .sWb, .sSh, .sCl
      End If
      If Not (.sWb = "" And .sSh = "" And .sCl = "") Then
         RefXLSetObj = RefSetObj(.sFP, .sWb, .sSh, .sCl, .aWb, .aSh, .aCl)
      End If
   End With

End Function

Public Function RefXLSetOff(aRefXL As tRefXL, _
                            Optional lOffRow As Long = 0, _
                            Optional lOffCol As Long = 0) As Boolean

' Applies the row and column offset to the reference range, if defined

   Dim bFail As Boolean

   bFail = RefXLSetObj(aRefXL)         ' --- Set original reference
   If Not bFail Then                   ' --- Success
      With aRefXL
         If Not .aCl Is Nothing Then   '     Range specified: Offset
            Set .aCl = .aCl.Offset(lOffRow, lOffCol)
         End If
      End With
   End If

   RefXLSetOff = bFail

End Function

Public Sub RefXtr_Ele(sRef As String, _
                      Optional sFP As String = "", Optional sWb As String = "", _
                      Optional sSh As String = "", Optional sCl As String = "")

' Get Excel reference parts from reference string
' Rules:
' A) When sRef begins with single quote, assume Workbook
'    single quoted part ends at last single quote character
'    indicate cases
'        'Path\[Workbook]Worksheet'...          -> Set bSQ
' B) if bSQ: If next character is opening square bracket or
'    if not bSQ, if first character is opening square bracket
'    indicate cases
'        '[Workbook]... or [Workbook]...        -> Set bWB
' C) If not bWB, then either case
'        Path\[Workbook]  or
'        'Sheet'! or Sheet!
'    If \[ exists, indicates case
'        Path\[Workbook]                        -> Set bFP AND bWB
'    else indicates presence of simple reference
' D) If bFP, get path from L1 to L2, set L1=L2
' E) If bWB, find next ] from L1, extract WB from L1 to L2, Set L1=L2
' F) if bSQ, find next '!, if not bSQ, find next !
'    if present, extract Sh from L1 to l2, Set L1=L2
'    if absent: if bWB, then indicates case
'        [Workbook]Sheet
'    if not bWB, then indicate case
'        Range
' Tested the 06.03.2018 on the following cases:
'    sRef = "A3"                        ' OK
'    sRef = "Blatt!A3"                  ' OK
'    sRef = "[Book]Blatt!B4"            ' OK
'    sRef = "'[Book]Blatt'!C5"          ' OK
'    sRef = "[Book]Blatt"               ' OK
'    sRef = "'[Book]Blatt!'"            ' OK
'    sRef = "'Path\[Book]Blatt'!D6"     ' OK

   Const CsSQ As String = "'"
   Const CsQO As String = "["
   Const CsQC As String = "]"
   Const CsXM As String = "!"

   Dim sPS As String, sBB As String
   Dim lSQ As Long, lQO As Long, lQC As Long, lXM As Long, lPS As Long
   Dim lP1 As Long, lP2 As Long, lRS As Long
   Dim bSQ As Boolean, bFP As Boolean, bWb As Boolean

   sPS = Application.PathSeparator
   sBB = sPS & CsQO                    '     Workbook 1st delimiter if bFP

   lSQ = Len(CsSQ)
   lQO = Len(CsQO)
   lQC = Len(CsQC)
   lXM = Len(CsXM)
   lPS = Len(sPS)

   lRS = Len(sRef)
   lP1 = 1
   bSQ = Left(sRef, lSQ) = CsSQ        ' --- Single-quoted?
   If bSQ Then lP1 = lP1 + lSQ         ' --- Increment past quote
   bWb = (Mid(sRef, lP1, lQO) = CsQO)
   If bWb Then
      lP1 = lP1 + lQO
   Else
      lP2 = QRS_LibStr.StrInN(sRef, sBB, -1, lP1)
      bWb = lP2 > 0
      bFP = bWb                        '     Workbook after file path
   End If
                                       ' --- File path present
   If bFP Then
      sFP = Mid(sRef, lP1, lP2 - lP1)
      lP1 = lP2 + lPS + lQO            ' --- Increment past delimiter
   End If
   If bWb Then                         ' --- Workbook present
      lP2 = InStr(lP1, sRef, CsQC)     '     Next closing delimiter
      If lP2 = 0 Then Exit Sub         '     Not present -> bailout
      sWb = Mid(sRef, lP1, lP2 - lP1)
      lP1 = lP2 + lQC                  ' --- Increment past delimiter
   End If
   If bSQ Then sBB = CsSQ & CsXM Else sBB = CsXM
   lP2 = InStr(lP1, sRef, sBB)
   If lP2 = 0 Then                     ' --- No sheet name end delimiter
      If bWb Then                      ' --- Contains workbook?
         If bSQ Then                   '     Case '[Workbook]Sheet'
            sSh = Mid(sRef, lP1, lRS - (lSQ + lP1))
         Else                          '     Case [Workbook]Sheet
            sSh = Mid(sRef, lP1)
         End If
      Else
         sCl = Mid(sRef, lP1)          '     Case Range only
      End If
   Else
      sSh = Mid(sRef, lP1, lP2 - lP1)
      lP1 = lP2 + Len(sBB)
      sCl = Mid(sRef, lP1)
   End If

End Sub

Public Function RngAreSame(aRngA As Range, aRngB As Range) As Boolean

' Returns true if range aRngA refers to same range, sheet and book as aRngB

   Const CsAbs As String = "$"

   Dim aRefA As tRefXL, aRefB As tRefXL
   Dim bMatch

   RngGetRefXL aRngA, aRefA: RefXLGetRef aRefA
   RngGetRefXL aRngB, aRefB: RefXLGetRef aRefB
   bMatch = aRefA.sRef = aRefB.sRef

   RngAreSame = bMatch

End Function

Public Sub RngClrVal(aCl As Range, _
                     Optional lNRow As Long = 0, Optional lNCol As Long = 0)

' Clears the specified range
' if lNRow or lNCol > 0, clears that many rows or columns
'                        from the top left cell in aCl
' if lNRow or lNCol < 0, clears that many rows or columns
'                        towards the top left of aCl

   Const Cl01 As Long = 1

   Dim aClTL As Range
   Dim lCntR As Long, lCntC As Long    ' --- Range size
   Dim lOffR As Long, lOffC As Long    ' --- Top left offset

   lCntR = aCl.Rows.Count
   lCntC = aCl.Columns.Count
   If lNRow < 0 Then lOffR = lNRow     ' --- Top left offset
   If lNCol < 0 Then lOffC = lNCol
   If lNRow = 0 Then lNRow = lCntR     ' --- Use range size
   If lNCol = 0 Then lNCol = lCntC
   Set aClTL = aCl.Cells(1, 1).Offset(lOffR, lOffC)
   lOffR = Abs(lNRow) - Cl01           ' --- Offsets are 1 less
   lOffC = Abs(lNCol) - Cl01           '     than output size
   Range(aClTL, aClTL.Offset(lOffR, lOffC)).ClearContents

End Sub

Public Sub RngColorBG(aCl As Range, lColor() As Long, _
                    bNdx As Boolean, bUnicolor As Boolean)

' Sets the cell colors for all cells in the range covered by lColor
' if bNdx is set, interprets lColor values as color indices
' if bUnicolor is set, uses lColor(1,1) for all cells

   Dim aClTL As Range
   Dim lNRow As Long, lNCol As Long
   Dim lIRow As Long, lICol As Long

   QRS_LibArr.ArrBoundL lColor(), , lNRow, , lNCol
   Set aClTL = aCl.Cells(1, 1)
   If bUnicolor Then
      With Range(aClTL, aClTL.Offset(lNRow - 1, lNCol - 1))
         If bNdx Then
            .Interior.ColorIndex = lColor(1, 1)
         Else
            .Interior.Color = lColor(1, 1)
         End If
      End With
   Else
      If bNdx Then
         For lIRow = 1 To lNRow
            For lICol = 1 To lNCol
               With aClTL.Cells(lIRow, lICol).Interior
                  .ColorIndex = lColor(lIRow, lICol)
               End With
            Next lICol
         Next lIRow
      Else
         For lIRow = 1 To lNRow
            For lICol = 1 To lNCol
               With aClTL.Cells(lIRow, lICol).Interior
                  .ColorIndex = lColor(lIRow, lICol)
               End With
            Next lICol
         Next lIRow
      End If
   End If

End Sub

Public Sub RngFill(aCl As Range, vVal, _
                   Optional lNRow As Long = 0, Optional lNCol As Long = 0)

' Fills the specified range with vVal
' if lNRow or lNCol > 0, clears that many rows or columns
'                        from the top left cell in aCl
' if lNRow or lNCol < 0, clears that many rows or columns
'                        towards the top left of aCl

   Const Cl01 As Long = 1

   Dim aClTL As Range
   Dim lCntR As Long, lCntC As Long    ' --- Range size
   Dim lOffR As Long, lOffC As Long    ' --- Top left offset

   lCntR = aCl.Rows.Count
   lCntC = aCl.Columns.Count
   If lNRow < 0 Then lOffR = lNRow     ' --- Top left offset
   If lNCol < 0 Then lOffC = lNCol
   If lNRow = 0 Then lNRow = lCntR     ' --- Use range size
   If lNCol = 0 Then lNCol = lCntC
   Set aClTL = aCl.Cells(1, 1).Offset(lOffR, lOffC)
   lOffR = Abs(lNRow) - Cl01           ' --- Offsets are 1 less
   lOffC = Abs(lNCol) - Cl01           '     than output size
   Range(aClTL, aClTL.Offset(lOffR, lOffC)).Value = vVal

End Sub

Public Function RngGetAddR(aCl As Range) As String

' Returns relative range addresses by removing $ characters

   RngGetAddR = QRS_LibStr.StrRmv(aCl.Address, "$")

End Function

Public Sub RngGetRefXL(aCl As Range, aRefXL As tRefXL)

' Separates reference elements of a cell into range, worksheet and workbook

   With aRefXL
      Set .aCl = aCl
      Set .aSh = aCl.Parent
      Set .aWb = .aSh.Parent
      RefGetTxt .aWb, .aSh, .aCl, .sFP, .sWb, .sSh, .sCl
   End With

End Sub

Public Sub RngGetTxt(aCl As Range, _
                     Optional sCl As String = "", _
                     Optional sSh As String = "", _
                     Optional sWb As String = "", _
                     Optional sFP As String = "")

' Returns different elements of the full cell reference

   Dim aRefXL As tRefXL
   
   RngGetRefXL aCl, aRefXL
   RefXLGetTxt aRefXL

   With aRefXL
      sCl = .sCl
      sSh = .sSh
      sWb = .sWb
      sFP = .sFP
   End With

End Sub

Public Sub RngPutArrF(aCl As Range, fArr() As Double)

' Output fArr to the range having aCl as top left cell
' 19.03.2018 -> The world will never be as before:
' Value assignment to Range accepts array of type double

   Dim lRow1 As Long, lCol1 As Long
   Dim lOffR As Long, lOffC As Long

   QRS_LibArr.ArrBoundF fArr(), lRow1, lOffR, lCol1, lOffC
   lOffR = lOffR - lRow1
   lOffC = lOffC - lCol1
   If lOffR = 0 And lOffC = 0 Then
      aCl.Value = fArr(lRow1, lCol1)
   Else
      Range(aCl, aCl.Offset(lOffR, lOffC)).Value = fArr()
   End If

End Sub

Public Sub RngPutArrL(aCl As Range, lArr() As Long)

' Output lArr to the range having aCl as top left cell
' 19.03.2018 -> The world will never be as before:
' Value assignment to Range accepts array of type long

   Dim lRow1 As Long, lCol1 As Long
   Dim lOffR As Long, lOffC As Long

   QRS_LibArr.ArrBoundL lArr(), lRow1, lOffR, lCol1, lOffC
   lOffR = lOffR - lRow1
   lOffC = lOffC - lCol1
   If lOffR = 0 And lOffC = 0 Then
      aCl.Value = lArr(lRow1, lCol1)
   Else
      Range(aCl, aCl.Offset(lOffR, lOffC)).Value = lArr()
   End If

End Sub

Public Sub RngPutArrS(aCl As Range, sArr() As String)

' Output sArr to the range having aCl as top left cell

   Dim lRow1 As Long, lCol1 As Long
   Dim lOffR As Long, lOffC As Long

   QRS_LibArr.ArrBoundS sArr(), lRow1, lOffR, lCol1, lOffC
   lOffR = lOffR - lRow1
   lOffC = lOffC - lCol1
   If lOffR = 0 And lOffC = 0 Then
      aCl.Value = sArr(lRow1, lCol1)
   Else
      Range(aCl, aCl.Offset(lOffR, lOffC)).Value = sArr()
   End If

End Sub

Public Sub RngPutArrV(aCl As Range, vArr() As Variant)

' Output vArr to the range having aCl as top left cell

   Dim lRow1 As Long, lCol1 As Long
   Dim lOffR As Long, lOffC As Long

   QRS_LibArr.ArrBoundV vArr(), lRow1, lOffR, lCol1, lOffC
   lOffR = lOffR - lRow1
   lOffC = lOffC - lCol1
   If lOffR = 0 And lOffC = 0 Then
      aCl.Value = vArr(lRow1, lCol1)
   Else
      Range(aCl, aCl.Offset(lOffR, lOffC)).Value = vArr()
   End If

End Sub

Public Function GetXlLngID(Optional sLng As String = "", _
                           Optional sLoc As String = "", _
                           Optional sTag As String = "") As Long

' Returns the language ID of the current Excel and the language
'               Language Location       LangID   Tag      Win   Code

   Const sNL = "Dutch   ;              ;0x0013  ;nl        ;  19"
   Const sEN = "English ;              ;0x0009  ;en        ;   9"
   Const sUS = "English ;United States ;0x0409  ;en-US     ;1033"
   Const sUK = "English ;United Kingdom;0x0809  ;en-GB     ;2057"
   Const sOZ = "English ;Australia     ;0x0C09  ;en-AU     ;3081"
   Const sFR = "French  ;              ;0x000C  ;fr        ;  12"
   Const sBe = "French  ;Belgium       ;0x080C  ;fr-BE     ;2060"
   Const sFC = "French  ;Switzerland   ;0x100C  ;fr-CH     ;4108"
   Const sDE = "German  ;              ;0x0007  ;de        ;   7"
   Const sGE = "German  ;Germany       ;0x0407  ;de-DE     ;1031"
   Const sGC = "German  ;Switzerland   ;0x0807  ;de-CH     ;2055"
   Const sGA = "German  ;Austria       ;0x0C07  ;de-AT     ;3079"
   Const sGL = "German  ;Liechtenstein ;0x1407  ;de-LI     ;5127"
   Const sIT = "Italian ;              ;0x0010  ;it        ;  16"
   Const sII = "Italian ;Italy         ;0x0410  ;it-IT     ;1040"
   Const sIC = "Italian ;Switzerland   ;0x0810  ;it-CH     ;2064"
   Const sES = "Spanish ;              ;0x000A  ;es        ;  10"
   Const sSP = "Spanish ;Spain         ;0x040A  ;es-ES_trad;1034"
   Const sSX = "Spanish ;Mexico        ;0x080A  ;es-MX     ;2058"

' Excel has different language settings for different components or modes
' The component for which the language settings is requested is specified
' by selecting the msoLanguageID argument of the LanguageID method
' However, the language ID is read-only

   Dim sInf As String, sLst() As String
   Dim lID As Long, lLI As Long
   Dim bT As Boolean

   lID = Application.LanguageSettings.LanguageID(msoLanguageIDExeMode)

   Select Case lID
   Case 7                              ' --- DE
      sInf = sDE
   Case 9                              ' --- EN
      sInf = sEN
   Case 10                             ' --- ES
      sInf = sES
   Case 12                             ' --- FR
      sInf = sFR
   Case 16                             ' --- IT
      sInf = sIT
   Case 19                             ' --- NL
      sInf = sNL
   Case 1031                           ' --- DE-DE
      sInf = sGE
   Case 1033                           ' --- EN-US
      sInf = sUS
   Case 1034                           ' --- ES-SP
      sInf = sSP
   Case 1040                           ' --- IT-IT
      sInf = sII
   Case 2055                           ' --- DE-CH
      sInf = sGC
   Case 2057                           ' --- EN-UK
      sInf = sUK
   Case 2058                           ' --- ES-MX
      sInf = sSX
   Case 2060                           ' --- FR-BE
      sInf = sBe
   Case 2064                           ' --- IT-CH
      sInf = sIC
   Case 3079                           ' --- DE-AT
      sInf = sGA
   Case 3081                           ' --- EN-AU
      sInf = sOZ
   Case 5127                           ' --- DE-LX
      sInf = sGL
   End Select
   If Not sInf = "" Then
      bT = True
      QRS_LibStr.LstStrLst sInf, sLst()
      QRS_LibLst.LstXtrC5S sLst(), 1, sLng, sLoc, , sTag, lLI, bT
   End If

End Function

Public Function TxtRCRowCol(sRng As String, _
                            Optional lRow As Long, Optional lCol As Long, _
                            Optional sRow As String, Optional sCol As String)

' Returns the row and column of the range indicated with column name
' Checks the column with by character codes

   Dim sRel As String
   Dim lL As Long

   sRel = QRS_LibStr.StrRpl(sRng, "$", "")
   For lL = 1 To Len(sRel)
      If Asc(Mid(sRel, lL, 1)) < 64 Then Exit For
   Next lL

   sRow = Mid(sRel, lL)
   sCol = Left(sRel, lL - 1)

   lCol = ColABC2123(sCol)
   lRow = CLng(sRow)

End Function

Public Function ColABC2123(sColABC As String) As Long

' Returns the column number from its alpha label
' The alpha label is a 26-based value
   
   Const Cl01 As Long = 1

   Dim lL As Long, lP As Long, lD As Long, lC As Long

   lD = Cl01                           ' --- Least digit significance

   lL = Len(sColABC)                   ' --- Input string length
   For lP = lL To Cl01 Step -1         ' --- Least to most significant
      lC = lC + ChrABC2123(Mid(sColABC, lP, Cl01)) * lD
      lD = lD * MCl_LibXL_Chr          ' --- Increment digit significance
   Next lP

   ColABC2123 = lC

End Function

Public Function Col1232ABC(lCol123 As Long) As String

' Returns the column alpha label from its number

   Const Cl01 As Long = 1

   Dim sABC As String

   Dim lC As Long, lD As Long

   lD = lCol123                        ' --- Work on copy
   lC = lD
   While lD > MCl_LibXL_Chr
      QRS_Lib0.DivModLon lD - Cl01, MCl_LibXL_Chr, lD, lC
      sABC = Chr1232ABC(lC + Cl01) & sABC
      lC = lD
   Wend
   sABC = Chr1232ABC(lC) & sABC

   Col1232ABC = sABC

End Function

Public Function ChrABC2123(sABC As String) As Long

' Returns 1 for "A" to 26 for "Z"
' These values correspond to column numbers

   ChrABC2123 = Asc(Left(UCase(sABC), 1)) - MCl_LibXL_ChA

End Function

Public Function Chr1232ABC(l123 As Long) As String

' Returns "A" for 1 to "Z" for 26
' These values correspont to columnn labels

   Const Cl01 As Long = 1

   Chr1232ABC = Chr((l123 - Cl01) Mod MCl_LibXL_Chr + Cl01 + MCl_LibXL_ChA)

End Function
