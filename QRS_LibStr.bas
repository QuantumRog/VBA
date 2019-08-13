Attribute VB_Name = "QRS_LibStr"
Option Explicit

' Module : QRS_LibStr
' Project: any
' Purpose: Some very basic string utility VBA routines
'          Including the new masterpiece StrFmt date formatting function
' By     : QRS, Roger Strebel
' Date   : 21.01.2018
'          04.03.2018                  StrPad added and tested
'          13.03.2018                  StrRpl improved, StrRmv added
'          19.03.2018                  StrFld, StrFmt and StrHexLon added
'          30.03.2018                  StrBinLon added
'          24.04.2018                  StrPartN tested, works OK
'          18.06.2018                  ArrStrStr added, tested, works OK
'          15.08.2018                  StrKwdKey added
'          17.08.2018                  StrSplit2 handles quoted strings
'          19.08.2018                  StrNexDel, StrOcc, LstStrLst do quotes
'          01.11.2018                  ArrStrLon added, tested, works OK
'          28.01.2019                  ArrStrArr added, tested, works OK
'          06.02.2019                  StrRpl bug fixed
'          07.02.2019                  ArrStrArr 1-row 2D case implemented
'          16.02.2019                  StrRpt added
' --- The public interface
'     ArrStrArr                        String to string array        07.02.2019
'     ArrStrLon                        String to long array or list  01.11.2018
'     ArrStrStr                        String array to string        18.06.2018
'     LstLstStr                        String list to string         19.02.2018
'     LstStrLst                        String to string list         19.08.2018
'     StrBinLon                        Binary string to decimal long 30.03.2018
'     StrFld                           String delimited field        19.03.2018
'     StrFmt                           String format dates           19.03.2018
'     StrHexLon                        Hex string to decimal long    19.03.2018
'     StrInN                           N-th occurrence in string     21.01.2018
'     StrInR                           InStr from the right          21.01.2018
'     StrKwdKey                        Replace all Kwd by Key        15.08.2018
'     StrNexDel                        Next delimiter with text      19.08.2018
'     StrOcc                           Occurrence count out of text  19.08.2018
'     StrPad                           Pad to specific width (LRC)   04.03.2018
'     StrPartN                         N-th part of a quoted string  24.04.2018
'     StrRev                           Reverse string                21.01.2018
'     StrRmv                           Remove partial strings        13.03.2018
'     StrRpt                           Repeat string                 16.02.2019
'     StrRpl                           Replace in string             06.02.2019
'     StrSplit2                        Split string in 2 parts       18.02.2018
' --- The private sphere

Public Function ArrStrArr(sStr As String, sArr() As String, _
                          Optional sDelRow As String = "[", _
                          Optional sDelCol As String = ";", _
                          Optional sTxtBeg As String = "", _
                          Optional sTxtEnd As String = "", _
                          Optional bTight As Boolean = True) As Boolean

' Generate a string array from a string containing row and col delimiters
' Allows to produce a one-row 2D array if sStr starts with a row delimiter
' if bTight is not set and sArr already exists, the size is preserved
' The number of columns is determined in the first row

   Dim lIRow As Long, lNRow As Long, lKRow As Long
   Dim lICol As Long, lNCol As Long
   Dim sLst() As String, sRow() As String, sEle As String
   Dim b2D As Boolean, bD1 As Boolean, bOK As Boolean

   lIRow = InStr(1, sStr, sDelRow)
   b2D = lIRow > 0                     ' --- 2D output required
   bD1 = lIRow = 1                     ' --- Leading row delimiter

   If b2D Then                         ' --- 2D processing
      lIRow = 1
      If bD1 Then                      '     Split into rows - 1 row
         sEle = Mid(sStr, Len(sDelRow) + 1)
         LstStrLst sEle, sLst(), sDelRow
      Else                             '     Split into rows - n rows
         LstStrLst sStr, sLst(), sDelRow
      End If
      lNRow = UBound(sLst())
      LstStrLst sLst(lIRow), sRow(), sDelCol
      lNCol = UBound(sRow())           '     Use 1st row to count columns
      QRS_LibArr.ArrAllocS sArr(), lNRow + 1 - lIRow, lNCol

      bOK = True
      lKRow = 1
      While bOK And Not lIRow > lNRow
         For lICol = 1 To lNCol
            sArr(lKRow, lICol) = sRow(lICol)
         Next lICol
         lIRow = lIRow + 1
         lKRow = lKRow + 1
         If Not lIRow > lNRow Then
            LstStrLst sLst(lIRow), sRow(), sDelCol
         End If
      Wend
   Else
      LstStrLst sStr, sRow(), sDelCol
      lNCol = UBound(sRow())
      If lNCol > 0 Then
         QRS_LibLst.LstAllocS sArr(), lNCol
         For lICol = 1 To lNCol
            sArr(lICol) = sRow(lICol)
         Next lICol
      End If
   End If

   ArrStrArr = Not bOK

End Function

Public Function ArrStrLon(sStr As String, lArr() As Long, _
                          Optional sDelRow As String = "/", _
                          Optional sDelCol As String = ",") As Boolean

' Generate a list or an array of long
' For lists the input string sStr is of the form "L1;L2;...;LN"
' For tables the input string sStr if of the form
'     >R1C1;R1C2;...;R1CN>R2C1;R2C2;...;R2CN>...>RMC1;RMC2;...;RMCN"
' Returns true if parsing failed (non-numeric elements)

   Dim lIRow As Long, lNRow As Long, lKRow As Long
   Dim lICol As Long, lNCol As Long
   Dim sLst() As String, sRow() As String, sEle As String
   Dim b2D As Boolean, bD1 As Boolean, bOK As Boolean

   lIRow = InStr(1, sStr, sDelRow)
   b2D = lIRow > 0                     ' --- 2D output required
   bD1 = lIRow = 1                     ' --- Leading row delimiter

   If b2D Then                         ' --- 2D processing
      LstStrLst sStr, sLst(), sDelRow  '     Split into rows
      If bD1 Then lIRow = 2 Else lIRow = 1
      lNRow = UBound(sLst())
      LstStrLst sLst(lIRow), sRow(), sDelCol
      lNCol = UBound(sRow())           '     Use 1st row to count columns
      QRS_LibArr.ArrAllocL lArr(), lNRow + 1 - lIRow, lNCol

      bOK = True
      lKRow = 1
      While bOK And Not lIRow > lNRow
         For lICol = 1 To lNCol
            sEle = sRow(lICol)
            bOK = IsNumeric(sEle)
            If bOK Then
               lArr(lKRow, lICol) = CLng(sEle)
            Else
               Exit For
            End If
         Next lICol
         lIRow = lIRow + 1
         lKRow = lKRow + 1
         If Not lIRow > lNRow Then
            LstStrLst sLst(lIRow), sRow(), sDelCol
         End If
      Wend
   Else
      LstStrLst sStr, sRow(), sDelCol
      lNCol = UBound(sRow())
      If lNCol > 0 Then
         QRS_LibLst.LstAllocL lArr(), lNCol
         For lICol = 1 To lNCol
            sEle = sRow(lICol)
            bOK = IsNumeric(sEle)
            If bOK Then
               lArr(lICol) = CLng(sEle)
            Else
               Exit For
            End If
         Next lICol
      End If
   End If

   ArrStrLon = Not bOK

End Function

Public Function ArrStrStr(sArr() As String, lRow1 As Long, lRowL As Long, _
                          Optional sDelFld As String = ";", _
                          Optional sDelRow As String = vbNewLine, _
                          Optional bSkipEmpty As Boolean = False)

' Returns a string from the string array provided
' lRow1 and lRowL designate the first and last row included in the string
' (<0: from last row, =0: from first row)
' sDelFld specifies the field separation string
' sDelRow specifies the  row  separation string

   Const Cl01 As Long = 1, Cl02 As Long = 2

   Dim lNRow As Long, lNCol As Long
   Dim lXtr1 As Long, lXtrL As Long
   Dim lIRow As Long, lICol As Long
   Dim sStr As String, sRow As String

   QRS_LibArr.ArrAllocS sArr(), lNRow, lNCol
   lXtr1 = QRS_LibArr.NdxGetSX1(Cl01, lNRow, lRow1)
   lXtrL = QRS_LibArr.NdxGetSXL(Cl01, lNRow, lRowL)

   For lIRow = lXtr1 To lXtrL
      sRow = sArr(lIRow, Cl01)
      For lICol = Cl02 To lNCol
         sRow = sRow & sDelFld & sArr(lIRow, lICol)
      Next lICol
      If lIRow = lXtr1 Then sStr = sRow Else sStr = sStr & sDelRow & sRow
   Next lIRow

   ArrStrStr = sStr
                          End Function

Public Sub LstLstStr(sLst() As String, sStr As String, _
                     Optional sDel As String = ";")

' Generate a delimited string from a string list

   Dim lEI As Long, lE1 As Long, lEL As Long

   QRS_LibLst.LstBoundS sLst(), lE1, lEL

   sStr = sLst(lE1)
   lE1 = lE1 + 1
   For lEI = lE1 To lEL
      sStr = sStr & sDel & sLst(lEI)
   Next lEI

End Sub

Public Sub LstStrLst(sStr As String, sLst() As String, _
                     Optional sDel As String = ";", _
                     Optional sTxtBeg As String = "", _
                     Optional sTxtEnd As String = "")

' Generate a string list from a string with entry delimiters
' Delimiter characters enclosed in text qualifiers are ignored
' The use of the split()-function is not in order as the returned
' list is zero-based

   Dim lI As Long, lN As Long, lP As Long, lQ As Long
   Dim bLast As Boolean

   lN = StrOcc(sStr, sDel, sTxtBeg, sTxtEnd) + 1
   If lN > 0 Then
      QRS_LibLst.LstAllocS sLst(), lN
      lN = 0
      While Not bLast
         lN = lN + 1
         bLast = StrNexDel(sStr, sDel, sTxtBeg, sTxtEnd, lP, lQ)
         sLst(lN) = Mid(sStr, lP, lQ - lP)
         lP = lQ
      Wend
   End If

End Sub

Public Function StrBinLon(sStrIn As String) As Long

' Returns a long integer from a binary string
' Checks for "1" or "0". If other characters occur, returns zero
' If the input string contains  more  than 32 bits, returns zero

   Const Cl01 As Long = 1, Cl02 As Long = 2, ClM0 As Long = -2147483647
   Const CiA0 As Integer = 48, CiA1 As Integer = 49

   Dim lMsk As Long, lLen As Long, lBit As Long, l As Long
   Dim iA As Integer

   lLen = Len(sStrIn)                  ' --- Input string length
   If lLen > 32 Then GoTo Ende         '     Avoid overflow

   lBit = lLen
   lMsk = Cl01                         ' --- LSB
   iA = Asc(Mid(sStrIn, lBit, Cl01))   '     rightmost character
   If Not (iA = CiA1 Or iA = CiA0) Then GoTo Ende
   If iA = CiA1 Then l = l + lMsk

   For lBit = lLen - 1 To Cl02 Step -Cl01
      lMsk = lMsk * Cl02               ' --- All bits before MSB
      iA = Asc(Mid(sStrIn, lBit, Cl01)) '    characters to left
      If Not (iA = CiA1 Or iA = CiA0) Then GoTo Ende
      If iA = CiA1 Then l = l + lMsk   '     Set bit by addition
   Next lBit

   iA = Asc(Mid(sStrIn, lBit, Cl01))
   If Not (iA = CiA1 Or iA = CiA0) Then GoTo Ende
   If iA = CiA1 Then
      If lLen < 32 Then                ' --- Set bit by addition
         lMsk = lMsk * Cl02
         l = l + lMsk
      Else
         l = ClM0 + l - Cl01           ' --- Set MSB by inversion
      End If
   End If

Ende:

   StrBinLon = l

End Function

Public Function StrHexLon(sStrIn As String) As Long

' Returns a decimal value from a string containing a hexadecimal value
' Handles "0x" prefix and "h" suffix and lower case hex digits

   Const Cl01 As Long = 1
   Const Cl0x As Long = 16
   Const ClA0 As Long = -7             '     "A" to 10
   Const ClUC As Long = -33            '     "a" to "A"
   Const Cl00 As Long = -48            ' --- "0" to 0

   Dim sTmp As String
   Dim lP As Long, lQ As Long, lR As Long, lN As Long, lD As Long

   sTmp = sStrIn
   sTmp = StrRmv(sTmp, "0x")
   sTmp = StrRmv(sTmp, "h")

   lN = Len(sTmp)
   lQ = Cl01
   For lP = lN To 1 Step -1
      lD = CLng(Asc(Mid(sTmp, lP, 1))) + Cl00
      If lD > Cl0x Then lD = (lD And ClUC) + ClA0
      lR = lR + lD * lQ
      lQ = lQ * Cl0x
   Next lP

   StrHexLon = lR

End Function

Public Sub StrFld(sStrIn As String, sFldBeg As String, sFldEnd As String, _
                  lPos1 As Long, sStrFld As String, _
                  Optional lPosBeg As Long = 0, Optional lPosEnd As Long = 0)

' Extracts a field in a string from starting and ending delimiter characters
' Returns the field value in sStrFld and the field begin and end positions
' field limit positions are outside delimiters
' May be called several times with lPosEnd as argument value for lPos1
' to extract consecutive fields

   Const Cl01 As Long = 1

   Dim lPB As Long, lPE As Long        ' --- Delimiter positions
   Dim lLB As Long, lLE As Long        ' --- Delimiter lengths

   lLB = Len(sFldBeg)
   lLE = Len(sFldEnd)
   sStrFld = ""                        ' --- Avoid static remanence

   If lPos1 = 0 Then lPB = Cl01 Else lPB = lPos1
   lPB = InStr(lPB, sStrIn, sFldBeg, vbBinaryCompare)
   If lPB = 0 Then GoTo Ende

   lPE = lPB + Len(sFldBeg)
   lPE = InStr(lPE, sStrIn, sFldEnd, vbBinaryCompare)
   If lPE = 0 Then GoTo Ende

   sStrFld = Mid(sStrIn, lPB + lLB, lPE - (lPB + lLB))

Ende:

   lPosBeg = lPB
   lPosEnd = lPE + lLE

End Sub

Public Function StrFmt(sStrIn As String, dDate As Date) As String

' Detects formatting characters in a string and replaces them
' by the formatted elements of dDate
' The formatting characters are used as found to format string
' NOTE: month and weekday name formatting depends on locale settings

   Const CsFmtBeg As String = "<"
   Const CsFmtEnd As String = ">"

   Dim sTmp As String, sFmt As String, sDat As String
   Dim lPB As Long, lPE As Long, lLD As Long

   lLD = Len(CsFmtBeg) + Len(CsFmtEnd) ' --- Total delimiter length
   sTmp = sStrIn
   lPE = 1                             ' --- Extract first format field
   StrFld sTmp, CsFmtBeg, CsFmtEnd, lPE, sFmt, lPB, lPE
   While Not lPB = 0
      sDat = Format(dDate, sFmt)       ' --- Apply format and replace
      sTmp = Left(sTmp, lPB - 1) & sDat & Mid(sTmp, lPE)
                                       ' --- Shift by length difference
      lPE = lPE + Len(sDat) - (lLD + Len(sFmt))
      StrFld sTmp, CsFmtBeg, CsFmtEnd, lPE, sFmt, lPB, lPE
   Wend

   StrFmt = sTmp

End Function

Public Sub StrSplit2(sStrIn As String, sStrAt As String, _
                     Optional sStrL As String = "", _
                     Optional sStrR As String = "", _
                     Optional lFrom As Long = 1, _
                     Optional sTxtQ As String = "")

' Split a string in two parts, determined by sStrAt
' sStrL returns the part left  to sStrAt
' sStrR returns the part right to sStrAt
' If sStrAt is not contained, sStrL and sStrR contain all sStrIn

   Dim lP As Long, lL As Long, lQ As Long

   lL = Len(sStrAt)
   If lL = 0 Then Exit Sub

   If sTxtQ = "" Then
      lQ = 0                           ' --- No quotes specified
   Else
      lQ = InStr(lFrom, sStrIn, sTxtQ, vbTextCompare)
   End If

   If lQ = 0 Then                      ' --- Consider no quotes
      lP = InStr(lFrom, sStrIn, sStrAt, vbTextCompare)
      If lP = 0 Then                   ' --- sStrAt not contained
         sStrL = Mid(sStrIn, lFrom)    '     Return full string in
         sStrR = sStrL                 '     both sStrL and sStrR
      Else
         sStrL = Mid(sStrIn, lFrom, lP - lFrom)
         sStrR = Mid(sStrIn, lP + lL)
      End If
   Else
      lP = InStr(lFrom, sStrIn, sStrAt, vbTextCompare)
      If lP = 0 Then                   ' --- sStrAt not contained
         sStrL = Mid(sStrIn, lFrom)    '     Return full string in
         sStrR = sStrL                 '     both sStrL and sStrR
      Else
         If lP < lQ Then               ' --- Quote past sStrAt
            sStrL = Mid(sStrIn, lFrom, lP - lFrom)
            sStrR = Mid(sStrIn, lP + lL)
         Else                          ' --- Quote before sStrAt
            lQ = InStr(lQ + Len(sTxtQ), sStrIn, sTxtQ, vbTextCompare)
            If lQ = 0 Then             ' --- Unbalanced quotes
               sStrL = Mid(sStrIn, lFrom, lP - lFrom)
               sStrR = Mid(sStrIn, lP + lL)
            Else
               If lP < lQ Then         ' --- separator in quotes
                  lP = InStr(lQ + Len(sTxtQ), sStrIn, sStrAt, vbTextCompare)
                  If lP = 0 Then       ' --- No separator outside quotes
                     sStrL = Mid(sStrIn, lFrom)
                     sStrR = sStrL
                  Else                 ' --- separator past quotes
                     sStrL = Mid(sStrIn, lFrom, lP - lFrom)
                     sStrR = Mid(sStrIn, lP + lL)
                  End If
               Else
                     sStrL = Mid(sStrIn, lFrom, lP - lFrom)
                     sStrR = Mid(sStrIn, lP + lL)
               End If
            End If
         End If
      End If
   End If

End Sub

Public Function StrInN(sStrIn As String, sWhat As String, _
                       Optional lN As Long = 1, _
                       Optional lFrom As Long = 1) As Long

' Return the position of the lN-th occurrence of sWhat in sStrIn
' if lN < 0, searches the occurrences from right to left
' if sWhat occurs less than lN times in sStrIn, returns
' zero or a negative value indicating the occurrence count

   Const Cl1 As Long = 1, Cl2 As Long = 2

   Dim sStr1 As String, sStr2 As String
   Dim lI As Long, lM As Long, lP As Long, lLW As Long, lLS As Long

   lLW = Len(sWhat)
   If lLW = 0 Then
      lP = 0
   Else
      lP = 1
      If lN < 0 Then                      ' --- Search from the right
         lM = Abs(lN)                     '     Avoid returning altered value
         For lI = Cl1 To lM
            lP = StrInR(sStrIn, sWhat, lP)
         Next lI
      Else                                ' --- search from the left
         lP = InStr(lFrom, sStrIn, sWhat, vbTextCompare)
         If lP > 0 Then
            For lI = Cl2 To lN
               lP = InStr(lP + lLW, sStrIn, sWhat, vbTextCompare)
               If lP = 0 Then Exit For    '     Not found
            Next lI
            If Not lI > lN Then lP = Cl1 - lI ' Occurrence count
         End If
      End If
   End If

   StrInN = lP                         ' --- Return value

End Function

Public Function StrInR(sStrIn As String, sWhat As String, _
                       Optional lStart As Long = 1) As Long

' Return the position of sWhat in sStrIn, counted from the right
' if sWhat is longer than sStrIn, does not search
' lStart constrains the search, analogous to StrIn:
' 1: search from the right end
' 2: omit last character in sStrIn (will not find "BC" in "ABC")
' Returns 0 if sWhat is not contained
' Returns position of first character of sWhat in sStrIn

   Const Cl1 As Long = 1

   Dim lI As Long, lP As Long, lLW As Long, lLS As Long

   lLS = Len(sStrIn) + Cl1 - lStart
   lLW = Len(sWhat)

   If lLW > lLS Then                   ' --- sWhat longer than sStrIn
      lI = 0
   Else
      lP = lLS + Cl1 - lLW             ' --- First position to search
      For lI = lP To 1 Step -Cl1       ' --- Count down
         If StrComp(Mid(sStrIn, lI, lLW), sWhat) = 0 Then Exit For
      Next lI
   End If

   StrInR = lI

End Function

Public Function StrKwdKey(sStrTpl As String, sLstKwd() As String, _
       sLstKey() As String) As String

' Replaces all keywords in the template by key values

   Dim sStr As String
   Dim lE1 As Long, lEI As Long, lEL As Long

   QRS_LibLst.LstBoundS sLstKwd(), lE1, lEL
   sStr = StrRpl(sStrTpl, sLstKwd(lE1), sLstKey(lE1))
   For lEI = lE1 + 1 To lEL
      sStr = StrRpl(sStr, sLstKwd(lEI), sLstKey(lEI))
   Next lEI

   StrKwdKey = sStr

End Function

Public Function StrNexDel(sTxt As String, sDel As String, _
                          sBeg As String, sEnd As String, _
                          lPosPrev As Long, lPosNext As Long) As Boolean

' Finds the next occurence of the delimiter sDel outside sBeg and sEnd
' starting at lFrom and returns its position in lNext
' Can be called repetitively with lFrom as the previous lNext
' Returns lPosPrev and lPosNext so that the string contained can be
' extracted directly by using lPosPrev and lPosNext (as difference)
' Returns true if no more delimiter was found

   Const Cl01 As Long = 1
   Const CbT As Boolean = True

   Dim lLB As Long, lLE As Long, lLT As Long, lLD As Long
   Dim lPB As Long, lPE As Long

   lLB = Len(sBeg): lLE = Len(sEnd): lLT = Len(sTxt): lLD = Len(sDel)
                                       ' --- Search start position
   If lPosPrev = 0 Then lPosPrev = Cl01 Else lPosPrev = lPosPrev + lLD
   lPB = lPosPrev                      '     Copy for qualifier search

   lPosNext = InStr(lPosPrev, sTxt, sDel, vbTextCompare)
   If lPosNext = 0 Then                ' --- No delimiter found
      StrNexDel = CbT                  '     Last field
      lPosNext = lLT + Cl01
      Exit Function
   End If

   If sBeg = "" Or sEnd = "" Then      ' --- no text qualifiers specified
      Exit Function                    '     done, exit
   End If
                                       ' --- Look for text begin qualifier
   lPB = InStr(lPB, sTxt, sBeg, vbTextCompare)
   If lPB = 0 Then
      Exit Function                    '     None found, exit
   End If
                                       ' --- Delimiter past current field
   If lPB > lPosNext Then Exit Function

   lPE = lPB + lLB                     ' --- Look for text end qualifier
   lPE = InStr(lPE, sTxt, sEnd, vbTextCompare)
   If lPE = 0 Then Exit Function       '     none found, unbalanced text
   lPE = lPE + lLE
   If lPosNext < lPE Then              ' --- Current delimiter inside text
      lPosNext = InStr(lPE, sTxt, sDel, vbTextCompare)
      If lPosNext = 0 Then
         StrNexDel = CbT
         lPosNext = lLT + 1
      End If
   End If

End Function

Public Function StrOcc(sStrIn As String, sOcc As String, _
                       Optional sTxtBeg As String = "", _
                       Optional sTxtEnd As String = "") As Long

' Counts substring occurrences in a string when they are
' out of delimited parts. Handles the following cases:
'    No delimiters
'    Matching Beg-End pairs
'    Single End before first Beg is ignored
'    Single Beg  after last  End is ignored

   Const Cl01 As Long = 1

   Dim lPO As Long, lLO As Long        ' --- sOcc current position and length
   Dim lPB As Long, lPE As Long        ' --- Current text marker positions
   Dim lLB As Long, lLE As Long        ' --- Text marker lengths
   Dim lN As Long
   Dim bLast As Boolean

   lLB = Len(sTxtBeg)                  ' --- Text begin marker length
   lLE = Len(sTxtEnd)                  ' --- Text  end  marker length

   If lLB = 0 And lLE = 0 Then         ' --- No delimiters -> Simple
      lLO = Len(sOcc)                  '     Occurrence length
      lN = 0
      lPO = InStr(Cl01, sStrIn, sOcc)
      While lPO > 0
         lN = lN + Cl01
         lPO = InStr(lPO + lLO, sStrIn, sOcc)
      Wend
   Else
      While Not bLast                  ' --- Use StrNexDel
         lN = lN + Cl01
         bLast = StrNexDel(sStrIn, sOcc, sTxtBeg, sTxtEnd, lPB, lPE)
         lPB = lPE                     ' --- Shift parameters
      Wend
   End If

   StrOcc = lN

End Function

Public Function StrPartN(sStrIn As String, lPartN As Long, sDel As String, _
                         sTxtB As String, sTxtE As String) As String

' Returns the lPartN-th part of a string with delimiters.
' Delimiters enclosed by quotes are ignored.
' separate quote start and end strings may be specified
' if sTxtB contains a value and sTxtE is empty, assumes sTxtE = sTxtB
' If the mumber of parts present in sStrIn is less than lNPart
' then returns an empty string
' If no delimiter is contained in the string, returns the whole string
' This version does not support part search from the right
' (Last part or second last)

   Const Cl01 As Long = 1

   Dim lNP As Long
   Dim lP0 As Long, lPL As Long        ' --- Part  start and end position
   Dim lQB As Long, lQE As Long        ' --- Quote start and end positions
   Dim lPD As Long                     ' --- Delimiter position
   Dim lLB As Long, lLE As Long        ' --- Quote start and end lengths
   Dim lLD As Long, lLS As Long        ' --- Delimiter and string lengths
   Dim b As Boolean                    ' --- Done

   If sTxtE = "" And Not sTxtB = "" Then sTxtE = sTxtB

   lLB = Len(sTxtB)                    ' --- Quote  start   length
   lLE = Len(sTxtE)                    ' --- Quote   end    length
   lLD = Len(sDel)                     ' --- Part delimiter length
   lLS = Len(sStrIn) + Cl01            ' --- Input  string  length
   If lLB = 0 Then lQB = lLS           ' --- No quote: Put past end
   lP0 = Cl01 - lLD
   lPD = Cl01

   Do
                                       ' --- Next delimiter
      lPD = InStr(lPD, sStrIn, sDel, vbTextCompare)
      b = lPD = 0                      '     No more found
      If b Then lPD = lLS              '     Set past string end

      If lPD > lQE Then                ' --- Past current quote end
         lQB = InStr(lQE + lLE, sStrIn, sTxtB, vbTextCompare)
         If lQB = 0 Then               ' --- Find new quote start
            lQB = lLS                  '     No more found
            lQE = lQB                  '     Set past string end
         Else                          ' --- Find new quote end
            lQE = InStr(lQB + lLB, sStrIn, sTxtE, vbTextCompare)
            If lQE = 0 Then lQE = lLS  ' --- Unclosed quote-> Till end
         End If
      End If
      If lPD > lQE Or Not lPD > lQB Then   ' Delimiter is out of quote
         lNP = lNP + Cl01              '     Count
         lPL = lP0 + lLD               '     Shift part start
         lP0 = lPD                     '     Set   part end
         b = b Or lNP = lPartN         ' --- Part count as requested
      End If
      lPD = lPD + lLD
   Loop Until b

   If lNP = lPartN Then
      StrPartN = Mid(sStrIn, lPL, lP0 - lPL)
   End If

End Function

Public Function StrPad(sStrIn As String, lWdth As Long, _
                       Optional sAlign As String = "L", _
                       Optional sPad As String = " ") As String

' Pads or clips the input string to match the specified length
' sAlign controls the padding and clipping:
'   "L": Left-align, clip on the right
'   "R": Right-align, clip on the left
'   "C": Center, clip on both sides

   Dim sFill As String
   Dim lLS As Long, lPC As Long, lCP As Long

   lLS = Len(sStrIn)
   lPC = lWdth - lLS                   ' --- >0: Pad, <0: Clip
   If lPC > 0 Then lCP = Asc(Left(sPad, 1))

   Select Case UCase(Left(sAlign, 1))
   Case "L"                            ' --- Align left, clip on right
      sFill = Left(sStrIn, lWdth)
      If lPC > 0 Then sFill = sFill & String(lPC, lCP)
   Case "R"                            ' --- Align right, clip on left
      sFill = Right(sStrIn, lWdth)
      If lPC > 0 Then sFill = String(lPC, lCP) & sFill
   Case "C"                            ' --- Center, clip both ends
      lLS = Abs(lPC) / 2
      If lPC < 0 Then
         sFill = Mid(sStrIn, lLS + 1, lWdth)
      Else
         sFill = String(lLS, lCP) & sStrIn & String(lPC - lLS, lCP)
      End If
   Case Else                           ' --- No operation
      sFill = sStrIn
   End Select
   StrPad = sFill

End Function

Public Function StrRev(sStrIn As String) As String

' Reverse character order in sStrIn using the VBA StrReverse function

'   Const Cl1 As Long = 1
'
'   Dim sRet As String
'   Dim lI As Long, lL As Long
'
'   lL = Len(sStrIn)
'   For lI = Cl1 To lL
'      sRet = Mid(sStrIn, lI, Cl1) & sRet
'   Next lI

   StrRev = StrReverse(sStrIn)         ' --- Shoulda known it long ago

End Function

Public Function StrRmv(sStrIn As String, sRmv As String) As String

' Remove all instances of sRmv in sStrIn.
' Is a particular version of StrRpl

   Const Cl1 As Long = 1

   Dim lI As Long, lP As Long, lLR As Long
   Dim s As String

   s = sStrIn
   lLR = Len(sRmv)                     ' --- Search string length
   lP = InStr(Cl1, s, sRmv, vbTextCompare)
   While Not lP = 0
      s = Left(s, lP - Cl1) & Mid(s, lP + lLR)
      lP = InStr(lP, s, sRmv, vbTextCompare)
   Wend

   StrRmv = s

End Function

Public Function StrRpl(sStrIn As String, sRpl As String, sBy As String) As String

' Replace all instances of sRpl by sBy in sStrIn
' Skips inserted part

   Const Cl1 As Long = 1

   Dim lI As Long, lP As Long, lLB As Long, lLR As Long
   Dim s As String

   s = sStrIn
   lLB = Len(sBy)                      ' --- Replacement length
   lLR = Len(sRpl)                     ' --- Search string length
   lP = InStr(Cl1, s, sRpl, vbTextCompare)
   While Not lP = 0
      s = Left(s, lP - Cl1) & sBy & Mid(s, lP + lLR)
      lP = InStr(lP + lLB, s, sRpl, vbTextCompare)
   Wend

   StrRpl = s

End Function

Public Function StrRpt(sStrIn As String, lN As Long) As String

' Repeat sStrIn lN times
' The VBA String() function only repeats one character

   Dim lI As Long
   Dim s As String

   For lI = 1 To lN
      s = s & sStrIn
   Next lI

   StrRpt = s

End Function
