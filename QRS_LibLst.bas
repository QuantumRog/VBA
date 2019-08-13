Attribute VB_Name = "QRS_LibLst"
Option Explicit

' Module : QRS_LibLst
' Project: any
' Purpose: Some very basic typed 1D list utility VBA routines
' By     : QRS, Roger Strebel
' Date   : 19.02.2018
'          25.03.2018                  List reversion routines added
'          26.03.2018                  LstGet routines added
'          01.04.2018                  LstPut routines added and tested
'          21.07.2018                  Find routines added
'          25.07.2018                  Insertion routines added
'          07.02.2019                  LstSetC5V/R5V added, LstXtrC5S improved
' --- The public interface
'     LstIsAllD                        Is dates list allocated?      18.02.2018
'     LstIsAllF                        Is double list allocated?     18.02.2018
'     LstIsAllL                        Is long list allocated?       18.02.2018
'     LstIsAllS                        Is string list allocated?     18.02.2018
'     LstIsAllV                        Is variant list allocated?    18.02.2018
'     LstAllocD                        Allocate/sizeof dates list    18.02.2018
'     LstAllocF                        Allocate/sizeof double list   18.02.2018
'     LstAllocL                        Allocate/sizeof long list     18.02.2018
'     LstAllocS                        Allocate/sizeof strging list  18.02.2018
'     LstAllocV                        Allocate/sizeof variant list  18.02.2018
'     LstBoundD                        Bounds of dates list          18.02.2018
'     LstBoundF                        Bounds of double list         18.02.2018
'     LstBoundL                        Bounds of long list           18.02.2018
'     LstBoundS                        Bounds of string list         18.02.2018
'     LstBoundV                        Bounds of variant list        18.02.2018
'     LstFind_D
'     LstFind_F                        Find in unordered double list 07.02.2019
'     LstFind_L                        Find in unordered long list   17.07.2018
'     LstFind_S                        Find in unordered string list 21.07.2018
'     LstFind_V
'     LstGenerD                        Generate   dates  list        25.03.2018
'     LstGenerF                        Generate real value list      25.03.2018
'     LstGenerL                        Generate long integer list    25.03.2018
'     LstGet_FF                        Extract sublist of real       26.03.2018
'     LstGet_LL                        Extract sublist of long       26.03.2018
'     LstGet_SS                        Extract sublist of string     26.03.2018
'     LstInsDat                        Insert element to string list
'     LstInsDbl                        Insert element to string list
'     LstInsLon                        Insert element to string list 25.07.2018
'     LstInsStr                        Insert element to string list 25.07.2018
'     LstMergeF                        Merge two ordered real lists  25.03.2018
'     LstMergeL                        Merge two ordered long lists  25.03.2018
'     LstMergeS                        Merge 2 ordered string lists  25.03.2018
'     LstPut_FF                        Output list to list of real   01.04.2018
'     LstPut_LL                        Output list to list of long   01.04.2018
'     LstPut_SS                        Output list to list of string 01.04.2018
'     LstReverD                        Reverse list of date values   26.03.2018
'     LstReverF                        Reverse list of real values   25.03.2018
'     LstReverL                        Reverse list of long values   25.03.2018
'     LstReverS                        Reverse list of strings       25.03.2018
'     LstReverV                        Reverse list of variant       25.03.2018
'     LstSetC5V                        Set 5 continuous variants     07.02.2019
'     LstSetR5V                        Set 5 random variants         07.02.2019
'     LstXtrC5S                        Extract 5 continuous strings  07.02.2019

Public Function LstIsAllD(dLst() As Date) As Boolean

' Returns true if date list has been allocated

   LstIsAllD = Not Not dLst()

End Function

Public Function LstIsAllF(fLst() As Double) As Boolean

' Returns true if double list has been allocated

   LstIsAllF = Not Not fLst()

End Function

Public Function LstIsAllL(lLst() As Long) As Boolean

' Returns true if long list has been allocated

   LstIsAllL = Not Not lLst()

End Function

Public Function LstIsAllS(sLst() As String) As Boolean

' Returns true if string list has been allocated

   LstIsAllS = Not Not sLst()

End Function

Public Function LstIsAllV(vLst() As Variant) As Boolean

' Returns true if variant list has been allocated

   LstIsAllV = Not Not vLst()

End Function

Public Sub LstAllocD(dLst() As Date, _
                     Optional lNRow As Long = 0)

' Allocates or re-allocates list if necessary
' Allocation is necessary if unallocated
' Re-allocation is necessary if size does not match
' if lNRow=0 then returns list size

   Const Cl01 As Long = 1

   Dim lARow As Long
   Dim bDoAll As Boolean
   
   bDoAll = Not LstIsAllD(dLst())
   If bDoAll Then                      ' --- not allocated
      bDoAll = Not (lNRow = 0)
   Else                                ' --- is allocated
      lARow = UBound(dLst(), 1) + Cl01 - LBound(dLst(), 1)
      If lNRow = 0 Then lNRow = lARow
      bDoAll = Not (lNRow = lARow)
   End If
   If bDoAll Then ReDim dLst(1 To lNRow)

End Sub

Public Sub LstAllocF(fLst() As Double, _
                     Optional lNRow As Long = 0)

' Allocates or re-allocates list if necessary
' Allocation is necessary if unallocated
' Re-allocation is necessary if size does not match
' if lNRow=0 then returns list size

   Const Cl01 As Long = 1

   Dim lARow As Long
   Dim bDoAll As Boolean
   
   bDoAll = Not LstIsAllF(fLst())
   If bDoAll Then                      ' --- not allocated
      bDoAll = Not (lNRow = 0)
   Else                                ' --- is allocated
      lARow = UBound(fLst(), 1) + Cl01 - LBound(fLst(), 1)
      If lNRow = 0 Then lNRow = lARow
      bDoAll = Not (lNRow = lARow)
   End If
   If bDoAll Then ReDim fLst(1 To lNRow)

End Sub

Public Sub LstAllocL(lLst() As Long, _
                     Optional lNRow As Long = 0)

' Allocates or re-allocates list if necessary
' Allocation is necessary if unallocated
' Re-allocation is necessary if size does not match
' if lNRow=0 then returns list size

   Const Cl01 As Long = 1

   Dim lARow As Long
   Dim bDoAll As Boolean
   
   bDoAll = Not LstIsAllL(lLst())
   If bDoAll Then                      ' --- not allocated
      bDoAll = Not (lNRow = 0)
   Else                                ' --- is allocated
      lARow = UBound(lLst(), 1) + Cl01 - LBound(lLst(), 1)
      If lNRow = 0 Then lNRow = lARow
      bDoAll = Not (lNRow = lARow)
   End If
   If bDoAll Then ReDim lLst(1 To lNRow)

End Sub

Public Sub LstAllocS(sLst() As String, _
                     Optional lNRow As Long = 0)

' Allocates or re-allocates list if necessary
' Allocation is necessary if unallocated
' Re-allocation is necessary if size does not match
' if lNRow=0 then returns list size

   Const Cl01 As Long = 1

   Dim lARow As Long
   Dim bDoAll As Boolean
   
   bDoAll = Not LstIsAllS(sLst())
   If bDoAll Then                      ' --- not allocated
      bDoAll = Not (lNRow = 0)
   Else                                ' --- is allocated
      lARow = UBound(sLst(), 1) + Cl01 - LBound(sLst(), 1)
      If lNRow = 0 Then lNRow = lARow
      bDoAll = Not (lNRow = lARow)
   End If
   If bDoAll Then ReDim sLst(1 To lNRow)

End Sub

Public Sub LstAllocV(vLst() As Variant, _
                     Optional lNRow As Long = 0)

' Allocates or re-allocates list if necessary
' Allocation is necessary if unallocated
' Re-allocation is necessary if size does not match
' if lNRow=0 then returns list size

   Const Cl01 As Long = 1

   Dim lARow As Long
   Dim bDoAll As Boolean
   
   bDoAll = Not LstIsAllV(vLst())
   If bDoAll Then                      ' --- not allocated
      bDoAll = Not (lNRow = 0)
   Else                                ' --- is allocated
      lARow = UBound(vLst(), 1) + Cl01 - LBound(vLst(), 1)
      If lNRow = 0 Then lNRow = lARow
      bDoAll = Not (lNRow = lARow)
   End If
   If bDoAll Then ReDim vLst(1 To lNRow)

End Sub

Public Sub LstBoundD(dLst() As Date, _
                     Optional lE1 As Long = 0, Optional lEL As Long = 0)

' Returns bounds of dates list, if allocated

   If LstIsAllD(dLst()) Then
      lE1 = LBound(dLst(), 1)
      lEL = UBound(dLst(), 1)
   End If

End Sub

Public Sub LstBoundF(fLst() As Double, _
                     Optional lE1 As Long = 0, Optional lEL As Long = 0)

' Returns bounds of double precision real value list, if allocated

   If LstIsAllF(fLst()) Then
      lE1 = LBound(fLst(), 1)
      lEL = UBound(fLst(), 1)
   End If

End Sub

Public Sub LstBoundL(lLst() As Long, _
                     Optional lE1 As Long = 0, Optional lEL As Long = 0)

' Returns bounds of long integer list, if allocated

   If LstIsAllL(lLst()) Then
      lE1 = LBound(lLst(), 1)
      lEL = UBound(lLst(), 1)
   End If

End Sub

Public Sub LstBoundS(sLst() As String, _
                     Optional lE1 As Long = 0, Optional lEL As Long = 0)

' Returns bounds of string list, if allocated

   If LstIsAllS(sLst()) Then
      lE1 = LBound(sLst(), 1)
      lEL = UBound(sLst(), 1)
   End If

End Sub

Public Sub LstBoundV(vLst() As Variant, _
                     Optional lE1 As Long = 0, Optional lEL As Long = 0)

' Returns bounds of variant list, if allocated

   If LstIsAllV(vLst()) Then
      lE1 = LBound(vLst(), 1)
      lEL = UBound(vLst(), 1)
   End If

End Sub

Public Function LstInsDat(dLst() As Date, dEle As Date, _
                          Optional lEle As Long = 0)

' Expands dLst() by one element and inserts dEle at position lEle
' if lEle is at the lower bound, all the list is shifted
' if lEle is zero, the element is appended at the end

   Dim lE1 As Long, lEL As Long, lEI As Long, lEO As Long

   LstBoundD dLst(), lE1, lEO
   If lEO > 0 Then
      lEL = lEO + 1
      ReDim Preserve dLst(lE1 To lEL)
   Else
      lE1 = 1
      lEO = 0
      lEL = 1
      LstAllocD dLst(), lEL
   End If
   If lEle = 0 Then lEle = lEL
   For lEI = lEO To lEle Step -1
      dLst(lEI + 1) = dLst(lEI)
   Next lEI
   dLst(lEle) = dEle

End Function

Public Function LstInsLon(lLst() As Long, lEle As Long, _
                          Optional lPos As Long = 0)

' Expands lLst() by one element and inserts lEle at position lEle
' if lEle is at the lower bound, all the list is shifted
' if lEle is zero, the element is appended at the end

   Dim lE1 As Long, lEL As Long, lEI As Long, lEO As Long

   LstBoundL lLst(), lE1, lEO
   If lEO > 0 Then
      lEL = lEO + 1
      ReDim Preserve lLst(lE1 To lEL)
   Else
      lE1 = 1
      lEO = 0
      lEL = 1
      LstAllocL lLst(), lEL
   End If
   If lPos = 0 Then lPos = lEL
   For lEI = lEO To lPos Step -1
      lLst(lEI + 1) = lLst(lEI)
   Next lEI
   lLst(lPos) = lEle

End Function

Public Function LstInsDbl(fLst() As Double, fEle As Double, _
                          Optional lEle As Long = 0)

' Expands fLst() by one element and inserts fEle at position lEle
' if lEle is at the lower bound, all the list is shifted
' if lEle is zero, the element is appended at the end

   Dim lE1 As Long, lEL As Long, lEI As Long, lEO As Long

   LstBoundF fLst(), lE1, lEO
   If lEO > 0 Then
      lEL = lEO + 1
      ReDim Preserve fLst(lE1 To lEL)
   Else
      lE1 = 1
      lEO = 0
      lEL = 1
      LstAllocF fLst(), lEL
   End If
   If lEle = 0 Then lEle = lEL
   For lEI = lEO To lEle Step -1
      fLst(lEI + 1) = fLst(lEI)
   Next lEI
   fLst(lEle) = fEle

End Function

Public Function LstInsStr(sLst() As String, sEle As String, _
                          Optional lEle As Long = 0)

' Expands sLst() by one element and inserts sEle at position lEle
' if lEle is at the lower bound, all the list is shifted
' if lEle is zero, the element is appended at the end

   Dim lE1 As Long, lEL As Long, lEI As Long, lEO As Long

   LstBoundS sLst(), lE1, lEO
   If lEO > 0 Then
      lEL = lEO + 1
      ReDim Preserve sLst(lE1 To lEL)
   Else
      lE1 = 1
      lEO = 0
      lEL = 1
      LstAllocS sLst(), lEL
   End If
   If lEle = 0 Then lEle = lEL
   For lEI = lEO To lEle Step -1
      sLst(lEI + 1) = sLst(lEI)
   Next lEI
   sLst(lEle) = sEle

End Function

Public Function LstFind_F(fLst() As Double, fFnd As Double, _
                          Optional lFrom As Long = 0) As Long

   Const ClM1 As Long = -1

   Dim lE1 As Long, lEL As Long, lEI As Long
   Dim bMatch As Boolean

   LstBoundF fLst(), lE1, lEL
   If lFrom > lEL Then Exit Function
   If lFrom > 0 Then lE1 = lFrom + ClM1
   While lEI < lEL And Not bMatch
      lEI = lEI + 1
      bMatch = fLst(lEI) = fFnd
   Wend
   If Not bMatch Then lEI = ClM1

   LstFind_F = lEI

End Function

Public Function LstFind_L(lLst() As Long, lFnd As Long, _
                          Optional lFrom As Long = 0) As Long

   Const ClM1 As Long = -1

   Dim lE1 As Long, lEL As Long, lEI As Long
   Dim bMatch As Boolean

   LstBoundL lLst(), lE1, lEL
   If lFrom > lEL Then Exit Function
   If lFrom > 0 Then lE1 = lFrom + ClM1
   While lEI < lEL And Not bMatch
      lEI = lEI + 1
      bMatch = lLst(lEI) = lFnd
   Wend
   If Not bMatch Then lEI = ClM1

   LstFind_L = lEI

End Function

Public Function LstFind_S(sLst() As String, sFnd As String, _
                          Optional lFrom As Long = 0) As Long

   Const ClM1 As Long = -1

   Dim lE1 As Long, lEL As Long, lEI As Long
   Dim bMatch As Boolean

   LstBoundS sLst(), lE1, lEL
   If lFrom > lEL Then Exit Function
   If lFrom > 0 Then lE1 = lFrom + ClM1
   While lEI < lEL And Not bMatch
      lEI = lEI + 1
      bMatch = sLst(lEI) = sFnd
   Wend
   If Not bMatch Then lEI = ClM1

   LstFind_S = lEI

End Function

Public Sub LstGenerD(dLst() As Date, lCnt As Long, _
                     Optional dIni As Date = 0, _
                     Optional dInc As Date = 1)

' Generate a list of date values from a starting value and increment

   Const Cl01 As Long = 1

   Dim d As Date
   Dim lIE As Long

   LstAllocD dLst(), lCnt
   d = dIni
   For lIE = Cl01 To lCnt
      dLst(lIE) = d
      d = d + dInc
   Next lIE

End Sub

Public Sub LstGenerF(fLst() As Double, lCnt As Long, _
                     Optional fIni As Double = 0, _
                     Optional fInc As Double = 1)

' Generate a list of real values from a starting value and increment

   Const Cl01 As Long = 1

   Dim f As Double
   Dim lIE As Long

   LstAllocF fLst(), lCnt
   f = fIni
   For lIE = Cl01 To lCnt
      fLst(lIE) = f
      f = f + fInc
   Next lIE

End Sub

Public Sub LstGenerL(lLst() As Long, lCnt As Long, _
                     Optional lIni As Long = 0, _
                     Optional lInc As Long = 1)

' Generate a list of long integer values from a starting value and increment

   Const Cl01 As Long = 1

   Dim l As Long
   Dim lIE As Long

   LstAllocL lLst(), lCnt
   l = lIni
   For lIE = Cl01 To lCnt
      lLst(lIE) = l
      l = l + lInc
   Next lIE

End Sub

Public Sub LstGet_FF(fSrc() As Double, fDst() As Double, _
                        Optional lEle1 As Long = 0, _
                        Optional lEleL As Long = 0)

' Extract part of a double list to double list
' When lEle1=0, starts at first source column
' When lEle1>0, starts at source column number lCol1
' When lEle1<0, starts at source column Abs(lCol1) from the right
' When lEleL=0, ends at last row

   Const Cl01 As Long = 1

   Dim lSE1 As Long, lSEL As Long
   Dim lDE1 As Long, lDEL As Long
   Dim lUE1 As Long, lUEL As Long
   Dim lIES As Long, lIED As Long

   LstBoundF fSrc(), lSE1, lSEL
   QRS_LibArr.NdxGetXSD lSE1, lSEL, lEle1, lEleL, lDE1, lDEL, lUE1, lUEL
   LstAllocF fDst(), lUEL

   lIED = lUE1
   For lIES = lEle1 To lEleL
      fDst(lIED) = fSrc(lIES)
      lIED = lIED + Cl01
   Next lIES

End Sub

Public Sub LstGet_LL(lSrc() As Long, lDst() As Long, _
                        Optional lEle1 As Long = 0, _
                        Optional lEleL As Long = 0)

' Extract part of a double list to double list
' When lEle1=0, starts at first source column
' When lEle1>0, starts at source column number lCol1
' When lEle1<0, starts at source column Abs(lCol1) from the right
' When lEleL=0, ends at last row

   Const Cl01 As Long = 1

   Dim lSE1 As Long, lSEL As Long
   Dim lDE1 As Long, lDEL As Long
   Dim lUE1 As Long, lUEL As Long
   Dim lIES As Long, lIED As Long

   LstBoundL lSrc(), lSE1, lSEL
   QRS_LibArr.NdxGetXSD lSE1, lSEL, lEle1, lEleL, lDE1, lDEL, lUE1, lUEL
   LstAllocL lDst(), lUEL

   lIED = lUE1
   For lIES = lEle1 To lEleL
      lDst(lIED) = lSrc(lIES)
      lIED = lIED + Cl01
   Next lIES

End Sub

Public Sub LstGet_SS(sSrc() As String, sDst() As String, _
                        Optional lEle1 As Long = 0, _
                        Optional lEleL As Long = 0)

' Extract part of a double list to double list
' When lEle1=0, starts at first source column
' When lEle1>0, starts at source column number lCol1
' When lEle1<0, starts at source column Abs(lCol1) from the right
' When lEleL=0, ends at last row

   Const Cl01 As Long = 1

   Dim lSE1 As Long, lSEL As Long
   Dim lDE1 As Long, lDEL As Long
   Dim lUE1 As Long, lUEL As Long
   Dim lIES As Long, lIED As Long

   LstBoundS sSrc(), lSE1, lSEL
   QRS_LibArr.NdxGetXSD lSE1, lSEL, lEle1, lEleL, lDE1, lDEL, lUE1, lUEL
   LstAllocS sDst(), lUEL

   lIED = lUE1
   For lIES = lEle1 To lEleL
      sDst(lIED) = sSrc(lIES)
      lIED = lIED + Cl01
   Next lIES

End Sub

Public Sub LstPut_FF(fSrc() As Double, fDst() As Double, _
                        Optional lEle1 As Long = 0)

' Output a double list to a column in a double array
' When lEle1=0, starts at first dest column
' When lEle1>0, starts at dest column number lCol1
' When lEle1<0, starts at dest column Abs(lCol1) from the right
' When lEleL=0, ends at last row

   Const Cl01 As Long = 1

   Dim lSE1 As Long, lSEL As Long
   Dim lDE1 As Long, lDEL As Long
   Dim lVE1 As Long, lVEL As Long
   Dim lIES As Long, lIED As Long
   Dim lEleL As Long

   LstBoundF fSrc(), lSE1, lSEL
   LstBoundF fDst(), lDE1, lDEL
   QRS_LibArr.NdxGetXDS lDE1, lDEL, lSE1, lSEL, lEle1, lEleL, lVE1, lVEL

   lIES = lVE1                         ' --- Valid source 1 index
   For lIED = lEle1 To lEleL
      fDst(lIED) = fSrc(lIES)
      lIES = lIES + Cl01
   Next lIED

End Sub

Public Sub LstPut_LL(lSrc() As Long, lDst() As Long, _
                        Optional lEle1 As Long = 0)

' Output a long integer list to long integer list from specific element index
' When lEle1=0, starts at first dest column
' When lEle1>0, starts at dest column number lCol1
' When lEle1<0, starts at dest column Abs(lCol1) from the right
' When lEleL=0, ends at last row

   Const Cl01 As Long = 1

   Dim lSE1 As Long, lSEL As Long
   Dim lDE1 As Long, lDEL As Long
   Dim lVE1 As Long, lVEL As Long
   Dim lIES As Long, lIED As Long
   Dim lEleL As Long

   LstBoundL lSrc(), lSE1, lSEL
   LstBoundL lDst(), lDE1, lDEL
   QRS_LibArr.NdxGetXDS lDE1, lDEL, lSE1, lSEL, lEle1, lEleL, lVE1, lVEL

   lIES = lVE1                         ' --- Valid source1 index
   For lIED = lEle1 To lEleL
      lDst(lIED) = lSrc(lIES)
      lIES = lIES + Cl01
   Next lIED

End Sub

Public Sub LstPut_SS(sSrc() As String, sDst() As String, _
                        Optional lEle1 As Long = 0)

' Output a string list to string list from specific element index
' When lEle1=0, starts at first dest column
' When lEle1>0, starts at dest column number lCol1
' When lEle1<0, starts at dest column Abs(lCol1) from the right
' When lEleL=0, ends at last row

   Const Cl01 As Long = 1

   Dim lSE1 As Long, lSEL As Long
   Dim lDE1 As Long, lDEL As Long
   Dim lVE1 As Long, lVEL As Long
   Dim lIES As Long, lIED As Long
   Dim lEleL As Long

   LstBoundS sSrc(), lSE1, lSEL
   LstBoundS sDst(), lDE1, lDEL
   QRS_LibArr.NdxGetXDS lDE1, lDEL, lSE1, lSEL, lEle1, lEleL, lVE1, lVEL

   lIES = lVE1                         ' --- Clipped source1 index
   For lIED = lEle1 To lEleL
      sDst(lIED) = sSrc(lIES)
      lIES = lIES + Cl01
   Next lIED

End Sub

Public Sub LstMergeF(fLstA() As Double, fLstB() As Double, fLstC() As Double)

' Merges two ordered input lists fLstA() and fLstB()
' into an ordered output list fLstC()

   Const Cl01 As Long = 1

   Dim fA As Double, fB As Double
   Dim lE1A As Long, lELA As Long, lNA As Long, lIA As Long
   Dim lE1B As Long, lELB As Long, lNB As Long, lIB As Long
   Dim lNC As Long, lIC As Long
   Dim bAD As Boolean, bBD As Boolean  ' --- "Done" flags
   Dim bA As Boolean                   ' --- take from list A

   LstBoundF fLstA(), lE1A, lELA: lNA = lELA + Cl01 - lE1A
   LstBoundF fLstB(), lE1B, lELB: lNB = lELB + Cl01 - lE1B
   lNC = lNA + lNB
   LstAllocF fLstC(), lNC

   lIA = lE1A:   fA = fLstA(lIA)
   lIB = lE1B:   fB = fLstB(lIB)

   For lIC = Cl01 To lNC
      bA = bBD                         ' --- fLstB through? -> UseA
      If Not (bAD Or bBD) Then bA = fA < fB
      If bA Then
         fLstC(lIC) = fA
         lIA = lIA + Cl01
         bAD = lIA > lNA
         If Not bAD Then fA = fLstA(lIA)
      Else
         fLstC(lIC) = fB
         lIB = lIB + Cl01
         bBD = lIB > lNB
         If Not bBD Then fB = fLstB(lIB)
      End If
   Next lIC

End Sub

Public Sub LstMergeL(lLstA() As Long, lLstB() As Long, lLstC() As Long)

' Merges two ordered input lists lLstA() and lLstB()
' into an ordered output list lLstC()

   Const Cl01 As Long = 1

   Dim lA As Long, lB As Long
   Dim lE1A As Long, lELA As Long, lNA As Long, lIA As Long
   Dim lE1B As Long, lELB As Long, lNB As Long, lIB As Long
   Dim lNC As Long, lIC As Long
   Dim bAD As Boolean, bBD As Boolean  ' --- "Done" flags
   Dim bA As Boolean                   ' --- take from list A

   LstBoundL lLstA(), lE1A, lELA: lNA = lELA + Cl01 - lE1A
   LstBoundL lLstB(), lE1B, lELB: lNB = lELB + Cl01 - lE1B
   lNC = lNA + lNB
   LstAllocL lLstC(), lNC

   lIA = lE1A:   lA = lLstA(lIA)
   lIB = lE1B:   lB = lLstB(lIB)

   For lIC = Cl01 To lNC
      bA = bBD                         ' --- lLstB through? -> UseA
      If Not (bAD Or bBD) Then bA = lA < lB
      If bA Then
         lLstC(lIC) = lA
         lIA = lIA + Cl01
         bAD = lIA > lNA
         If Not bAD Then lA = lLstA(lIA)
      Else
         lLstC(lIC) = lB
         lIB = lIB + Cl01
         bBD = lIB > lNB
         If Not bBD Then lB = lLstB(lIB)
      End If
   Next lIC

End Sub

Public Sub LstMergeS(sLstA() As String, sLstB() As String, sLstC() As String)

' Merges two ordered input lists sLstA() and sLstB()
' into an ordered output list sLstC()

   Const Cl01 As Long = 1

   Dim sA As String, sB As String
   Dim lE1A As Long, lELA As Long, lNA As Long, lIA As Long
   Dim lE1B As Long, lELB As Long, lNB As Long, lIB As Long
   Dim lNC As Long, lIC As Long
   Dim bAD As Boolean, bBD As Boolean  ' --- "Done" flags
   Dim bA As Boolean                   ' --- take from list A

   LstBoundS sLstA(), lE1A, lELA: lNA = lELA + Cl01 - lE1A
   LstBoundS sLstB(), lE1B, lELB: lNB = lELB + Cl01 - lE1B
   lNC = lNA + lNB
   LstAllocS sLstC(), lNC

   lIA = lE1A:   sA = sLstA(lIA)
   lIB = lE1B:   sB = sLstB(lIB)

   For lIC = Cl01 To lNC
      bA = bBD                         ' --- sLstB through? -> UseA
      If Not (bAD Or bBD) Then bA = StrComp(sA, sB, vbTextCompare) < 0
      If bA Then
         sLstC(lIC) = sA
         lIA = lIA + Cl01
         bAD = lIA > lNA
         If Not bAD Then sA = sLstA(lIA)
      Else
         sLstC(lIC) = sB
         lIB = lIB + Cl01
         bBD = lIB > lNB
         If Not bBD Then sB = sLstB(lIB)
      End If
   Next lIC

End Sub

Public Sub LstReverD(dLst() As Date)

' Reverse the element order in dLst()
' If the numer of elements in dLst() is odd,
' the middle element is not processed

   Const Cl01 As Long = 1

   Dim d As Date
   Dim lE1 As Long, lEL As Long, lNE As Long
   Dim lS1 As Long, lSL As Long, lNS As Long
   Dim bOdd As Boolean

   LstBoundD dLst(), lE1, lEL
   lNE = lEL + Cl01 - lE1
   lNS = lNE
   bOdd = lNS And Cl01 = Cl01          ' --- Odd length
   If bOdd Then lNS = lNS - Cl01       ' --- Subtract 1
   lNS = lNS / 2                       '     Half length

   lSL = lEL
   For lS1 = lE1 To lNS
      d = dLst(lS1)
      dLst(lS1) = dLst(lSL)
      dLst(lSL) = d
      lSL = lSL - Cl01
   Next lS1

End Sub

Public Sub LstReverF(fLst() As Double)

' Reverse the element order in fLst()
' If the numer of elements in fLst() is odd,
' the middle element is not processed

   Const Cl01 As Long = 1

   Dim f As Double
   Dim lE1 As Long, lEL As Long, lNE As Long
   Dim lS1 As Long, lSL As Long, lNS As Long
   Dim bOdd As Boolean

   LstBoundF fLst(), lE1, lEL
   lNE = lEL + Cl01 - lE1
   lNS = lNE
   bOdd = lNS And Cl01 = Cl01          ' --- Odd length
   If bOdd Then lNS = lNS - Cl01       ' --- Subtract 1
   lNS = lNS / 2                       '     Half length

   lSL = lEL
   For lS1 = lE1 To lNS
      f = fLst(lS1)
      fLst(lS1) = fLst(lSL)
      fLst(lSL) = f
      lSL = lSL - Cl01
   Next lS1

End Sub

Public Sub LstReverL(lLst() As Long)

' Reverse the element order in lLst()
' If the numer of elements in lLst() is odd,
' the middle element is not processed

   Const Cl01 As Long = 1

   Dim l As Long
   Dim lE1 As Long, lEL As Long, lNE As Long
   Dim lS1 As Long, lSL As Long, lNS As Long
   Dim bOdd As Boolean

   LstBoundL lLst(), lE1, lEL
   lNE = lEL + Cl01 - lE1
   bOdd = lNE And Cl01 = Cl01
   lNS = lNE
   If bOdd Then lNS = lNS - Cl01
   lNS = lNS / 2

   lSL = lEL
   For lS1 = lE1 To lNS
      l = lLst(lS1)
      lLst(lS1) = lLst(lSL)
      lLst(lSL) = l
      lSL = lSL - Cl01
   Next lS1

End Sub

Public Sub LstReverS(sLst() As String)

' Reverse the element order in sLst()
' If the numer of elements in sLst() is odd,
' the middle element is not processed

   Const Cl01 As Long = 1

   Dim s As String
   Dim lE1 As Long, lEL As Long, lNE As Long
   Dim lS1 As Long, lSL As Long, lNS As Long
   Dim bOdd As Boolean

   LstBoundS sLst(), lE1, lEL
   lNE = lEL + Cl01 - lE1
   bOdd = lNE And Cl01 = Cl01
   lNS = lNE
   If bOdd Then lNS = lNS - Cl01
   lNS = lNS / 2

   lSL = lEL
   For lS1 = lE1 To lNS
      s = sLst(lS1)
      sLst(lS1) = sLst(lSL)
      sLst(lSL) = s
      lSL = lSL - Cl01
   Next lS1

End Sub

Public Sub LstReverV(vLst() As Variant)

' Reverse the element order in vLst()
' If the numer of elements in vLst() is odd,
' the middle element is not processed

   Const Cl01 As Long = 1

   Dim v As Variant
   Dim lE1 As Long, lEL As Long, lNE As Long
   Dim lS1 As Long, lSL As Long, lNS As Long
   Dim bOdd As Boolean

   LstBoundV vLst(), lE1, lEL
   lNE = lEL + Cl01 - lE1
   bOdd = lNE And Cl01 = Cl01
   lNS = lNE
   If bOdd Then lNS = lNS - Cl01
   lNS = lNS / 2

   lSL = lEL
   For lS1 = lE1 To lNS
      v = vLst(lS1)
      vLst(lS1) = vLst(lSL)
      vLst(lSL) = v
      lSL = lSL - Cl01
   Next lS1

End Sub

Public Function LstSetC5V(vLst() As Variant, lEle1 As Long, _
                          Optional A1, Optional A2, Optional A3, _
                          Optional A4, Optional A5) As Boolean

' Set 5 consecutive elements in vLst() from lEle1
' if any of the value arguments are omitted,
' the corresponding element it not set.
' All 5 elements are handled anyway

   Dim lE1 As Long, lEL As Long, lEI As Long
   Dim bErr As Boolean, bFail As Boolean

   LstBoundV vLst(), lE1, lEL
   lEle1 = QRS_LibArr.NdxGetSX1(lE1, lEL, lEle1)

   If Not IsMissing(A1) Then
      lEI = lEle1
      bErr = lEI < lE1 Or lEI > lEL
      bFail = bFail Or bErr
      If Not bErr Then vLst(lEI) = A1
   End If

   If Not IsMissing(A2) Then
      lEI = lEle1 + 1
      bErr = lEI < lE1 Or lEI > lEL
      bFail = bFail Or bErr
      If Not bErr Then vLst(lEI) = A2
   End If

   If Not IsMissing(A3) Then
      lEI = lEle1 + 2
      bErr = lEI < lE1 Or lEI > lEL
      bFail = bFail Or bErr
      If Not bErr Then vLst(lEI) = A3
   End If

   If Not IsMissing(A4) Then
      lEI = lEle1 + 3
      bErr = lEI < lE1 Or lEI > lEL
      bFail = bFail Or bErr
      If Not bErr Then vLst(lEI) = A4
   End If

   If Not IsMissing(A5) Then
      lEI = lEle1 + 4
      bErr = lEI < lE1 Or lEI > lEL
      bFail = bFail Or bErr
      If Not bErr Then vLst(lEI) = A5
   End If

   LstSetC5V = bFail                   ' --- Return true if any set failed

End Function

Public Function LstSetR5V(vLst() As Variant, _
                          Optional A1, Optional lE1 As Long = 1, _
                          Optional A2, Optional lE2 As Long = 0, _
                          Optional A3, Optional lE3 As Long = 0, _
                          Optional A4, Optional lE4 As Long = 0, _
                          Optional A5, Optional lE5 As Long = 0)

' Puts elements in into a variant list
' If arguments are present for the element values V1 to V5,
' the values are updated in the list.
' If any of the elements indices are 0, they are assumed as
' the increment from the previously present index, e.g:
'    1,0,0,5,0 would result in putting elements 1,2,3,5,6

   Dim lE0 As Long, lEL As Long, lEI As Long
   Dim bErr As Boolean, bFail As Boolean

   LstBoundV vLst(), lE0, lEL
   lEI = QRS_LibArr.NdxGetSX1(lE0, lEL, lE1)

   If Not IsMissing(A1) Then
      bErr = lEI < lE0 Or lEI > lEL
      bFail = bFail Or bErr
      If Not bErr Then vLst(lEI) = A1
   End If

   If Not IsMissing(A2) Then
      If lE2 = 0 Then lEI = lEI + 1 Else lEI = lE2
      bErr = lEI < lE0 Or lEI > lEL
      bFail = bFail Or bErr
      If Not bErr Then vLst(lEI) = A2
   End If

   If Not IsMissing(A3) Then
      If lE3 = 0 Then lEI = lEI + 1 Else lEI = lE3
      bErr = lEI < lE0 Or lEI > lEL
      bFail = bFail Or bErr
      If Not bErr Then vLst(lEI) = A3
   End If

   If Not IsMissing(A4) Then
      If lE4 = 0 Then lEI = lEI + 1 Else lEI = lE4
      bErr = lEI < lE0 Or lEI > lEL
      bFail = bFail Or bErr
      If Not bErr Then vLst(lEI) = A4
   End If

   If Not IsMissing(A5) Then
      If lE5 = 0 Then lEI = lEI + 1 Else lEI = lE5
      bErr = lEI < lE0 Or lEI > lEL
      bFail = bFail Or bErr
      If Not bErr Then vLst(lEI) = A5
   End If

   LstSetR5V = bFail                   ' --- return true if any set failed

End Function

Public Function LstXtrC5S(sLst() As String, lEle1 As Long, _
                          Optional A1, Optional A2, Optional A3, _
                          Optional A4, Optional A5, _
                          Optional bTrim As Boolean) As Boolean

' Extracts 5 consecutive elements from the list starting a lEle1

   Dim lS1 As Long, lSL As Long, lSX As Long
   Dim bErr As Boolean, bFail As Boolean

   LstBoundS sLst(), lS1, lSL
   lEle1 = QRS_LibArr.NdxGetSX1(lS1, lSL, lEle1)

   If Not IsMissing(A1) Then
      lSX = lEle1
      bErr = lSX < lS1 Or lSX > lSL
      bFail = bFail Or bErr
      If Not bErr Then QRS_Lib0.CnvStrTyp sLst(lSX), A1, bTrim
   End If

   If Not IsMissing(A2) Then
      lSX = lEle1 + 1
      bErr = lSX < lS1 Or lSX > lSL
      bFail = bFail Or bErr
      If Not bErr Then QRS_Lib0.CnvStrTyp sLst(lSX), A2, bTrim
   End If

   If Not IsMissing(A3) Then
      lSX = lEle1 + 2
      bErr = lSX < lS1 Or lSX > lSL
      bFail = bFail Or bErr
      If Not bErr Then QRS_Lib0.CnvStrTyp sLst(lSX), A3, bTrim
   End If

   If Not IsMissing(A4) Then
      lSX = lEle1 + 3
      bErr = lSX < lS1 Or lSX > lSL
      bFail = bFail Or bErr
      If Not bErr Then QRS_Lib0.CnvStrTyp sLst(lSX), A4, bTrim
   End If

   If Not IsMissing(A5) Then
      lSX = lEle1 + 4
      bErr = lSX < lS1 Or lSX > lSL
      bFail = bFail Or bErr
      If Not bErr Then QRS_Lib0.CnvStrTyp sLst(lSX), A5, bTrim
   End If

   LstXtrC5S = bFail                   ' --- Return true if any get failed

End Function

