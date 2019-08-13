Attribute VB_Name = "QRS_LibQRS"
Option Explicit

' Module : QRS_LibQRS
' Project: any
' Purpose: Quick Rank and Sort routines on lists and arrays
' By     : QRS, Roger Strebel
' Date   : 01.04.2018                  LstQSortF works
'          02.04.2018                  LstQRNdxF and LstQRankF work
'          24.04.2018                  LstQFindF works
' --- The public interface
'     ArrSortCF                        Order real array by column
'     ArrSortRF                        Order real array by row
'     LstQFindF                        Quick find in ordered list    24.04.2018
'     LstQFindL                        Quick find in ordered list    24.04.2018
'     LstQFindS                        Quick find in ordered list    24.04.2018
'     LstQRankF                        Rank of real list values      02.04.2018
'     LstQRankL                        Rank of real list values      02.04.2018
'     LstQRankS                        Rank of real list values      02.04.2018
'     LstQRNdxF                        Rank index of real values     02.04.2018
'     LstQRNdxL                        Rank index of long values     02.04.2018
'     LstQRNdxS                        Rank index of string values   02.04.2018
'     LstQSortF                        Order list of real values     01.04.2018
'     LstQSortL                        Order list of long values     01.04.2018
'     LstQSortS                        Order list of string values   01.04.2018
'     LstQUniqF                        Unique real values in list    02.04.2018
'     LstQUniqL                        Unique long values in list    02.04.2018
'     LstQUniqS                        Unique string values in list  02.04.2018
' --- The private sphere
'     LstMrgNdxF                       Merge value and index lists   02.04.2018
'     LstMrgSubF                       Merge consecutive sublists    26.03.2018
'     LstMrgNdxL                       Merge value and index lists   02.04.2018
'     LstMrgSubL                       Merge consecutive sublists    26.03.2018
'     LstMrgNdxS                       Merge value and index lists   02.04.2018
'     LstMrgSubS                       Merge consecutive sublists    26.03.2018

Public Function LstQFindF(fLst() As Double, fVal As Double) As Long

' Obtain index of a real value fVal in fLst()
' if a match is found, returns the index of the matching element in fLst()
' if no match is found, returns the index of where the element is inserted
' as a negative number (-1: insert as first element, -n+1: insert at end)

   Const ClM1 As Long = -1, Cl01 As Long = 1, Cl02 As Long = 2

   Dim fR As Double
   Dim lE1 As Long, lEL As Long, lE As Long, lI As Long, lS As Long
   Dim bDone As Boolean

   QRS_LibLst.LstBoundF fLst(), lE1, lEL
   QRS_Lib0.BitGetMSB lEL, lI          ' --- Half length
   lS = 1                              ' --- Sign of increment
   lE = lI

   Do
      If lE > lEL Then                 ' --- Past upper bound
         lS = ClM1                     '
      Else                             ' --- In range
         fR = fVal - fLst(lE)          ' --- Difference
         If fR < 0 Then                ' --- Too high
            bDone = lI = Cl01          '     Step = 1: Done
            lS = ClM1                  '     Increment down
         Else                          ' --- Too low or match
            If fR > 0 Then             '     Too low
               bDone = lI = Cl01       '     Step = 1: Done
               lS = Cl01               '     Increment up
            Else                       ' --- Match
               bDone = True            '     Done
            End If
         End If
      End If
      If Not bDone Then
         lI = lI / Cl02                ' --- Half step
         lE = lE + lI * lS             ' --- Apply increment
      End If
   Loop While Not bDone

   If Not fR = 0 Then
      lE = -lE
      If fR > 0 Then lE = lE - Cl01
   End If
   LstQFindF = lE

End Function

Public Function LstQFindL(lLst() As Long, lVal As Long) As Long

' Obtain index of an integer value lVal in lLst()
' if a match is found, returns the index of the matching element in fLst()
' if no match is found, returns the index of where the element is inserted
' as a negative number (-1: insert as first element, -n+1: insert at end)

   Const ClM1 As Long = -1, Cl01 As Long = 1, Cl02 As Long = 2

   Dim lR As Double
   Dim lE1 As Long, lEL As Long, lE As Long, lI As Long, lS As Long
   Dim bDone As Boolean

   QRS_LibLst.LstBoundL lLst(), lE1, lEL
   QRS_Lib0.BitGetMSB lEL, lI          ' --- Half length
   lS = 1                              ' --- Sign of increment
   lE = lI

   Do
      If lE > lEL Then                 ' --- Past upper bound
         lS = ClM1                     '
      Else                             ' --- In range
         lR = lVal - lLst(lE)          ' --- Difference
         If lR < 0 Then                ' --- Too high
            bDone = lI = Cl01          '     Step = 1: Done
            lS = ClM1                  '     Increment down
         Else                          ' --- Too low or match
            If lR > 0 Then             '     Too low
               bDone = lI = Cl01       '     Step = 1: Done
               lS = Cl01               '     Increment up
            Else                       ' --- Match
               bDone = True            '     Done
            End If
         End If
      End If
      If Not bDone Then
         lI = lI / Cl02                ' --- Half step
         lE = lE + lI * lS             ' --- Apply increment
      End If
   Loop While Not bDone

   If Not lR = 0 Then
      lE = -lE
      If lR > 0 Then lE = lE - Cl01
   End If
   LstQFindL = lE

End Function

Public Function LstQFindS(sLst() As String, sVal As String) As Long

' Obtain index of a string value sVal in sLst()
' if a match is found, returns the index of the matching element in fLst()
' if no match is found, returns the index of where the element is inserted
' as a negative number (-1: insert as first element, -n+1: insert at end)

   Const ClM1 As Long = -1, Cl01 As Long = 1, Cl02 As Long = 2

   Dim lR As Double
   Dim lE1 As Long, lEL As Long, lE As Long, lI As Long, lS As Long
   Dim bDone As Boolean

   QRS_LibLst.LstBoundS sLst(), lE1, lEL
   QRS_Lib0.BitGetMSB lEL, lI          ' --- Half length
   lS = 1                              ' --- Sign of increment
   lE = lI

   Do
      If lE > lEL Then                 ' --- Past upper bound
         lS = ClM1                     '
      Else                             ' --- In range
         lR = StrComp(sVal, sLst(lE), vbTextCompare)
         If lR < 0 Then                ' --- Too high
            bDone = lI = Cl01          '     Step = 1: Done
            lS = ClM1                  '     Increment down
         Else                          ' --- Too low or match
            If lR > 0 Then             '     Too low
               bDone = lI = Cl01       '     Step = 1: Done
               lS = Cl01               '     Increment up
            Else                       ' --- Match
               bDone = True            '     Done
            End If
         End If
      End If
      If Not bDone Then
         lI = lI / Cl02                ' --- Half step
         lE = lE + lI * lS             ' --- Apply increment
      End If
   Loop While Not bDone

   If Not lR = 0 Then
      lE = -lE
      If lR > 0 Then lE = lE - Cl01
   End If
   LstQFindS = lE

End Function

Public Sub LstQRankF(fLst() As Double, lRnk() As Long)

' Obtain ranks of real values in a list by their indices
'   Rank(fLst(i)) = lRnk(i)
' uses the quick sort method which first sorts pairs and then
' merges successively larger ordered lists

' This routine uses one buffer list that is allocated once
' and then the ordered sublists are kept track of by their
' indices

   Dim lNdx() As Long
   Dim lE1 As Long, lEL As Long, lE As Long

   LstQRNdxF fLst(), lNdx()
   QRS_LibLst.LstBoundL lNdx(), lE1, lEL
   QRS_LibLst.LstAllocL lRnk(), lEL

   For lE = lE1 To lEL
      lRnk(lNdx(lE)) = lE
   Next lE

End Sub

Public Sub LstQRankL(lLst() As Long, lRnk() As Long)

' Obtain ranks of long integer values in a list by their indices
'   Rank(lLst(i)) = lRnk(i)
' uses the quick sort method which first sorts pairs and then
' merges successively larger ordered lists

' This routine uses one buffer list that is allocated once
' and then the ordered sublists are kept track of by their
' indices

   Dim lNdx() As Long
   Dim lE1 As Long, lEL As Long, lE As Long

   LstQRNdxL lLst(), lNdx()
   QRS_LibLst.LstBoundL lNdx(), lE1, lEL
   QRS_LibLst.LstAllocL lRnk(), lEL

   For lE = lE1 To lEL
      lRnk(lNdx(lE)) = lE
   Next lE

End Sub

Public Sub LstQRankS(sLst() As String, lRnk() As Long, _
                     eCompareType As VbCompareMethod)

' Obtain ranks of string values in a list by their indices
'   Rank(sLst(i)) = lRnk(i)
' uses the quick sort method which first sorts pairs and then
' merges successively larger ordered lists

' This routine uses one buffer list that is allocated once
' and then the ordered sublists are kept track of by their
' indices

   Dim lNdx() As Long
   Dim lE1 As Long, lEL As Long, lE As Long

   LstQRNdxS sLst(), lNdx(), eCompareType
   QRS_LibLst.LstBoundL lNdx(), lE1, lEL
   QRS_LibLst.LstAllocL lRnk(), lEL

   For lE = lE1 To lEL
      lRnk(lNdx(lE)) = lE
   Next lE

End Sub

Public Sub LstQRNdxF(fLst() As Double, lNdx() As Long)

' Obtain indices by order of real values in a list.
'    lNdx(1) = min(fLst())
' uses the quick sort method on a copy of fLst()
' and ordering the index list in parallel
' Value lists and Index lists are doubled
' and copied back or forth on each pass

   Const Cl01 As Long = 1, Cl02 As Long = 2

   Dim fLstS() As Double, fLstT() As Double
   Dim lNdxT() As Long                 ' --- Buffer index list
   Dim f As Double
   Dim lE1 As Long, lEL As Long, lN2 As Long
   Dim lS1 As Long, lSL As Long, lNS As Long, lMS As Long
   Dim lT1 As Long, lTL As Long
   Dim l As Long
   Dim bOdd As Boolean, bDun As Boolean

   QRS_LibLst.LstGet_FF fLst(), fLstS() '    Copy fLst() to fLstT()
   QRS_LibLst.LstBoundF fLstS(), lE1, lEL  ' List bounds
   QRS_LibLst.LstAllocF fLstT(), lEL   '     Buffer values list
   QRS_LibLst.LstAllocL lNdxT(), lEL   '     Buffer index list
   QRS_LibLst.LstAllocL lNdx(), lEL    '     Index list
   For lS1 = lE1 To lEL                '     Index
      lNdx(lS1) = lS1
   Next lS1

   lNS = Cl02                          ' --- Sublist length = 2
   lMS = lEL - Cl01                    ' --- Order all pairs
   For lS1 = lE1 To lMS Step lNS
      lSL = lS1 + Cl01
      If fLstS(lS1) > fLstS(lSL) Then  ' --- Inverse order?
         f = fLstS(lS1)                '     -> swap value
         fLstS(lS1) = fLstS(lSL)
         fLstS(lSL) = f
         l = lNdx(lS1)                 '        swap index
         lNdx(lS1) = lNdx(lSL)
         lNdx(lSL) = l
      End If
   Next lS1

   ' --------------------------------------- Merge the sublists
   '                                         until their length
   '                                         exceeds half the list length

   QRS_Lib0.BitGetMSB lEL, lMS
   bOdd = False                        ' --- Odd true: Copy Lst -> LstT
   lNS = Cl02                          ' --- sublist length
   lN2 = lNS * Cl02                    ' --- two sublist lengths
   While Not lNS > lMS
      bOdd = Not bOdd                  '     Toggle back and forth
      lS1 = lE1                        ' --- Start from beginning
      While Not lS1 + lN2 > lEL        ' --- Process list pairs of regular size
         lT1 = lS1 + lNS               '     2nd sublist start
         lSL = lT1 - Cl01              '     1st sublist end
         lTL = lSL + lNS               '     2nd sublist end
         If bOdd Then                  ' --- Merge
            LstMrgNdxF fLstS(), lNdx(), lS1, lSL, lT1, lTL, fLstT(), lNdxT()
         Else
            LstMrgNdxF fLstT(), lNdxT(), lS1, lSL, lT1, lTL, fLstS(), lNdx()
         End If
         lS1 = lS1 + lN2               ' --- Increment 1st sublist start
      Wend
      lT1 = lS1 + lNS                  '     2nd sublist adjacent
      lSL = lT1 - Cl01                 '     First sublist end
      lTL = lSL + lNS                  '     2nd sublist default end
      If Not lSL < lEL Then            ' --- 1st sublist reaches end
         If bOdd Then                  '     Copy 1st sublist
            For lT1 = lS1 To lEL
               fLstT(lT1) = fLstS(lT1)
               lNdxT(lT1) = lNdx(lT1)
            Next lT1
         Else
            For lT1 = lS1 To lEL
               fLstS(lT1) = fLstT(lT1)
               lNdx(lT1) = lNdxT(lT1)
            Next lT1
         End If
      Else                             ' --- 2nd sublist reaches end
         If bOdd Then                  ' --- Merge uneven lists
            LstMrgNdxF fLstS(), lNdx(), lS1, lSL, lT1, lEL, fLstT(), lNdxT()
         Else
            LstMrgNdxF fLstT(), lNdxT(), lS1, lSL, lT1, lEL, fLstS(), lNdx()
         End If
      End If
      lNS = lN2                        ' --- Next aggregation step
      lN2 = lN2 * Cl02
   Wend

   If bOdd Then lNdx = lNdxT           ' --- Use temporary list

End Sub

Public Sub LstQSortF(fLst() As Double)

' Order a list of real values
' uses the quick sort method which first sorts pairs and then
' merges successively larger ordered lists

' This routine uses one buffer list that is allocated once
' and then the ordered sublists are kept track of by their
' indices

   Const Cl01 As Long = 1, Cl02 As Long = 2

   Dim fLstT() As Double
   Dim f As Double
   Dim lE1 As Long, lEL As Long, lN2 As Long
   Dim lS1 As Long, lSL As Long, lNS As Long, lMS As Long
   Dim lT1 As Long, lTL As Long
   Dim bOdd As Boolean, bDun As Boolean

   QRS_LibLst.LstBoundF fLst(), lE1, lEL ' - List bounds
   QRS_LibLst.LstAllocF fLstT(), lEL   '     Temporary auxiliary list

   lNS = Cl02                          ' --- Sublist length = 2
   lMS = lEL - Cl01                    ' --- Order all pairs
   For lS1 = lE1 To lMS Step lNS
      lSL = lS1 + Cl01
      If fLst(lS1) > fLst(lSL) Then    ' --- Inverse order?
         f = fLst(lS1)                 '     -> swap
         fLst(lS1) = fLst(lSL)
         fLst(lSL) = f
      End If
   Next lS1

   ' --------------------------------------- Merge the sublists
   '                                         until their length
   '                                         exceeds half the list length

   QRS_Lib0.BitGetMSB lEL, lMS
   bOdd = False                        ' --- Odd true: Copy Lst -> LstT
   lNS = Cl02                          ' --- sublist length
   lN2 = lNS * Cl02                    ' --- two sublist lengths
   While Not lNS > lMS
      bOdd = Not bOdd                  '     Toggle back and forth
      lS1 = lE1                        ' --- Start from beginning
      While Not lS1 + lN2 > lEL        ' --- Process list pairs of regular size
         lT1 = lS1 + lNS               '     2nd sublist start
         lSL = lT1 - Cl01              '     1st sublist end
         lTL = lSL + lNS               '     2nd sublist end
         If bOdd Then                  ' --- Merge
            LstMrgSubF fLst(), lS1, lSL, lT1, lTL, fLstT()
         Else
            LstMrgSubF fLstT(), lS1, lSL, lT1, lTL, fLst()
         End If
         lS1 = lS1 + lN2               ' --- Increment 1st sublist start
      Wend
      lT1 = lS1 + lNS                  '     2nd sublist adjacent
      lSL = lT1 - Cl01                 '     First sublist end
      lTL = lSL + lNS                  '     2nd sublist default end
      If Not lSL < lEL Then            ' --- 1st sublist reaches end
         If bOdd Then                  '     Copy 1st sublist
            For lT1 = lS1 To lEL
               fLstT(lT1) = fLst(lT1)
            Next lT1
         Else
            For lT1 = lS1 To lEL
               fLst(lT1) = fLstT(lT1)
            Next lT1
         End If
      Else                             ' --- 2nd sublist reaches end
         If bOdd Then                  ' --- Merge uneven lists
            LstMrgSubF fLst(), lS1, lSL, lT1, lEL, fLstT()
         Else
            LstMrgSubF fLstT(), lS1, lSL, lT1, lEL, fLst()
         End If
      End If
      lNS = lN2                        ' --- Next aggregation step
      lN2 = lN2 * Cl02
   Wend

   If bOdd Then fLst = fLstT           ' --- Use temporary list

End Sub

Public Sub LstQRNdxL(lLst() As Long, lNdx() As Long)

' Obtain indices by order of long integer values in a list.
'    lNdx(1) = min(lLst())
' uses the quick sort method on a copy of lLst()
' and ordering the index list in parallel
' Value lists and Index lists are doubled
' and copied back or forth on each pass

   Const Cl01 As Long = 1, Cl02 As Long = 2

   Dim lLstS() As Long, lLstT() As Long
   Dim lNdxT() As Long                 ' --- Buffer index list
   Dim lE1 As Long, lEL As Long, lN2 As Long
   Dim lS1 As Long, lSL As Long, lNS As Long, lMS As Long
   Dim lT1 As Long, lTL As Long
   Dim l As Long
   Dim bOdd As Boolean, bDun As Boolean

   QRS_LibLst.LstGet_LL lLst(), lLstS() '    Copy lLst() to lLstT()
   QRS_LibLst.LstBoundL lLstS(), lE1, lEL  ' List bounds
   QRS_LibLst.LstAllocL lLstT(), lEL   '     Buffer values list
   QRS_LibLst.LstAllocL lNdxT(), lEL   '     Buffer index list
   QRS_LibLst.LstAllocL lNdx(), lEL    '     Index list
   For lS1 = lE1 To lEL                '     Index
      lNdx(lS1) = lS1
   Next lS1

   lNS = Cl02                          ' --- Sublist length = 2
   lMS = lEL - Cl01                    ' --- Order all pairs
   For lS1 = lE1 To lMS Step lNS
      lSL = lS1 + Cl01
      If lLstS(lS1) > lLstS(lSL) Then  ' --- Inverse order?
         l = lLstS(lS1)                '     -> swap value
         lLstS(lS1) = lLstS(lSL)
         lLstS(lSL) = l
         l = lNdx(lS1)                 '        swap index
         lNdx(lS1) = lNdx(lSL)
         lNdx(lSL) = l
      End If
   Next lS1

   ' --------------------------------------- Merge the sublists
   '                                         until their length
   '                                         exceeds half the list length

   QRS_Lib0.BitGetMSB lEL, lMS
   bOdd = False                        ' --- Odd true: Copy Lst -> LstT
   lNS = Cl02                          ' --- sublist length
   lN2 = lNS * Cl02                    ' --- two sublist lengths
   While Not lNS > lMS
      bOdd = Not bOdd                  '     Toggle back and forth
      lS1 = lE1                        ' --- Start from beginning
      While Not lS1 + lN2 > lEL        ' --- Process list pairs of regular size
         lT1 = lS1 + lNS               '     2nd sublist start
         lSL = lT1 - Cl01              '     1st sublist end
         lTL = lSL + lNS               '     2nd sublist end
         If bOdd Then                  ' --- Merge
            LstMrgNdxL lLstS(), lNdx(), lS1, lSL, lT1, lTL, lLstT(), lNdxT()
         Else
            LstMrgNdxL lLstT(), lNdxT(), lS1, lSL, lT1, lTL, lLstS(), lNdx()
         End If
         lS1 = lS1 + lN2               ' --- Increment 1st sublist start
      Wend
      lT1 = lS1 + lNS                  '     2nd sublist adjacent
      lSL = lT1 - Cl01                 '     First sublist end
      lTL = lSL + lNS                  '     2nd sublist default end
      If Not lSL < lEL Then            ' --- 1st sublist reaches end
         If bOdd Then                  '     Copy 1st sublist
            For lT1 = lS1 To lEL
               lLstT(lT1) = lLstS(lT1)
               lNdxT(lT1) = lNdx(lT1)
            Next lT1
         Else
            For lT1 = lS1 To lEL
               lLstS(lT1) = lLstT(lT1)
               lNdx(lT1) = lNdxT(lT1)
            Next lT1
         End If
      Else                             ' --- 2nd sublist reaches end
         If bOdd Then                  ' --- Merge uneven lists
            LstMrgNdxL lLstS(), lNdx(), lS1, lSL, lT1, lEL, lLstT(), lNdxT()
         Else
            LstMrgNdxL lLstT(), lNdxT(), lS1, lSL, lT1, lEL, lLstS(), lNdx()
         End If
      End If
      lNS = lN2                        ' --- Next aggregation step
      lN2 = lN2 * Cl02
   Wend

   If bOdd Then lNdx = lNdxT           ' --- Use temporary list

End Sub

Public Sub LstQSortL(lLst() As Long)

' Order a list of long integer values
' uses the quick sort method which first sorts pairs and then
' merges successively larger ordered lists

' This routine uses one buffer list that is allocated once
' and then the ordered sublists are kept track of by their
' indices

   Const Cl01 As Long = 1, Cl02 As Long = 2

   Dim lLstT() As Long
   Dim l As Double
   Dim lE1 As Long, lEL As Long, lN2 As Long
   Dim lS1 As Long, lSL As Long, lNS As Long, lMS As Long
   Dim lT1 As Long, lTL As Long
   Dim bOdd As Boolean, bDun As Boolean

   QRS_LibLst.LstBoundL lLst(), lE1, lEL ' - List bounds
   QRS_LibLst.LstAllocL lLstT(), lEL   '     Temporary auxiliary list

   lNS = Cl02                          ' --- Sublist length = 2
   lMS = lEL - Cl01                    ' --- Order all pairs
   For lS1 = lE1 To lMS Step lNS
      lSL = lS1 + Cl01
      If lLst(lS1) > lLst(lSL) Then    ' --- Inverse order?
         l = lLst(lS1)                 '     -> swap
         lLst(lS1) = lLst(lSL)
         lLst(lSL) = l
      End If
   Next lS1

   ' --------------------------------------- Merge the sublists
   '                                         until their length
   '                                         exceeds half the list length

   QRS_Lib0.BitGetMSB lEL, lMS
   bOdd = False                        ' --- Odd true: Copy Lst -> LstT
   lNS = Cl02                          ' --- sublist length
   lN2 = lNS * Cl02                    ' --- two sublist lengths
   While Not lNS > lMS
      bOdd = Not bOdd                  '     Toggle back and forth
      lS1 = lE1                        ' --- Start from beginning
      While Not lS1 + lN2 > lEL        ' --- Process list pairs of regular size
         lT1 = lS1 + lNS               '     2nd sublist start
         lSL = lT1 - Cl01              '     1st sublist end
         lTL = lSL + lNS               '     2nd sublist end
         If bOdd Then                  ' --- Merge
            LstMrgSubL lLst(), lS1, lSL, lT1, lTL, lLstT()
         Else
            LstMrgSubL lLstT(), lS1, lSL, lT1, lTL, lLst()
         End If
         lS1 = lS1 + lN2               ' --- Increment 1st sublist start
      Wend
      lT1 = lS1 + lNS                  '     2nd sublist adjacent
      lSL = lT1 - Cl01                 '     First sublist end
      lTL = lSL + lNS                  '     2nd sublist default end
      If Not lSL < lEL Then            ' --- 1st sublist reaches end
         If bOdd Then                  '     Copy 1st sublist
            For lT1 = lS1 To lEL
               lLstT(lT1) = lLst(lT1)
            Next lT1
         Else
            For lT1 = lS1 To lEL
               lLst(lT1) = lLstT(lT1)
            Next lT1
         End If
      Else                             ' --- 2nd sublist reaches end
         If bOdd Then                  ' --- Merge uneven lists
            LstMrgSubL lLst(), lS1, lSL, lT1, lEL, lLstT()
         Else
            LstMrgSubL lLstT(), lS1, lSL, lT1, lEL, lLst()
         End If
      End If
      lNS = lN2                        ' --- Next aggregation step
      lN2 = lN2 * Cl02
   Wend

   If bOdd Then lLst = lLstT           ' --- Use temporary list

End Sub

Public Sub LstQRNdxS(sLst() As String, lNdx() As Long, _
                     eCompareType As VbCompareMethod)

' Obtain indices by order of string values in a list.
'    lNdx(1) = min(sLst())
' uses the quick sort method on a copy of sLst()
' and ordering the index list in parallel
' Value lists and Index lists are doubled
' and copied back or forth on each pass

   Const Cl01 As Long = 1, Cl02 As Long = 2

   Dim sLstS() As String, sLstT() As String
   Dim lNdxT() As Long                 ' --- Buffer index list
   Dim s As String
   Dim lE1 As Long, lEL As Long, lN2 As Long
   Dim lS1 As Long, lSL As Long, lNS As Long, lMS As Long
   Dim lT1 As Long, lTL As Long
   Dim l As Long
   Dim e As VbCompareMethod
   Dim bOdd As Boolean, bDun As Boolean

   e = eCompareType

   QRS_LibLst.LstGet_SS sLst(), sLstS() '    Copy sLst() to sLstT()
   QRS_LibLst.LstBoundS sLstS(), lE1, lEL  ' List bounds
   QRS_LibLst.LstAllocS sLstT(), lEL   '     Buffer values list
   QRS_LibLst.LstAllocL lNdxT(), lEL   '     Buffer index list
   QRS_LibLst.LstAllocL lNdx(), lEL    '     Index list
   For lS1 = lE1 To lEL                '     Index
      lNdx(lS1) = lS1
   Next lS1

   lNS = Cl02                          ' --- Sublist length = 2
   lMS = lEL - Cl01                    ' --- Order all pairs
   For lS1 = lE1 To lMS Step lNS
      lSL = lS1 + Cl01                 ' --- Inverse order?
      If StrComp(sLstS(lS1), sLstS(lSL), e) > 0 Then
         s = sLstS(lS1)                '     -> swap value
         sLstS(lS1) = sLstS(lSL)
         sLstS(lSL) = s
         l = lNdx(lS1)                 '        swap index
         lNdx(lS1) = lNdx(lSL)
         lNdx(lSL) = l
      End If
   Next lS1

   ' --------------------------------------- Merge the sublists
   '                                         until their length
   '                                         exceeds half the list length

   QRS_Lib0.BitGetMSB lEL, lMS
   bOdd = False                        ' --- Odd true: Copy Lst -> LstT
   lNS = Cl02                          ' --- sublist length
   lN2 = lNS * Cl02                    ' --- two sublist lengths
   While Not lNS > lMS
      bOdd = Not bOdd                  '     Toggle back and forth
      lS1 = lE1                        ' --- Start from beginning
      While Not lS1 + lN2 > lEL        ' --- Process list pairs of regular size
         lT1 = lS1 + lNS               '     2nd sublist start
         lSL = lT1 - Cl01              '     1st sublist end
         lTL = lSL + lNS               '     2nd sublist end
         If bOdd Then                  ' --- Merge
            LstMrgNdxS sLstS(), lNdx(), lS1, lSL, lT1, lTL, sLstT(), lNdxT(), e
         Else
            LstMrgNdxS sLstT(), lNdxT(), lS1, lSL, lT1, lTL, sLstS(), lNdx(), e
         End If
         lS1 = lS1 + lN2               ' --- Increment 1st sublist start
      Wend
      lT1 = lS1 + lNS                  '     2nd sublist adjacent
      lSL = lT1 - Cl01                 '     First sublist end
      lTL = lSL + lNS                  '     2nd sublist default end
      If Not lSL < lEL Then            ' --- 1st sublist reaches end
         If bOdd Then                  '     Copy 1st sublist
            For lT1 = lS1 To lEL
               sLstT(lT1) = sLstS(lT1)
               lNdxT(lT1) = lNdx(lT1)
            Next lT1
         Else
            For lT1 = lS1 To lEL
               sLstS(lT1) = sLstT(lT1)
               lNdx(lT1) = lNdxT(lT1)
            Next lT1
         End If
      Else                             ' --- 2nd sublist reaches end
         If bOdd Then                  ' --- Merge uneven lists
            LstMrgNdxS sLstS(), lNdx(), lS1, lSL, lT1, lEL, sLstT(), lNdxT(), e
         Else
            LstMrgNdxS sLstT(), lNdxT(), lS1, lSL, lT1, lEL, sLstS(), lNdx(), e
         End If
      End If
      lNS = lN2                        ' --- Next aggregation step
      lN2 = lN2 * Cl02
   Wend

   If bOdd Then lNdx = lNdxT           ' --- Use temporary list

End Sub

Public Sub LstQSortS(sLst() As String, eCompareType As VbCompareMethod)

' Order a list of string values
' uses the quick sort method which first sorts pairs and then
' merges successively larger ordered lists

' This routine uses one buffer list that is allocated once
' and then the ordered sublists are kept track of by their
' indices

   Const Cl01 As Long = 1, Cl02 As Long = 2

   Dim sLstT() As String
   Dim s As Double
   Dim lE1 As Long, lEL As Long, lN2 As Long
   Dim lS1 As Long, lSL As Long, lNS As Long, lMS As Long
   Dim lT1 As Long, lTL As Long
   Dim bOdd As Boolean, bDun As Boolean

   QRS_LibLst.LstBoundS sLst(), lE1, lEL ' - List bounds
   QRS_LibLst.LstAllocS sLstT(), lEL   '     Temporary auxiliary list

   lNS = Cl02                          ' --- Sublist length = 2
   lMS = lEL - Cl01                    ' --- Order all pairs
   For lS1 = lE1 To lMS Step lNS
      lSL = lS1 + Cl01                 ' --- Inverse order?
      If StrComp(sLst(lS1), sLst(lSL), eCompareType) > 0 Then
         s = sLst(lS1)                 '     -> swap
         sLst(lS1) = sLst(lSL)
         sLst(lSL) = s
      End If
   Next lS1

   ' --------------------------------------- Merge the sublists
   '                                         until their length
   '                                         exceeds half the list length

   QRS_Lib0.BitGetMSB lEL, lMS
   bOdd = False                        ' --- Odd true: Copy Lst -> LstT
   lNS = Cl02                          ' --- sublist length
   lN2 = lNS * Cl02                    ' --- two sublist lengths
   While Not lNS > lMS
      bOdd = Not bOdd                  '     Toggle back and forth
      lS1 = lE1                        ' --- Start from beginning
      While Not lS1 + lN2 > lEL        ' --- Process list pairs of regular size
         lT1 = lS1 + lNS               '     2nd sublist start
         lSL = lT1 - Cl01              '     1st sublist end
         lTL = lSL + lNS               '     2nd sublist end
         If bOdd Then                  ' --- Merge
            LstMrgSubS sLst(), lS1, lSL, lT1, lTL, sLstT(), eCompareType
         Else
            LstMrgSubS sLstT(), lS1, lSL, lT1, lTL, sLst(), eCompareType
         End If
         lS1 = lS1 + lN2               ' --- Increment 1st sublist start
      Wend
      lT1 = lS1 + lNS                  '     2nd sublist adjacent
      lSL = lT1 - Cl01                 '     First sublist end
      lTL = lSL + lNS                  '     2nd sublist default end
      If Not lSL < lEL Then            ' --- 1st sublist reaches end
         If bOdd Then                  '     Copy 1st sublist
            For lT1 = lS1 To lEL
               sLstT(lT1) = sLst(lT1)
            Next lT1
         Else
            For lT1 = lS1 To lEL
               sLst(lT1) = sLstT(lT1)
            Next lT1
         End If
      Else                             ' --- 2nd sublist reaches end
         If bOdd Then                  ' --- Merge uneven lists
            LstMrgSubS sLst(), lS1, lSL, lT1, lEL, sLstT(), eCompareType
         Else
            LstMrgSubS sLstT(), lS1, lSL, lT1, lEL, sLst(), eCompareType
         End If
      End If
      lNS = lN2                        ' --- Next aggregation step
      lN2 = lN2 * Cl02
   Wend

   If bOdd Then sLst = sLstT           ' --- Use temporary list

End Sub

Public Sub LstQUniqF(fLst() As Double, fLstUnq() As Double)

' Returns the distinct values in fLst in fLstUnq
' The distinct values are ordered

   Const Cl01 As Long = 1

   Dim lE1 As Long, lEL As Long, lE As Long, lN As Long
   Dim f As Double

   QRS_LibLst.LstBoundF fLst(), lE1, lEL
   QRS_LibLst.LstAllocF fLstUnq(), lEL

   LstQSortF fLst()

   f = fLst(lE1)
   fLstUnq(lE1) = f
   lN = Cl01
   For lE = lE1 + Cl01 To lEL
      If Not fLst(lE) = f Then
         lN = lN + Cl01
         f = fLst(lE)
         fLstUnq(lN) = f
      End If
   Next lE
                                       ' --- Tight fit list
   If lN < lEL Then ReDim Preserve fLstUnq(Cl01 To lN)

End Sub

Public Sub LstQUniqL(lLst() As Long, lLstUnq() As Long)

' Returns the distinct values in lLst in lLstUnq
' The distinct values are ordered

   Const Cl01 As Long = 1

   Dim lE1 As Long, lEL As Long, lE As Long, lN As Long
   Dim l As Long

   QRS_LibLst.LstBoundL lLst(), lE1, lEL
   QRS_LibLst.LstAllocL lLstUnq(), lEL

   LstQSortL lLst()

   l = lLst(lE1)
   lLstUnq(lE1) = l
   lN = Cl01
   For lE = lE1 + Cl01 To lEL
      If Not lLst(lE) = l Then
         lN = lN + Cl01
         l = lLst(lE)
         lLstUnq(lN) = l
      End If
   Next lE
                                       ' --- Tight fit list
   If lN < lEL Then ReDim Preserve lLstUnq(Cl01 To lN)

End Sub

Public Sub LstQUniqS(sLst() As String, sLstUnq() As String, _
                     eCompareType As VbCompareMethod)

' Returns the distinct values in sLst in sLstUnq
' The distinct values are ordered

   Const Cl01 As Long = 1

   Dim lE1 As Long, lEL As Long, lE As Long, lN As Long
   Dim s As String

   QRS_LibLst.LstBoundS sLst(), lE1, lEL
   QRS_LibLst.LstAllocS sLstUnq(), lEL

   LstQSortS sLst(), eCompareType

   s = sLst(lE1)
   sLstUnq(lE1) = s
   lN = Cl01
   For lE = lE1 + Cl01 To lEL
      If Not sLst(lE) = s Then
         lN = lN + Cl01
         s = sLst(lE)
         sLstUnq(lN) = s
      End If
   Next lE
                                       ' --- Tight fit list
   If lN < lEL Then ReDim Preserve sLstUnq(Cl01 To lN)

End Sub


Private Sub LstMrgSubF(fLstA() As Double, _
                       lE1A As Long, lELA As Long, _
                       lE1B As Long, lELB As Long, _
                       fLstC() As Double)

' Merges two ordered sublists of fLstA(),
' specified by their first and last elements,
' into an ordered sublist of fLstC() starting at lE1A
' This routine is used repeatedly by LstQSortF and contains no boundary tests

   Const Cl01 As Long = 1

   Dim fA As Double, fB As Double
   Dim lIA As Long, lIB As Long, lIC As Long
   Dim bAD As Boolean, bBD As Boolean  ' --- "Done" flags
   Dim bA As Boolean                   ' --- take from list A

   lIA = lE1A:   fA = fLstA(lIA)
   lIB = lE1B:   fB = fLstA(lIB)

   For lIC = lE1A To lELB
      bA = bBD                         ' --- fLstB through? -> UseA
      If Not (bAD Or bBD) Then bA = fA < fB
      If bA Then
         fLstC(lIC) = fA
         lIA = lIA + Cl01
         bAD = lIA > lELA
         If Not bAD Then fA = fLstA(lIA)
      Else
         fLstC(lIC) = fB
         lIB = lIB + Cl01
         bBD = lIB > lELB
         If Not bBD Then fB = fLstA(lIB)
      End If
   Next lIC

End Sub

Private Sub LstMrgNdxF(fLstA() As Double, lNdxA() As Long, _
                       lE1A As Long, lELA As Long, _
                       lE1B As Long, lELB As Long, _
                       fLstC() As Double, lNdxC() As Long)

' Merges two ordered sublists of fLstA() along with sublists of lNdxA()
' specified by their first and last elements,
' into an ordered sublist of fLstC() along with lNdxC(), starting at lE1A
' This routine is used repeatedly by LstQRankF and contains no boundary tests

   Const Cl01 As Long = 1

   Dim fA As Double, fB As Double
   Dim lIA As Long, lIB As Long, lIC As Long
   Dim bAD As Boolean, bBD As Boolean  ' --- "Done" flags
   Dim bA As Boolean                   ' --- take from list A

   lIA = lE1A:   fA = fLstA(lIA)
   lIB = lE1B:   fB = fLstA(lIB)

   For lIC = lE1A To lELB
      bA = bBD                         ' --- fLstB through? -> UseA
      If Not (bAD Or bBD) Then bA = fA < fB
      If bA Then
         fLstC(lIC) = fA
         lNdxC(lIC) = lNdxA(lIA)
         lIA = lIA + Cl01
         bAD = lIA > lELA
         If Not bAD Then fA = fLstA(lIA)
      Else
         fLstC(lIC) = fB
         lNdxC(lIC) = lNdxA(lIB)
         lIB = lIB + Cl01
         bBD = lIB > lELB
         If Not bBD Then fB = fLstA(lIB)
      End If
   Next lIC

End Sub

Private Sub LstMrgSubL(lLstA() As Long, _
                       lE1A As Long, lELA As Long, _
                       lE1B As Long, lELB As Long, _
                       lLstC() As Long)

' Merges two ordered sublists of lLstA(),
' specified by their first and last elements,
' into an ordered sublist of lLstC() starting at lE1A
' This routine is used repeatedly by LstQSortF and contains no boundary tests

   Const Cl01 As Long = 1

   Dim lA As Long, lB As Long
   Dim lIA As Long, lIB As Long, lIC As Long
   Dim bAD As Boolean, bBD As Boolean  ' --- "Done" flags
   Dim bA As Boolean                   ' --- take from list A

   lIA = lE1A:   lA = lLstA(lIA)
   lIB = lE1B:   lB = lLstA(lIB)

   For lIC = lE1A To lELB
      bA = bBD                         ' --- fLstB through? -> UseA
      If Not (bAD Or bBD) Then bA = lA < lB
      If bA Then
         lLstC(lIC) = lA
         lIA = lIA + Cl01
         bAD = lIA > lELA
         If Not bAD Then lA = lLstA(lIA)
      Else
         lLstC(lIC) = lB
         lIB = lIB + Cl01
         bBD = lIB > lELB
         If Not bBD Then lB = lLstA(lIB)
      End If
   Next lIC

End Sub

Private Sub LstMrgNdxL(lLstA() As Long, lNdxA() As Long, _
                       lE1A As Long, lELA As Long, _
                       lE1B As Long, lELB As Long, _
                       lLstC() As Long, lNdxC() As Long)

' Merges two ordered sublists of fLstA() along with sublists of lNdxA()
' specified by their first and last elements,
' into an ordered sublist of fLstC() along with lNdxC(), starting at lE1A
' This routine is used repeatedly by LstQRankF and contains no boundary tests

   Const Cl01 As Long = 1

   Dim lA As Long, lB As Long
   Dim lIA As Long, lIB As Long, lIC As Long
   Dim bAD As Boolean, bBD As Boolean  ' --- "Done" flags
   Dim bA As Boolean                   ' --- take from list A

   lIA = lE1A:   lA = lLstA(lIA)
   lIB = lE1B:   lB = lLstA(lIB)

   For lIC = lE1A To lELB
      bA = bBD                         ' --- fLstB through? -> UseA
      If Not (bAD Or bBD) Then bA = lA < lB
      If bA Then
         lLstC(lIC) = lA
         lNdxC(lIC) = lNdxA(lIA)
         lIA = lIA + Cl01
         bAD = lIA > lELA
         If Not bAD Then lA = lLstA(lIA)
      Else
         lLstC(lIC) = lB
         lNdxC(lIC) = lNdxA(lIB)
         lIB = lIB + Cl01
         bBD = lIB > lELB
         If Not bBD Then lB = lLstA(lIB)
      End If
   Next lIC

End Sub

Private Sub LstMrgSubS(sLstA() As String, _
                       lE1A As Long, lELA As Long, _
                       lE1B As Long, lELB As Long, _
                       sLstC() As String, _
                       eCompareType As VbCompareMethod)

' Merges two ordered sublists of sLstA(),
' specified by their first and last elements,
' into an ordered sublist of sLstC() starting at lE1A
' This routine is used repeatedly by LstQSortF and contains no boundary tests

   Const Cl01 As Long = 1

   Dim sA As String, sB As String
   Dim lIA As Long, lIB As Long, lIC As Long
   Dim bAD As Boolean, bBD As Boolean  ' --- "Done" flags
   Dim bA As Boolean                   ' --- take from list A

   lIA = lE1A:   sA = sLstA(lIA)
   lIB = lE1B:   sB = sLstA(lIB)

   For lIC = lE1A To lELB
      bA = bBD                         ' --- fLstB through? -> UseA
      If Not (bAD Or bBD) Then bA = StrComp(sA, sB, eCompareType) < 0
      If bA Then
         sLstC(lIC) = sA
         lIA = lIA + Cl01
         bAD = lIA > lELA
         If Not bAD Then sA = sLstA(lIA)
      Else
         sLstC(lIC) = sB
         lIB = lIB + Cl01
         bBD = lIB > lELB
         If Not bBD Then sB = sLstA(lIB)
      End If
   Next lIC

End Sub

Private Sub LstMrgNdxS(sLstA() As String, lNdxA() As Long, _
                       lE1A As Long, lELA As Long, _
                       lE1B As Long, lELB As Long, _
                       sLstC() As String, lNdxC() As Long, _
                       eCompareType As VbCompareMethod)

' Merges two ordered sublists of fLstA() along with sublists of lNdxA()
' specified by their first and last elements,
' into an ordered sublist of fLstC() along with lNdxC(), starting at lE1A
' This routine is used repeatedly by LstQRankF and contains no boundary tests

   Const Cl01 As Long = 1

   Dim sA As String, sB As String
   Dim lIA As Long, lIB As Long, lIC As Long
   Dim bAD As Boolean, bBD As Boolean  ' --- "Done" flags
   Dim bA As Boolean                   ' --- take from list A

   lIA = lE1A:   sA = sLstA(lIA)
   lIB = lE1B:   sB = sLstA(lIB)

   For lIC = lE1A To lELB
      bA = bBD                         ' --- fLstB through? -> UseA
      If Not (bAD Or bBD) Then bA = StrComp(sA, sB, eCompareType) < 0
      If bA Then
         sLstC(lIC) = sA
         lNdxC(lIC) = lNdxA(lIA)
         lIA = lIA + Cl01
         bAD = lIA > lELA
         If Not bAD Then sA = sLstA(lIA)
      Else
         sLstC(lIC) = sB
         lNdxC(lIC) = lNdxA(lIB)
         lIB = lIB + Cl01
         bBD = lIB > lELB
         If Not bBD Then sB = sLstA(lIB)
      End If
   Next lIC

End Sub

