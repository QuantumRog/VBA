Attribute VB_Name = "QRS_LibFnd"
Option Explicit

' Module : QRS_LibFnd
' Purpose: Functions for finding values in tables
'          Uses auxiliary functions of QRS_LibArr
' By     : QRS, Roger Strebel
' Date   : 28.01.2019                  ColFind1 functions created
'          17.05.2019                  numeric functions added
' --- The public interface
'     ColFind1_F                       Find 1 double  in 1 column    28.01.2019
'     ColFind1_L                       Find 1 integer in 1 column    28.01.2019
'     ColFind1_S                       Find 1 string  in 1 column    28.01.2019
'     ColFind1_V                       Find 1 variant in 1 column    28.01.2019

Public Function ColFind1_F(fVal As Double, fArr() As Double, _
                           lCoLook As Long, _
                           Optional lRow1 As Long = 0) As Long

' Find 1 variant value in a column of vArr() and return index
' If lRow1 is set to 0, uses the first row
' Returns -1 if not found

   Dim lR1 As Long, lRL As Long, lC1 As Long, lCL As Long

   QRS_LibArr.ArrBoundF fArr(), lR1, lRL, lC1, lCL
   lRow1 = QRS_LibArr.NdxGetSX1(lR1, lRL, lRow1)
   lC1 = QRS_LibArr.NdxGetSX1(lC1, lCL, lCoLook)

   For lR1 = lRow1 To lRL
      If fArr(lR1, lC1) = fVal Then Exit For
   Next lR1
   If lR1 > lRL Then lR1 = -1
   ColFind1_F = lR1

End Function

Public Function ColFind1_L(lVal As Long, lArr() As Long, _
                           lCoLook As Long, _
                           Optional lRow1 As Long = 0) As Long

' Find 1 variant value in a column of vArr() and return index
' If lRow1 is set to 0, uses the first row
' Returns -1 if not found

   Dim lR1 As Long, lRL As Long, lC1 As Long, lCL As Long

   QRS_LibArr.ArrBoundL lArr(), lR1, lRL, lC1, lCL
   lRow1 = QRS_LibArr.NdxGetSX1(lR1, lRL, lRow1)
   lC1 = QRS_LibArr.NdxGetSX1(lC1, lCL, lCoLook)

   For lR1 = lRow1 To lRL
      If lArr(lR1, lC1) = lVal Then Exit For
   Next lR1
   If lR1 > lRL Then lR1 = -1
   ColFind1_L = lR1

End Function

Public Function ColFind1_S(sVal As String, sArr() As String, _
                           lCoLook As Long, _
                           Optional lRow1 As Long = 0) As Long

' Find 1 variant value in a column of vArr() and return index
' If lRow1 is set to 0, uses the first row
' Returns -1 if not found

   Dim lR1 As Long, lRL As Long, lC1 As Long, lCL As Long

   QRS_LibArr.ArrBoundS sArr(), lR1, lRL, lC1, lCL
   lRow1 = QRS_LibArr.NdxGetSX1(lR1, lRL, lRow1)
   lC1 = QRS_LibArr.NdxGetSX1(lC1, lCL, lCoLook)

   For lR1 = lRow1 To lRL
      If StrComp(sArr(lR1, lC1), sVal, vbTextCompare) = 0 Then Exit For
   Next lR1
   If lR1 > lRL Then lR1 = -1
   ColFind1_S = lR1

End Function

Public Function ColFind1_V(vVal As Variant, vArr() As Variant, _
                           lCoLook As Long, _
                           Optional lRow1 As Long = 0) As Long

' Find 1 variant value in a column of vArr() and return index
' If lRow1 is set to 0, uses the first row
' Returns -1 if not found

   Dim lR1 As Long, lRL As Long, lC1 As Long, lCL As Long

   QRS_LibArr.ArrBoundV vArr(), lR1, lRL, lC1, lCL
   lRow1 = QRS_LibArr.NdxGetSX1(lR1, lRL, lRow1)
   lC1 = QRS_LibArr.NdxGetSX1(lC1, lCL, lCoLook)

   For lR1 = lRow1 To lRL
      If vArr(lR1, lC1) = vVal Then Exit For
   Next lR1
   If lR1 > lRL Then lR1 = -1
   ColFind1_V = lR1

End Function


