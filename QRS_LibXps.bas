Attribute VB_Name = "QRS_LibXps"
Option Explicit

' Module : QRS_LibXps
' Project: any
' Purpose: VBA array transposition
' By     : QRS, Roger Strebel
' Date   : 17.05.2019
' --- The public interface
'     ArrXpsDbl                        Transpose array of double     17.05.2019
'     ArrXpsLon                        Transpose array of integer    17.05.2019
'     ArrXpsStr                        Transpose array of string
'     ArrXpsVar                        Transpose array of variant    17.05.2029
' --- The private sphere

Public Sub ArrXpsDbl(fArrSrc() As Double, _
                     fArrDst() As Double)

' A proper array transposition using a temporary array
' so that vArrSrc and vArrDst can point to the same array

   Dim fArrTmp() As Double
   Dim lR1 As Long, lRL As Long, lRI As Long
   Dim lC1 As Long, lCL As Long, lCI As Long

   If Not QRS_LibArr.ArrIsAllF(fArrSrc()) Then Exit Sub
   QRS_LibArr.GetSubDblDbl fArrSrc(), fArrTmp()
   QRS_LibArr.ArrBoundF fArrTmp(), lR1, lRL, lC1, lCL
   QRS_LibArr.ArrAllocF fArrDst(), lCL, lRL
   For lCI = lC1 To lCL
      For lRI = lR1 To lRL
         fArrDst(lCI, lRI) = fArrTmp(lRI, lCI)
      Next lRI
   Next lCI

End Sub

Public Sub ArrXpsLon(lArrSrc() As Long, _
                     lArrDst() As Long)

' A proper array transposition using a temporary array
' so that vArrSrc and vArrDst can point to the same array

   Dim lArrTmp() As Long
   Dim lR1 As Long, lRL As Long, lRI As Long
   Dim lC1 As Long, lCL As Long, lCI As Long

   If Not QRS_LibArr.ArrIsAllL(lArrSrc()) Then Exit Sub
   QRS_LibArr.GetSubLonLon lArrSrc(), lArrTmp()
   QRS_LibArr.ArrBoundL lArrTmp(), lR1, lRL, lC1, lCL
   QRS_LibArr.ArrAllocL lArrDst(), lCL, lRL
   For lCI = lC1 To lCL
      For lRI = lR1 To lRL
         lArrDst(lCI, lRI) = lArrTmp(lRI, lCI)
      Next lRI
   Next lCI

End Sub

Public Sub ArrXpsStr(sArrSrc() As String, _
                     sArrDst() As String)

   Dim sArrTmp() As String
   Dim lR1 As Long, lRL As Long, lRI As Long
   Dim lC1 As Long, lCL As Long, lCI As Long

   If Not QRS_LibArr.ArrIsAllS(sArrSrc()) Then Exit Sub
   QRS_LibArr.GetSubStrStr sArrSrc(), sArrTmp()
   QRS_LibArr.ArrBoundS sArrTmp(), lR1, lRL, lC1, lCL
   QRS_LibArr.ArrAllocS sArrDst(), lCL, lRL
   For lCI = lC1 To lCL
      For lRI = lR1 To lRL
         sArrDst(lCI, lRI) = sArrTmp(lRI, lCI)
      Next lRI
   Next lCI

End Sub

Public Sub ArrXpsVar(vArrSrc() As Variant, _
                     vArrDst() As Variant)

' A proper array transposition using a temporary array
' so that vArrSrc and vArrDst can point to the same array

   Dim vArrTmp() As Variant
   Dim lR1 As Long, lRL As Long, lRI As Long
   Dim lC1 As Long, lCL As Long, lCI As Long

   If Not QRS_LibArr.ArrIsAllV(vArrSrc()) Then Exit Sub
   QRS_LibArr.GetSubVarVar vArrSrc(), vArrTmp()
   QRS_LibArr.ArrBoundV vArrTmp(), lR1, lRL, lC1, lCL
   QRS_LibArr.ArrAllocV vArrDst(), lCL, lRL
   For lCI = lC1 To lCL
      For lRI = lR1 To lRL
         vArrDst(lCI, lRI) = vArrTmp(lRI, lCI)
      Next lRI
   Next lCI

End Sub

