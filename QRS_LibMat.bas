Attribute VB_Name = "QRS_LibMat"
Option Explicit

' Module : QRS_LibMat
' Purpose: Matrix functions
' By     : QRS, Roger Strebel, based upon different old and new sources
'          QRS matrices all have 1-based indices
' Date   : 17.03.2018                  Matrix memory handling
'          18.03.2018                  Matrix operations
' --- The public interface
'     MtxAddF                          Add fA and fB into fC         18.03.2018
'     MtxChkF                          Cholesky decompose fA
'     MtxDupF                          Duplicate fA into fB          18.03.2018
'     MtxIdnF                          Identity matrix               0
'     MtxIniF                          initialize real value matrix  17.03.2018
'     MtxIniL                          initialize long int matrix    17.03.2018
'     MtxInvF                          Inversion with pivot search
'     MtxMulF                          Multiply fA and fB into fC    18.03.2018
'     MtxNrmF                          Frobenius norm of fA
'     MtxSclF                          Matrix scalar multiplication  18.03.2018
'     MtxSizF                          Bounds of real value matrix   17.03.2018
'     MtxSizL                          Bounds of long integer matrix 17.03.2018
'     MtxSubF                          Subtract fB from fA into fC   18.03.2018
'     MtxTpsF                          Transpose Real value Matrix   18.03.2018
'     MtxTpsL                          Transpose Long integer Matrix 18.03.2018
'     MtxTrcF                          Diagonal trace of Matrix      03.04.2018
'     MtxTriF                          LU decomposition
'     MtxUseL                          Use Long integer matrix?      17.03.2018
'     MtxUseF                          Use real value matrix?        17.03.2018
' --- The private sphere

Public Const MCl_LibMat_X1 As Long = 1

Public Sub MtxAddF(fA() As Double, fB() As Double, fC() As Double)

' Adds fA and fB into fC if dimensions of fA and fB match
' Also works if either fA() or fB() is used as fC()

   Dim lMRow As Long, lMCol As Long
   Dim lIRow As Long, lICol As Long

   MtxSizF fA(), lMRow, lMCol
   If Not MtxUseF(fB(), lMRow, lMCol) Then Exit Sub
   If Not MtxUseF(fC(), lMRow, lMCol) Then
      ReDim fC(MCl_LibMat_X1 To lMRow, MCl_LibMat_X1 To lMCol) As Double
   End If

   For lIRow = MCl_LibMat_X1 To lMRow  ' --- Add element by element
      For lICol = MCl_LibMat_X1 To lMCol
         fC(lIRow, lICol) = fA(lIRow, lICol) + fB(lIRow, lICol)
      Next lICol
   Next lIRow

End Sub

Public Function MtxChkF(fA() As Double, fB() As Double) As Boolean

' Calculates the Cholesky LU decomposition of fA() into fB()
' The Cholesky decomposition only exists if fA() is positive semidefinite
' If the decomposition does not exist, returns true

   Dim lMRow As Long, lMCol As Long
   Dim bFail As Boolean

   MtxSizF fA(), lMRow, lMCol


Ende:

   MtxChkF = bFail

End Function

Public Sub MtxDupF(fA() As Double, fB() As Double)

' Duplicates fA() into fB() by re-allocation and elementwise copying
   
   Dim lMRow As Long, lMCol As Long
   Dim lIRow As Long, lICol As Long

   MtxSizF fA(), lMRow, lMCol
   If Not MtxUseF(fB(), lMRow, lMCol) Then
      ReDim fB(MCl_LibMat_X1 To lMRow, MCl_LibMat_X1 To lMCol) As Double
   End If

   For lIRow = MCl_LibMat_X1 To lMRow  ' --- Add element by element
      For lICol = MCl_LibMat_X1 To lMCol
         fB(lIRow, lICol) = fA(lIRow, lICol)
      Next lICol
   Next lIRow

End Sub

Public Sub MtxIdnF(lSiz As Long, fMat() As Double)

' Returns the identity matrix of dimension lSiz

   Const Cf01 As Double = 1#

   Dim lI As Long

   MtxIniF fMat(), lSiz, lSiz, 0

   For lI = 1 To lSiz
      fMat(lI, lI) = Cf01
   Next lI

End Sub

Public Sub MtxIniF(fMat() As Double, lNRow As Long, lNCol As Long, _
                   Optional fIni As Double = 0)

' Allocate or re-allocate to specified dimensions a real value matrix
' and set to initial value

   Dim lIRow As Long, lICol As Long

   If Not MtxUseF(fMat(), lNRow, lNCol) Then
      ReDim fMat(MCl_LibMat_X1 To lNRow, MCl_LibMat_X1 To lNCol) As Double
   End If

   For lIRow = MCl_LibMat_X1 To lNRow
      For lICol = MCl_LibMat_X1 To lNCol
         fMat(lIRow, lICol) = fIni
      Next lICol
   Next lIRow

End Sub

Public Sub MtxIniL(lMat() As Long, lNRow As Long, lNCol As Long, _
                   Optional lIni As Long = 0)

' Allocate or re-allocate to specified dimensions a long integer matrix
' and set to initial value

   Dim lIRow As Long, lICol As Long

   If Not MtxUseL(lMat(), lNRow, lNCol) Then
      ReDim lMat(MCl_LibMat_X1 To lNRow, MCl_LibMat_X1 To lNCol) As Long
   End If

   For lIRow = MCl_LibMat_X1 To lNRow
      For lICol = MCl_LibMat_X1 To lNCol
         lMat(lIRow, lICol) = lIni
      Next lICol
   Next lIRow

End Sub

Public Function MtxInvF(fA() As Double, fB() As Double) As Boolean

' Returns the inverse of Matrix fA() in fB()

   Const Cl01 As Long = 1

   Dim fT() As Double
   Dim fD As Double, fF As Double, f1 As Double
   Dim lIRow As Long, lMRow As Long
   Dim lICol As Long, lMCol As Long
   Dim lIRem As Long, l1 As Long
   Dim bFail As Boolean

   MtxSizF fA(), lMRow, lMCol
   bFail = Not lMRow = lMCol
   If bFail Then GoTo Ende

   f1 = 1#
   l1 = MCl_LibMat_X1

   MtxIdnF lMRow, fB()                 ' --- Identity matrix in fB()
   MtxDupF fA(), fT()                  ' --- Duplicate fA() in fT()

   For lIRow = l1 To lMRow - Cl01
      fD = f1 / fT(lIRow, lIRow)       '     diagonal element
      For lIRem = lIRow + Cl01 To lMRow '    all remaining rows
         If fT(lIRow, lIRow) = 0 Then fF = 0 Else fF = fT(lIRem, lIRow) * fD
                                       '     All elements in the row
         For lICol = l1 To lMCol       '     LU decomposition of fT() and fB()
            fT(lIRem, lICol) = fT(lIRem, lICol) - fT(lIRow, lICol) * fF
            fB(lIRem, lICol) = fB(lIRem, lICol) - fB(lIRow, lICol) * fF
         Next lICol
      Next lIRem
   Next lIRow

   For lIRow = lMRow To MCl_LibMat_X1 + Cl01 Step -Cl01
      For lIRem = l1 To 2
      Next lIRem
   Next lIRow

Ende:

End Function

Public Sub MtxMulF(fA() As Double, fB() As Double, fC() As Double)

' multiplies fA and fB into fC if dimensions of fA and transposed fB match
' Also works if either fA() or fB() is used as fC()

   Dim f As Double
   Dim lMRow As Long, lMCol As Long
   Dim lIRow As Long, lICol As Long, lIX As Long

   MtxSizF fA(), lMRow, lMCol
   If Not MtxUseF(fB(), lMCol, lMRow) Then Exit Sub
   If Not MtxUseF(fC(), lMRow, lMRow) Then
      ReDim fC(MCl_LibMat_X1 To lMRow, MCl_LibMat_X1 To lMRow) As Double
   End If

   For lIRow = MCl_LibMat_X1 To lMRow  ' --- Add element by element
      For lICol = MCl_LibMat_X1 To lMRow
         f = 0
         For lIX = MCl_LibMat_X1 To lMCol
            f = f + fA(lIRow, lIX) * fB(lIX, lICol)
         Next lIX
         fC(lIRow, lICol) = f
      Next lICol
   Next lIRow

End Sub

Public Function MtxNorF(fA() As Double) As Double

' Returns the frobenius (euclidean) norm of matrix fA()
' which is the sum of squares of all elements

   Dim fE As Double, fS As Double
   Dim lIRow As Long, lMRow As Long
   Dim lICol As Long, lMCol As Long, l1 As Long

   MtxSizF fA(), lMRow, lMCol

   l1 = MCl_LibMat_X1
   For lIRow = l1 To lMRow
      For lICol = l1 To lMCol
         fE = fA(lIRow, lICol)
         fS = fS + fE * fE
      Next lICol
   Next lIRow

   MtxNorF = fS

End Function

Public Sub MtxSclF(fA() As Double, fS As Double, fB() As Double)

' Multiplies all elements of fA() with fS into fB

   Dim lMRow As Long, lMCol As Long
   Dim lIRow As Long, lICol As Long

   MtxSizF fA(), lMRow, lMCol
   If Not MtxUseF(fB(), lMRow, lMRow) Then
      ReDim fC(MCl_LibMat_X1 To lMRow, MCl_LibMat_X1 To lMCol) As Double
   End If

   For lIRow = MCl_LibMat_X1 To lMRow  ' --- Scalar multiplication
      For lICol = MCl_LibMat_X1 To lMCol
         fB(lIRow, lICol) = fA(lIRow, lICol) * fS
      Next lICol
   Next lIRow

End Sub

Public Sub MtxSizF(fMat() As Double, _
                   Optional lNRow As Long = 0, Optional lNCol As Long = 0)

' Returns the bounds of the real value matrix in optional return arguments

   If ArrIsAllF(fMat()) Then
      lNRow = UBound(fMat(), 1) + MCl_LibMat_X1 - LBound(fMat(), 1)
      lNCol = UBound(fMat(), 2) + MCl_LibMat_X1 - LBound(fMat(), 2)
   End If

End Sub

Public Sub MtxSizL(lMat() As Long, _
                   Optional lNRow As Long = 0, Optional lNCol As Long = 0)

' Returns the bounds of the long int matrix in optional return arguments

   If ArrIsAllL(lMat()) Then
      lNRow = UBound(lMat(), 1) + MCl_LibMat_X1 - LBound(lMat(), 1)
      lNCol = UBound(lMat(), 2) + MCl_LibMat_X1 - LBound(lMat(), 2)
   End If

End Sub

Public Sub MtxSubF(fA() As Double, fB() As Double, fC() As Double)

' Subtractrs fB from fA into fC if dimensions of fA and fB match
' Also works if either fA() or fB() is used as fC()

   Dim lMRow As Long, lMCol As Long
   Dim lIRow As Long, lICol As Long

   MtxSizF fA(), lMRow, lMCol
   If Not MtxUseF(fB(), lMRow, lMCol) Then Exit Sub
   If Not MtxUseF(fC(), lMRow, lMCol) Then
      ReDim fC(MCl_LibMat_X1 To lMRow, MCl_LibMat_X1 To lMCol) As Double
   End If

   For lIRow = MCl_LibMat_X1 To lMRow  ' --- Subtract element by element
      For lICol = MCl_LibMat_X1 To lMCol
         fC(lIRow, lICol) = fA(lIRow, lICol) - fB(lIRow, lICol)
      Next lICol
   Next lIRow

End Sub

Public Sub MtxTpsF(fA() As Double, fB() As Double)

' Returns the transposed matrix of fA() in fB()
' The target matrix fB() is allocated or re-allocated
' Does not work if the reference of fB() refers to fA()

   Dim lMRow As Long, lMCol As Long
   Dim lIRow As Long, lICol As Long

   MtxSizF fA(), lMRow, lMCol          ' --- Source size
   If Not MtxUseF(fB(), lMCol, lMRow) Then ' Target match transposed size?
      ReDim fB(MCl_LibMat_X1 To lMCol, MCl_LibMat_X1 To lMRow) As Double
   End If

   For lIRow = MCl_LibMat_X1 To lMRow  ' --- Copy
      For lICol = MCl_LibMat_X1 To lMCol
         fB(lICol, lIRow) = fA(lIRow, lICol)
      Next lICol
   Next lIRow

End Sub

Public Sub MtxTpsL(lA() As Long, lB() As Long)

' Returns the transposed matrix of lA() in lB()
' The target matrix lB() is allocated or re-allocated
' Does not work if the reference of lB() refers to lA()

   Dim lMRow As Long, lMCol As Long
   Dim lIRow As Long, lICol As Long

   MtxSizL lA(), lMRow, lMCol
   If Not MtxUseL(lB(), lMRow, lMCol) Then
      ReDim lB(MCl_LibMat_X1 To lMCol, MCl_LibMat_X1 To lMRow) As Long
   End If

   For lIRow = MCl_LibMat_X1 To lMRow
      For lICol = MCl_LibMat_X1 To lMCol
         lB(lICol, lIRow) = lA(lIRow, lICol)
      Next lICol
   Next lIRow

End Sub

Public Sub MtxTrcF(fMat() As Double, fTrc() As Double)

' Returns the trace vector of fMat() in fTrc() as 1D list

   Dim lSiz As Long, lNdx As Long, l1 As Long

   l1 = MCl_LibMat_X1                  ' --- Low index

   MtxSizF fMat(), lSiz, lNdx          ' --- 2D size
   If lNdx < lSiz Then lSiz = lNdx     ' --- lower

   ReDim fTrc(l1 To lSiz)

   For lNdx = l1 To lSiz
      fTrc(lNdx) = fMat(lNdx, lNdx)
   Next lNdx

End Sub

Public Function MtxTriF(fA() As Double, fB() As Double, _
                        Optional bPivotSrch As Boolean = False) As Boolean

' LU decomposition of fA() in fB() using Gauss elimination
' with pivot search "close to 1"
' If decomposition successful, diagonal elements contain eigenvalues
' Returns true if fA() is singular
' Note: If bPivotSrch is set, the determinant may change sign

End Function

Public Function MtxUseF(fMat() As Double, _
                        Optional lNRow As Long = 0, _
                        Optional lNCol As Long = 0) As Boolean

' Returns true if real value matrix has been allocated
' and its dimensions match specified row and column count
' Dimension check is omitted for dimensions specified at zero

   Dim lMRow As Long, lMCol As Long
   Dim bUse As Boolean

   bUse = Not Not fMat()               ' --- True if allocated
   If bUse Then
      MtxSizF fMat(), lMRow, lMCol
      If lNRow > 0 Then bUse = lMRow = lNRow
      If lNCol > 0 Then bUse = lMCol = lNCol
   End If
   MtxUseF = bUse

End Function

Public Function MtxUseL(lMat() As Long, _
                        Optional lNRow As Long = 0, _
                        Optional lNCol As Long = 0) As Boolean

' Returns true if long integer matrix has been allocated
' and its dimensions match specified row and column count
' Dimension check is omitted for dimensions specified at zero

   Dim lMRow As Long, lMCol As Long
   Dim bUse As Boolean

   bUse = Not Not lMat()
   If bUse Then
      MtxSizL lMat(), lMRow, lMCol
      If lNRow > 0 Then bUse = lMRow = lNRow
      If lNCol > 0 Then bUse = lMCol = lNCol
   End If
   MtxUseL = bUse

End Function

