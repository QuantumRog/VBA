VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QRS_LUT_XY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Module : LUT_XY
' Purpose: A look-up table for two integer values
'          3 columns: Key, X, Y
' By     : QRS GmbH, Roger Strebel
' Date   : 01.11.2018
' --- The public interface
' Properties
'     Count       (get)                Number of entries
'     GetKey      (get)                Look up key from x and y
' Methods
'     Get_XY                           Look up x and y from key
'     InitFromArrColL                  Initialize from 3 array columns
' --- The private sphere
'

Private MlCount As Long
Private MlKXY() As Long

Public Property Get Count() As Long

   Count = MlCount

End Property

Public Property Get GetKey(lValX As Long, lValY As Long) As Long

   Dim lI As Long
   Dim bMatch As Boolean

   For lI = 1 To MlCount
      bMatch = MlKXY(lI, 2) = lValX And MlKXY(lI, 3) = lValY
      If bMatch Then Exit For
   Next lI
   If bMatch Then GetKey = MlKXY(lI, 1)

End Property

Public Sub GetRow(lRow As Long, _
                      lValX As Long, lValY As Long, lKey As Long)

   If lRow < 1 Or lRow > MlCount Then Exit Sub

   lKey = MlKXY(lRow, 1)
   lValX = MlKXY(lRow, 2)
   lValY = MlKXY(lRow, 3)

End Sub

Public Sub Get_XY(lKey As Long, lValX As Long, lValY As Long)

   Dim lI As Long
   Dim bMatch As Boolean

   For lI = 1 To MlCount
      bMatch = MlKXY(lI, 1) = lKey
      If bMatch Then Exit For
   Next lI
   If bMatch Then
      lValX = MlKXY(lI, 2)
      lValY = MlKXY(lI, 3)
   End If

End Sub

Public Sub InitFromArrColL(lArr() As Long, _
                           Optional lColKey As Long = 1, _
                           Optional lColX As Long = 2, _
                           Optional lColY As Long = 3)

   Dim lCol() As Long
   Dim b123 As Boolean

   If Not QRS_LibArr.ArrIsAllL(lArr()) Then Exit Sub

   b123 = lColY - lColX = 1 And lColX - lColKey = 1
   MlCount = UBound(lArr(), 1) + 1 - LBound(lArr(), 1)
   If b123 Then
      QRS_LibArr.GetColLonLon lArr, MlKXY(), lColKey, lColY
   Else
      QRS_LibArr.ArrAllocL MlKXY(), MlCount, 3
      QRS_LibA2L.GetColLonLon lArr(), lCol(), lColKey
      QRS_LibA2L.PutColLonLon lCol(), lArr(), 1
      QRS_LibA2L.GetColLonLon lArr(), lCol(), lColX
      QRS_LibA2L.PutColLonLon lCol(), lArr(), 2
      QRS_LibA2L.GetColLonLon lArr(), lCol(), lColY
      QRS_LibA2L.PutColLonLon lCol(), lArr(), 3
   End If
End Sub

