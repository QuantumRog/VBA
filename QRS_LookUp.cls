VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QRS_LookUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Module : LookUp
' Purpose: a 2D table contains a key column and a value column
'          Somethimes the key corresponding to a value must be
'          looked up, sometimes the value corresponding to the
'          key. Of both columns there are ordred copies as 1D-
'          lists for quick look-up
' Idea   : A look-up table for two columns in a table of strings
'          Quick look-up using ordered auxiliary lists and index
'          on both columns
' Use    :  1. Prepare the 2D table
'           2. Instantiate the module
'           3. Init module using SetTbl method
'           4. Look up value from key using GetVal method
'           5. Look up key from value using GetKey method
' By     : QRS GmbH, Roger Strebel
' Date   : 06.02.2019
' --- The public interface
' Properties
'     ValLst   (get     )              Extract ordered values list
' Methods
'     GetVal                           Look up value from key
'     GetKey                           Look up key from value
'     SetTbl                           Set table and init look-up
' --- The private sphere
'     ColOrder                         Build ordered lists

Private MsLUT() As String
Private MsKey() As String              ' --- Ordered key list
Private MlKey() As Long                ' --- Key indices
Private MsVal() As String              ' --- Ordered value list
Private MlVal() As Long                ' --- Value indices

Private MlColKey As Long
Private MlColVal As Long
Private MlLen As Long

Public Property Get ValLst() As String()

' Returns the ordered values column as a 1D list
' Call syntax:  sColLst() = oInstance.KeyVal()

   If Not MlLen > 0 Then Exit Property
   KeyVal = MsVal()

End Property

Public Function GetVal(sKey As String, _
                       Optional lNdx As Long = 0) As String

' Returns the value corresponding to the key provided

   lNdx = QRS_LibQRS.LstQFindS(MsKey(), sKey)
   If lNdx > 0 Then
      GetVal = MsLUT(MlKey(lNdx), MlColVal)
   End If

End Function

Public Function GetKey(sVal As String, _
                       Optional lNdx As Long = 0) As String

' Returns the key corresponding to the row
' containing sVal in the value column

   lNdx = QRS_LibQRS.LstQFindS(MsVal(), sVal)
   If lNdx > 0 Then
      GetKey = MsLUT(MlVal(lNdx), MlColKey)
   End If

End Function

Public Sub SetTbl(sTbl() As String, lColKey As Long, lColVal As Long)

' Sets the table and initalizes the value and key look-ups

   MlColKey = lColKey
   MlColVal = lColVal
   MsLUT() = sTbl()
   MlLen = UBound(sTbl(), 1)
   ColOrder MlColKey, MsKey(), MlKey()
   ColOrder MlColVal, MsVal(), MlVal()

End Sub


Private Sub ColOrder(lCol As Long, sLst() As String, lNdx() As Long)

' Orders column lCol of MsLUT() into sLst() and keeps track of
' the sLUT() indices by the corresponding row indices in lNdx()

   Dim lTmp() As Long
   Dim lEle As Long

   QRS_LibLst.LstAllocL lNdx(), MlLen  ' --- Prepare unordered lists
   QRS_LibA2L.GetColStrStr MsLUT(), sLst(), lCol
   QRS_LibQRS.LstQRNdxS sLst(), lTmp(), vbTextCompare

   For lEle = 1 To MlLen
      lNdx(lEle) = lTmp(lEle)
      sLst(lEle) = MsLUT(lTmp(lEle), lCol)
   Next lEle

End Sub

