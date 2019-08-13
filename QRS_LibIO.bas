Attribute VB_Name = "QRS_LibIO"
Option Explicit

' Module : QRS_LibIO
' Purpose: Data Input/Output to files, format data for output
' By     : QRS, Roger Strebel
' Date   : 25.03.2018                  First ideas
'          19.08.2018                  Text file reader
' --- The public interface
'     PutRefT3
'     ReadFileTxt                      Read text file, return array  19

Public Function ReadFileTxt(sFile As String, vTbl(), _
                            Optional sSep As String = ",", _
                            Optional sTxtBeg As String = "", _
                            Optional sTxtEnd As String = "") As Boolean

' Returns the file content of sFile
' sFile is a delimited text file and
' text is marked by sTxtBeg and sTxtEnd
' text may contain the delimiter

   Const Cl1 As Long = 1

   Dim iFile As Integer
   Dim sRow As String
   Dim lNRow As Long, lNCol As Long
   Dim lP As Long, lQ As Long
   Dim bFail As Boolean

   bFail = Not QRS_LibDOS.FileExists(sFile)
   If bFail Then GoTo ReadFileTxt_Ende

   iFile = FreeFile
   Open sFile For Input As #iFile
   While Not EOF(iFile)
      lNRow = lNRow + Cl1
      Line Input #iFile, sRow
   Wend
   Close #iFile

   If Not lNRow > 0 Then GoTo ReadFileTxt_Ende
   lNCol = QRS_LibStr.StrOcc(sRow, sSep, sTxtBeg, sTxtEnd) + Cl1

   QRS_LibArr.ArrAllocV vTbl(), lNRow, lNCol
   lNRow = 0

   Open sFile For Input As #iFile
   While Not EOF(iFile)
      lNRow = lNRow + Cl1
      Line Input #iFile, sRow
      lNCol = 0
      While Not bFail
         lNCol = lNCol + Cl1
         bFail = QRS_LibStr.StrNexDel(sRow, sSep, sTxtBeg, sTxtEnd, lP, lQ)
         vTbl(lNRow, lNCol) = Mid(sRow, lP, lQ - lP)
         lNCol = lQ
      Wend
      bFail = False
   Wend
   Close #iFile

ReadFileTxt_Ende:

   ReadFileTxt = bFail

End Function

Public Function PutRefT3(sHdr() As String, dDate() As Date, fVal() As Double _
                         ) As Boolean
                         
' Outputs a
                         
End Function

Public Function PutTxtT3(sHdr() As String, dDate() As Date, fVal() As Double, _
                         Optional sDel As String = ";", _
                         Optional sFmtDate As String = "", _
                         Optional sFmtVal As String = "")

' Outputs header, date column and values to a delimited text file

End Function

Public Function PutTxtV(vArr() As Variant, sFile As String, _
                        Optional sDel As String = ";") As Boolean

' Outputs the variant array into a delimited text file
' No formatting

   If QRS_LibDOS.FileExists(sFile) Then QRS_LibDOS.FileDelete sFile

End Function

Public Function PutTxtS(sArr() As String, sFile As String, _
                        Optional sDel As String = ";") As Boolean

End Function

Public Function PutColDblStr(fArr() As Double, sArr() As String, _
                             lColSrc As Long, lColDst As Long, _
                             Optional lRowDst1 As Long = 1, _
                             Optional sFmt As String = "") As Boolean

' Copies a column of double to a string array and formats the values
' Source and destination columns may be specified with negative values
' for right-bound counting
' The first output row is specified in lRowDst1 to handle header rows
' if fArr() as more rows than sArr() output is truncated, returns true
' if lColSrc or lColDst are out of bounds, aborts and returns true

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long

   ArrBoundF fArr(), lSR1, lSRL, lSC1, lSCL
   ArrBoundS sArr(), lDR1, lDRL, lDC1, lDCL
   lICS = QRS_LibArr.NdxGetSX1(lSC1, lSCL, lColSrc)
   lICD = QRS_LibArr.NdxGetSX1(lDC1, lDCL, lColDst)

   If sFmt = "" Then
      For lIRS = 1 To 2
         sArr(lIRD, lICD) = Format(fArr(lIRS, lICS), sFmt)
      Next lIRS
      lIRS = lIRS + Cl01
   Else
   End If
End Function

