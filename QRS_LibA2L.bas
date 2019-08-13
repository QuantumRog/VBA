Attribute VB_Name = "QRS_LibA2L"
Option Explicit

' Module : QRS_LibA2L
' Purpose: Functions for getting lists from arrays
'          and putting lists to arrays
'          Uses auxiliary functions of QRS_LibArr and QRS_LibLst
' By     : QRS, Roger Strebel
' Date   : 26.03.2018                  GetCol/GetRow functions added
'          01.04.2018                  PutCol/PutRow functions added
'          30.07.2018                  GetColVarLon added
'          07.02.2019                  PutColStrVar added
' --- The public interface
'     GetColDatDat                     List from one date column     26.03.2018
'     GetColDblDbl                     List from one real column     26.03.2018
'     GetColLonLon                     List from one long column     26.03.2018
'     GetColStrStr                     List from one string column   26.03.2018
'     GetColVarDat                     Date list from var column     26.03.2018
'     GetColVarLon                     Long list from var column     30.07.2018
'     GetRowDatDat                     List from one date column     26.03.2018
'     GetRowDblDbl                     List from one real column     26.03.2018
'     GetRowLonLon                     List from one real column     26.03.2018
'     GetRowStrStr                     List from one string column   26.03.2018
'     GetRowStrVar                     Var list from one string col  26.03.2018
'     PutColDatVar                     Date List to one var column   01.04.2018
'     PutColDblDbl                     List to one real column       01.04.2018
'     PutColDblVar                     Real list to one var column   01.04.2018
'     PutColLonLon                     List to one long column       01.04.2018
'     PutColStrStr                     List to one string column     01.04.2018
'     PutColStrVar                     String to one var column
'     PutRowDblDbl                     List to one real column       01.04.2018
'     PutRowDblVar                     Real list to one var column   01.04.2018
'     PutRowLonLon                     List to one long column       01.04.2018
'     PutRowStrStr                     List to one string column     01.04.2018
'     PutRowStrVar                     String List to one var column 01.04.2018

Public Sub GetColDatDat(dArrSrc() As Date, dLstDst() As Date, _
                        Optional lCol As Long = 0, _
                        Optional lRow1 As Long = 0, _
                        Optional lRowL As Long = 0)

' Extract one column from a date source array to date list
' When lCol=0, starts at first source column
' When lCol>0, starts at source column number lCol1
' When lCol<0, starts at source column Abs(lCol1) from the right
' When lCol=0, ends at last row

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lUR1 As Long, lURL As Long, lUC1 As Long, lUCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lCol1 As Long, lColL As Long

   lCol1 = lCol
   lColL = lCol1
   QRS_LibArr.ArrBoundD dArrSrc(), lSR1, lSRL, lSC1, lSCL
   QRS_LibArr.NdxGetXSD lSR1, lSRL, lRow1, lRowL, lDR1, lDRL, lUR1, lURL
   QRS_LibArr.NdxGetXSD lSC1, lSCL, lCol1, lColL, lDC1, lDCL, lUC1, lUCL
   lColL = lCol1
   QRS_LibLst.LstAllocD dLstDst(), lURL

   lIRD = lUR1
   For lIRS = lRow1 To lRowL
      dLstDst(lIRD) = dArrSrc(lIRS, lCol1)
      lIRD = lIRD + Cl01
   Next lIRS

End Sub

Public Sub GetColDblDbl(fArrSrc() As Double, fLstDst() As Double, _
                        Optional lCol As Long = 0, _
                        Optional lRow1 As Long = 0, _
                        Optional lRowL As Long = 0)

' Extract one column from a double source array to double list
' When lCol=0, starts at first source column
' When lCol>0, starts at source column number lCol1
' When lCol<0, starts at source column Abs(lCol1) from the right
' When lCol=0, ends at last row

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lUR1 As Long, lURL As Long, lUC1 As Long, lUCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lCol1 As Long, lColL As Long

   lCol1 = lCol
   lColL = lCol1
   QRS_LibArr.ArrBoundF fArrSrc(), lSR1, lSRL, lSC1, lSCL
   QRS_LibArr.NdxGetXSD lSR1, lSRL, lRow1, lRowL, lDR1, lDRL, lUR1, lURL
   QRS_LibArr.NdxGetXSD lSC1, lSCL, lCol1, lColL, lDC1, lDCL, lUC1, lUCL
   lColL = lCol1
   QRS_LibLst.LstAllocF fLstDst(), lURL

   lIRD = lUR1
   For lIRS = lRow1 To lRowL
      fLstDst(lIRD) = fArrSrc(lIRS, lCol1)
      lIRD = lIRD + Cl01
   Next lIRS

End Sub

Public Sub GetColLonLon(lArrSrc() As Long, lLstDst() As Long, _
                        Optional lCol As Long = 0, _
                        Optional lRow1 As Long = 0, _
                        Optional lRowL As Long = 0)

' Extract one column from a long source array to long list
' When lCol=0, starts at first source column
' When lCol>0, starts at source column number lCol1
' When lCol<0, starts at source column Abs(lCol1) from the right
' When lCol=0, ends at last row

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lUR1 As Long, lURL As Long, lUC1 As Long, lUCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lCol1 As Long, lColL As Long

   lCol1 = lCol
   lColL = lCol1
   QRS_LibArr.ArrBoundL lArrSrc(), lSR1, lSRL, lSC1, lSCL
   QRS_LibArr.NdxGetXSD lSR1, lSRL, lRow1, lRowL, lDR1, lDRL, lUR1, lURL
   QRS_LibArr.NdxGetXSD lSC1, lSCL, lCol1, lColL, lDC1, lDCL, lUC1, lUCL
   lColL = lCol1
   QRS_LibLst.LstAllocL lLstDst(), lURL

   lIRD = lUR1
   For lIRS = lRow1 To lRowL
      lLstDst(lIRD) = lArrSrc(lIRS, lCol1)
      lIRD = lIRD + Cl01
   Next lIRS

End Sub

Public Sub GetColStrStr(sArrSrc() As String, sLstDst() As String, _
                        Optional lCol As Long = 0, _
                        Optional lRow1 As Long = 0, _
                        Optional lRowL As Long = 0)

' Extract one column from a string source array to string list
' When lCol=0, starts at first source column
' When lCol>0, starts at source column number lCol1
' When lCol<0, starts at source column Abs(lCol1) from the right
' When lCol=0, ends at last row

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lUR1 As Long, lURL As Long, lUC1 As Long, lUCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lCol1 As Long, lColL As Long

   lCol1 = lCol
   lColL = lCol1
   QRS_LibArr.ArrBoundS sArrSrc(), lSR1, lSRL, lSC1, lSCL
   QRS_LibArr.NdxGetXSD lSR1, lSRL, lRow1, lRowL, lDR1, lDRL, lUR1, lURL
   QRS_LibArr.NdxGetXSD lSC1, lSCL, lCol1, lColL, lDC1, lDCL, lUC1, lUCL
   lColL = lCol1
   QRS_LibLst.LstAllocS sLstDst(), lURL

   lIRD = lUR1
   For lIRS = lRow1 To lRowL
      sLstDst(lIRD) = sArrSrc(lIRS, lCol1)
      lIRD = lIRD + Cl01
   Next lIRS

End Sub

Public Sub GetColVarDat(vArrSrc() As Variant, dLstDst() As Date, _
                        Optional lCol As Long = 0, _
                        Optional lRow1 As Long = 0, _
                        Optional lRowL As Long = 0)

' Extract one column from a variant source array to date list
' When lCol=0, starts at first source column
' When lCol>0, starts at source column number lCol1
' When lCol<0, starts at source column Abs(lCol1) from the right
' When lCol=0, ends at last row

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lUR1 As Long, lURL As Long, lUC1 As Long, lUCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lCol1 As Long, lColL As Long

   lCol1 = lCol
   lColL = lCol1
   QRS_LibArr.ArrBoundV vArrSrc(), lSR1, lSRL, lSC1, lSCL
   QRS_LibArr.NdxGetXSD lSR1, lSRL, lRow1, lRowL, lDR1, lDRL, lUR1, lURL
   QRS_LibArr.NdxGetXSD lSC1, lSCL, lCol1, lColL, lDC1, lDCL, lUC1, lUCL
   lColL = lCol1
   QRS_LibLst.LstAllocD dLstDst(), lURL

   lIRD = lUR1
   For lIRS = lRow1 To lRowL
      dLstDst(lIRD) = vArrSrc(lIRS, lCol1)
      lIRD = lIRD + Cl01
   Next lIRS

End Sub

Public Sub GetColVarLon(vArrSrc() As Variant, lLstDst() As Long, _
                        Optional lCol As Long = 0, _
                        Optional lRow1 As Long = 0, _
                        Optional lRowL As Long = 0)

' Extract one column from a variant source array to date list
' When lCol=0, starts at first source column
' When lCol>0, starts at source column number lCol1
' When lCol<0, starts at source column Abs(lCol1) from the right
' When lCol=0, ends at last row

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lUR1 As Long, lURL As Long, lUC1 As Long, lUCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lCol1 As Long, lColL As Long

   lCol1 = lCol
   lColL = lCol1
   QRS_LibArr.ArrBoundV vArrSrc(), lSR1, lSRL, lSC1, lSCL
   QRS_LibArr.NdxGetXSD lSR1, lSRL, lRow1, lRowL, lDR1, lDRL, lUR1, lURL
   QRS_LibArr.NdxGetXSD lSC1, lSCL, lCol1, lColL, lDC1, lDCL, lUC1, lUCL
   lColL = lCol1
   QRS_LibLst.LstAllocL lLstDst(), lURL

   lIRD = lUR1
   For lIRS = lRow1 To lRowL
      lLstDst(lIRD) = vArrSrc(lIRS, lCol1)
      lIRD = lIRD + Cl01
   Next lIRS

End Sub

Public Sub GetColVarStr(vArrSrc() As Variant, sLstDst() As String, _
                        Optional lCol As Long = 0, _
                        Optional lRow1 As Long = 0, _
                        Optional lRowL As Long = 0)

' Extract one column from a variant source array to date list
' When lCol=0, starts at first source column
' When lCol>0, starts at source column number lCol1
' When lCol<0, starts at source column Abs(lCol1) from the right
' When lCol=0, ends at last row

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lUR1 As Long, lURL As Long, lUC1 As Long, lUCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lCol1 As Long, lColL As Long

   lCol1 = lCol
   lColL = lCol1
   QRS_LibArr.ArrBoundV vArrSrc(), lSR1, lSRL, lSC1, lSCL
   QRS_LibArr.NdxGetXSD lSR1, lSRL, lRow1, lRowL, lDR1, lDRL, lUR1, lURL
   QRS_LibArr.NdxGetXSD lSC1, lSCL, lCol1, lColL, lDC1, lDCL, lUC1, lUCL
   lColL = lCol1
   QRS_LibLst.LstAllocS sLstDst(), lURL

   lIRD = lUR1
   For lIRS = lRow1 To lRowL
      sLstDst(lIRD) = vArrSrc(lIRS, lCol1)
      lIRD = lIRD + Cl01
   Next lIRS

End Sub

Public Sub GetRowDatDat(dArrSrc() As Date, dLstDst() As Date, _
                        Optional lRow As Long = 0, _
                        Optional lCol1 As Long = 0, _
                        Optional lColL As Long = 0)

' Extract one row from a date source array to date list
' When lRow1=0, starts at first source row
' When lRow1>0, starts at source row number lRow1
' When lRow1<0, starts at source row Abs(lRow1) from the bottom
' When lRowL=0, ends at last row

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lUR1 As Long, lURL As Long, lUC1 As Long, lUCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lRow1 As Long, lRowL As Long

   lRow1 = lRow
   lRowL = lRow1
   QRS_LibArr.ArrBoundD dArrSrc(), lSR1, lSRL, lSC1, lSCL
   QRS_LibArr.NdxGetXSD lSR1, lSRL, lRow1, lRowL, lDR1, lDRL, lUR1, lURL
   QRS_LibArr.NdxGetXSD lSC1, lSCL, lCol1, lColL, lDC1, lDCL, lUC1, lUCL
   QRS_LibLst.LstAllocD dLstDst(), lUCL

   lICD = lUC1
   For lICS = lCol1 To lColL
      dLstDst(lICD) = dArrSrc(lRow1, lICS)
      lICD = lICD + Cl01
   Next lICS

End Sub

Public Sub GetRowDblDbl(fArrSrc() As Double, fLstDst() As Double, _
                        Optional lRow As Long = 0, _
                        Optional lCol1 As Long = 0, _
                        Optional lColL As Long = 0)

' Extract one row from a double source array to double list
' When lRow1=0, starts at first source row
' When lRow1>0, starts at source row number lRow1
' When lRow1<0, starts at source row Abs(lRow1) from the bottom
' When lRowL=0, ends at last row

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lUR1 As Long, lURL As Long, lUC1 As Long, lUCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lRow1 As Long, lRowL As Long

   lRow1 = lRow
   lRowL = lRow1
   QRS_LibArr.ArrBoundF fArrSrc(), lSR1, lSRL, lSC1, lSCL
   QRS_LibArr.NdxGetXSD lSR1, lSRL, lRow1, lRowL, lDR1, lDRL, lUR1, lURL
   QRS_LibArr.NdxGetXSD lSC1, lSCL, lCol1, lColL, lDC1, lDCL, lUC1, lUCL
   QRS_LibLst.LstAllocF fLstDst(), lUCL

   lICD = lUC1
   For lICS = lCol1 To lColL
      fLstDst(lICD) = fArrSrc(lRow1, lICS)
      lICD = lICD + Cl01
   Next lICS

End Sub

Public Sub GetRowLonLon(lArrSrc() As Long, lLstDst() As Long, _
                        Optional lRow As Long = 0, _
                        Optional lCol1 As Long = 0, _
                        Optional lColL As Long = 0)

' Extract one row from a long source array to long list
' When lRow1=0, starts at first source row
' When lRow1>0, starts at source row number lRow1
' When lRow1<0, starts at source row Abs(lRow1) from the bottom
' When lRowL=0, ends at last row

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lUR1 As Long, lURL As Long, lUC1 As Long, lUCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lRow1 As Long, lRowL As Long

   lRow1 = lRow
   lRowL = lRow1
   QRS_LibArr.ArrBoundL lArrSrc(), lSR1, lSRL, lSC1, lSCL
   QRS_LibArr.NdxGetXSD lSR1, lSRL, lRow1, lRowL, lDR1, lDRL, lUR1, lURL
   QRS_LibArr.NdxGetXSD lSC1, lSCL, lCol1, lColL, lDC1, lDCL, lUC1, lUCL
   QRS_LibLst.LstAllocL lLstDst(), lUCL

   lICD = lUC1
   For lICS = lCol1 To lColL
      lLstDst(lICD) = lArrSrc(lRow1, lICS)
      lICD = lICD + Cl01
   Next lICS

End Sub

Public Sub GetRowStrStr(sArrSrc() As String, sLstDst() As String, _
                        Optional lRow As Long = 0, _
                        Optional lCol1 As Long = 0, _
                        Optional lColL As Long = 0)

' Extract one row from a string source array to string list
' When lRow1=0, starts at first source row
' When lRow1>0, starts at source row number lRow1
' When lRow1<0, starts at source row Abs(lRow1) from the bottom
' When lRowL=0, ends at last row

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lUR1 As Long, lURL As Long, lUC1 As Long, lUCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lRow1 As Long, lRowL As Long

   lRow1 = lRow
   lRowL = lRow1
   QRS_LibArr.ArrBoundS sArrSrc(), lSR1, lSRL, lSC1, lSCL
   QRS_LibArr.NdxGetXSD lSR1, lSRL, lRow1, lRowL, lDR1, lDRL, lUR1, lURL
   QRS_LibArr.NdxGetXSD lSC1, lSCL, lCol1, lColL, lDC1, lDCL, lUC1, lUCL
   QRS_LibLst.LstAllocS sLstDst(), lUCL

   lICD = lUC1
   For lICS = lCol1 To lColL
      sLstDst(lICD) = sArrSrc(lRow1, lICS)
      lICD = lICD + Cl01
   Next lICS

End Sub

Public Sub GetRowStrVar(vArrSrc() As Variant, sLstDst() As String, _
                        Optional lRow As Long = 0, _
                        Optional lCol1 As Long = 0, _
                        Optional lColL As Long = 0)

' Extract one row from a variant source array to string list
' When lRow1=0, starts at first source row
' When lRow1>0, starts at source row number lRow1
' When lRow1<0, starts at source row Abs(lRow1) from the bottom
' When lRowL=0, ends at last row

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lUR1 As Long, lURL As Long, lUC1 As Long, lUCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lRow1 As Long, lRowL As Long

   lRow1 = lRow
   lRowL = lRow1
   QRS_LibArr.ArrBoundV vArrSrc(), lSR1, lSRL, lSC1, lSCL
   QRS_LibArr.NdxGetXSD lSR1, lSRL, lRow1, lRowL, lDR1, lDRL, lUR1, lURL
   QRS_LibArr.NdxGetXSD lSC1, lSCL, lCol1, lColL, lDC1, lDCL, lUC1, lUCL
   QRS_LibLst.LstAllocS sLstDst(), lUCL

   lICD = lUC1
   For lICS = lCol1 To lColL
      sLstDst(lICD) = vArrSrc(lRow1, lICS)
      lICD = lICD + Cl01
   Next lICS

End Sub

Public Sub PutColDatVar(dLstSrc() As Date, vArrDst() As Variant, _
                        Optional lCol As Long = 0, _
                        Optional lRow1 As Long = 0, _
                        Optional lRowL As Long = 0)

' Outputs a date list to one column of a variant destination array
' Position indication arguments all refer to destination array
' When lCol=0, starts at first source column
' When lCol>0, starts at source column number lCol1
' When lCol<0, starts at source column Abs(lCol1) from the right
' When lCol=0, ends at last row

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lVR1 As Long, lVRL As Long, lVC1 As Long, lVCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lCol1 As Long, lColL As Long    ' --- Partial row

   lCol1 = lCol
   QRS_LibLst.LstBoundD dLstSrc(), lSR1, lSRL
   QRS_LibArr.ArrBoundV vArrDst(), lDR1, lDRL, lDC1, lDCL
   QRS_LibArr.NdxGetXDS lDR1, lDRL, lSR1, lSRL, lRow1, lRowL, lVR1, lVRL
   QRS_LibArr.NdxGetXDS lDC1, lDCL, Cl01, Cl01, lCol1, lColL, lVC1, lVCL

   lIRS = lVR1
   For lIRD = lRow1 To lRowL           ' --- All elements of column
      vArrDst(lIRD, lCol1) = dLstSrc(lIRS)
      lIRS = lIRS + Cl01
   Next lIRD

End Sub

Public Sub PutColDblDbl(fLstSrc() As Double, fArrDst() As Double, _
                        Optional lCol As Long = 0, _
                        Optional lRow1 As Long = 0, _
                        Optional lRowL As Long = 0)

' Outputs a double list to one column of a double destination array
' Position indication arguments all refer to destination array
' When lCol=0, starts at first source column
' When lCol>0, starts at source column number lCol1
' When lCol<0, starts at source column Abs(lCol1) from the right
' When lCol=0, ends at last row

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lVR1 As Long, lVRL As Long, lVC1 As Long, lVCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lCol1 As Long, lColL As Long    ' --- Partial row

   lCol1 = lCol
   QRS_LibLst.LstBoundF fLstSrc(), lSR1, lSRL
   QRS_LibArr.ArrBoundF fArrDst(), lDR1, lDRL, lDC1, lDCL
   QRS_LibArr.NdxGetXDS lDR1, lDRL, lSR1, lSRL, lRow1, lRowL, lVR1, lVRL
   QRS_LibArr.NdxGetXDS lDC1, lDCL, Cl01, Cl01, lCol1, lColL, lVC1, lVCL

   lIRS = lVR1
   For lIRD = lRow1 To lRowL           ' --- All elements of column
      fArrDst(lIRD, lCol1) = fLstSrc(lIRS)
      lIRS = lIRS + Cl01
   Next lIRD

End Sub

Public Sub PutColLonLon(lLstSrc() As Long, lArrDst() As Long, _
                        Optional lCol As Long = 0, _
                        Optional lRow1 As Long = 0, _
                        Optional lRowL As Long = 0)

' Outputs a long list to one column of a long destination array
' Position indication arguments all refer to destination array
' When lCol=0, starts at first source column
' When lCol>0, starts at source column number lCol1
' When lCol<0, starts at source column Abs(lCol1) from the right
' When lCol=0, ends at last row

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lVR1 As Long, lVRL As Long, lVC1 As Long, lVCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lCol1 As Long, lColL As Long    ' --- Partial row

   lCol1 = lCol
   QRS_LibLst.LstBoundL lLstSrc(), lSR1, lSRL
   QRS_LibArr.ArrBoundL lArrDst(), lDR1, lDRL, lDC1, lDCL
   QRS_LibArr.NdxGetXDS lDR1, lDRL, lSR1, lSRL, lRow1, lRowL, lVR1, lVRL
   QRS_LibArr.NdxGetXDS lDC1, lDCL, Cl01, Cl01, lCol1, lColL, lVC1, lVCL

   lIRS = lVR1
   For lIRD = lRow1 To lRowL           ' --- All elements of column
      lArrDst(lIRD, lCol1) = lLstSrc(lIRS)
      lIRS = lIRS + Cl01
   Next lIRD

End Sub

Public Sub PutColStrStr(sLstSrc() As String, sArrDst() As String, _
                        Optional lCol As Long = 0, _
                        Optional lRow1 As Long = 0, _
                        Optional lRowL As Long = 0)

' Outputs a string list to one column of a string destination array
' Position indication arguments all refer to destination array
' When lCol=0, starts at first source column
' When lCol>0, starts at source column number lCol1
' When lCol<0, starts at source column Abs(lCol1) from the right
' When lCol=0, ends at last row

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lVR1 As Long, lVRL As Long, lVC1 As Long, lVCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lCol1 As Long, lColL As Long    ' --- Partial row

   lCol1 = lCol
   QRS_LibLst.LstBoundS sLstSrc(), lSR1, lSRL
   QRS_LibArr.ArrBoundS sArrDst(), lDR1, lDRL, lDC1, lDCL
   QRS_LibArr.NdxGetXDS lDR1, lDRL, lSR1, lSRL, lRow1, lRowL, lVR1, lVRL
   QRS_LibArr.NdxGetXDS lDC1, lDCL, Cl01, Cl01, lCol1, lColL, lVC1, lVCL

   lIRS = lVR1
   For lIRD = lRow1 To lRowL           ' --- All elements of column
      sArrDst(lIRD, lCol1) = sLstSrc(lIRS)
      lIRS = lIRS + Cl01
   Next lIRD

End Sub

Public Sub PutColStrVar(sLstSrc() As String, vArrDst() As Variant, _
                        Optional lCol As Long = 0, _
                        Optional lRow1 As Long = 0, _
                        Optional lRowL As Long = 0)

' Outputs a string list to one column of a string destination array
' Position indication arguments all refer to destination array
' When lCol=0, starts at first source column
' When lCol>0, starts at source column number lCol1
' When lCol<0, starts at source column Abs(lCol1) from the right
' When lCol=0, ends at last row

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lVR1 As Long, lVRL As Long, lVC1 As Long, lVCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lCol1 As Long, lColL As Long    ' --- Partial row

   lCol1 = lCol
   QRS_LibLst.LstBoundS sLstSrc(), lSR1, lSRL
   QRS_LibArr.ArrBoundV vArrDst(), lDR1, lDRL, lDC1, lDCL
   QRS_LibArr.NdxGetXDS lDR1, lDRL, lSR1, lSRL, lRow1, lRowL, lVR1, lVRL
   QRS_LibArr.NdxGetXDS lDC1, lDCL, Cl01, Cl01, lCol1, lColL, lVC1, lVCL

   lIRS = lVR1
   For lIRD = lRow1 To lRowL           ' --- All elements of column
      vArrDst(lIRD, lCol1) = sLstSrc(lIRS)
      lIRS = lIRS + Cl01
   Next lIRD

End Sub

Public Sub PutColDblVar(fLstSrc() As Double, vArrDst() As Variant, _
                        Optional lCol As Long = 0, _
                        Optional lRow1 As Long = 0, _
                        Optional lRowL As Long = 0)

' Outputs a double list to one column of a variant destination array
' Position indication arguments all refer to destination array
' When lCol=0, starts at first source column
' When lCol>0, starts at source column number lCol1
' When lCol<0, starts at source column Abs(lCol1) from the right
' When lCol=0, ends at last row

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lVR1 As Long, lVRL As Long, lVC1 As Long, lVCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lCol1 As Long, lColL As Long    ' --- Partial row

   lCol1 = lCol
   QRS_LibLst.LstBoundF fLstSrc(), lSR1, lSRL
   QRS_LibArr.ArrBoundV vArrDst(), lDR1, lDRL, lDC1, lDCL
   QRS_LibArr.NdxGetXDS lDR1, lDRL, lSR1, lSRL, lRow1, lRowL, lVR1, lVRL
   QRS_LibArr.NdxGetXDS lDC1, lDCL, Cl01, Cl01, lCol1, lColL, lVC1, lVCL

   lIRS = lVR1
   For lIRD = lRow1 To lRowL           ' --- All elements of column
      vArrDst(lIRD, lCol1) = fLstSrc(lIRS)
      lIRS = lIRS + Cl01
   Next lIRD

End Sub

Public Sub PutRowDblDbl(fLstSrc() As Double, fArrDst() As Double, _
                        Optional lRow As Long = 0, _
                        Optional lCol1 As Long = 0, _
                        Optional lColL As Long = 0)

' Outputs a double list to one row of a double destination array
' Position indication arguments all refer to destination array
' When lRow=0, starts at first source column
' When lRow>0, starts at source column number lCol1
' When lRow<0, starts at source column Abs(lCol1) from the right

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lVR1 As Long, lVRL As Long, lVC1 As Long, lVCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lRow1 As Long, lRowL As Long    ' --- Partial col

   lRow1 = lRow
   QRS_LibLst.LstBoundF fLstSrc(), lSC1, lSCL
   QRS_LibArr.ArrBoundF fArrDst(), lDR1, lDRL, lDC1, lDCL
   QRS_LibArr.NdxGetXDS lDC1, lDCL, lSC1, lSCL, lCol1, lColL, lVC1, lVCL
   QRS_LibArr.NdxGetXDS lDR1, lDRL, Cl01, Cl01, lRow1, lRowL, lVR1, lVRL

   lICS = lVC1
   For lICD = lCol1 To lColL           ' --- All elements of row
      fArrDst(lRow1, lICD) = fLstSrc(lICS)
      lICS = lICS + Cl01
   Next lICD

End Sub

Public Sub PutRowLonLon(lLstSrc() As Long, lArrDst() As Long, _
                        Optional lRow As Long = 0, _
                        Optional lCol1 As Long = 0, _
                        Optional lColL As Long = 0)

' Outputs a long list to one row of a long destination array
' Position indication arguments all refer to destination array
' When lRow=0, starts at first source column
' When lRow>0, starts at source column number lCol1
' When lRow<0, starts at source column Abs(lCol1) from the right

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lVR1 As Long, lVRL As Long, lVC1 As Long, lVCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lRow1 As Long, lRowL As Long    ' --- Partial col

   lRow1 = lRow
   QRS_LibLst.LstBoundL lLstSrc(), lSC1, lSCL
   QRS_LibArr.ArrBoundL lArrDst(), lDR1, lDRL, lDC1, lDCL
   QRS_LibArr.NdxGetXDS lDC1, lDCL, lSC1, lSCL, lCol1, lColL, lVC1, lVCL
   QRS_LibArr.NdxGetXDS lDR1, lDRL, Cl01, Cl01, lRow1, lRowL, lVR1, lVRL

   lICS = lVC1
   For lICD = lCol1 To lColL           ' --- All elements of row
      lArrDst(lRow1, lICD) = lLstSrc(lICS)
      lICS = lICS + Cl01
   Next lICD

End Sub

Public Sub PutRowDblVar(fLstSrc() As Double, vArrDst() As Variant, _
                        Optional lRow As Long = 0, _
                        Optional lCol1 As Long = 0, _
                        Optional lColL As Long = 0)

' Outputs a double list to one row of a variant destination array
' Position indication arguments all refer to destination array
' When lRow=0, starts at first source column
' When lRow>0, starts at source column number lCol1
' When lRow<0, starts at source column Abs(lCol1) from the right

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lVR1 As Long, lVRL As Long, lVC1 As Long, lVCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lRow1 As Long, lRowL As Long    ' --- Partial col

   lRow1 = lRow
   QRS_LibLst.LstBoundF fLstSrc(), lSC1, lSCL
   QRS_LibArr.ArrBoundV vArrDst(), lDR1, lDRL, lDC1, lDCL
   QRS_LibArr.NdxGetXDS lDC1, lDCL, lSC1, lSCL, lCol1, lColL, lVC1, lVCL
   QRS_LibArr.NdxGetXDS lDR1, lDRL, Cl01, Cl01, lRow1, lRowL, lVR1, lVRL

   lICS = lVC1
   For lICD = lCol1 To lColL           ' --- All elements of row
      vArrDst(lRow1, lICD) = fLstSrc(lICS)
      lICS = lICS + Cl01
   Next lICD

End Sub

Public Sub PutRowStrStr(sLstSrc() As String, sArrDst() As String, _
                        Optional lRow As Long = 0, _
                        Optional lCol1 As Long = 0, _
                        Optional lColL As Long = 0)

' Outputs a string list to one row of a string destination array
' Position indication arguments all refer to destination array
' When lRow=0, starts at first source column
' When lRow>0, starts at source column number lCol1
' When lRow<0, starts at source column Abs(lCol1) from the right

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lVR1 As Long, lVRL As Long, lVC1 As Long, lVCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lRow1 As Long, lRowL As Long    ' --- Partial col

   lRow1 = lRow
   QRS_LibLst.LstBoundS sLstSrc(), lSC1, lSCL
   QRS_LibArr.ArrBoundS sArrDst(), lDR1, lDRL, lDC1, lDCL
   QRS_LibArr.NdxGetXDS lDC1, lDCL, lSC1, lSCL, lCol1, lColL, lVC1, lVCL
   QRS_LibArr.NdxGetXDS lDR1, lDRL, Cl01, Cl01, lRow1, lRowL, lVR1, lVRL

   lICS = lVC1
   For lICD = lCol1 To lColL           ' --- All elements of row
      sArrDst(lRow1, lICD) = sLstSrc(lICS)
      lICS = lICS + Cl01
   Next lICD

End Sub

Public Sub PutRowStrVar(sLstSrc() As String, vArrDst() As Variant, _
                        Optional lRow As Long = 0, _
                        Optional lCol1 As Long = 0, _
                        Optional lColL As Long = 0)

' Outputs a string list to one row of a string destination array
' Position indication arguments all refer to destination array
' When lRow=0, starts at first source column
' When lRow>0, starts at source column number lCol1
' When lRow<0, starts at source column Abs(lCol1) from the right

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lVR1 As Long, lVRL As Long, lVC1 As Long, lVCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lRow1 As Long, lRowL As Long    ' --- Partial col

   lRow1 = lRow
   QRS_LibLst.LstBoundS sLstSrc(), lSC1, lSCL
   QRS_LibArr.ArrBoundV vArrDst(), lDR1, lDRL, lDC1, lDCL
   QRS_LibArr.NdxGetXDS lDC1, lDCL, lSC1, lSCL, lCol1, lColL, lVC1, lVCL
   QRS_LibArr.NdxGetXDS lDR1, lDRL, Cl01, Cl01, lRow1, lRowL, lVR1, lVRL

   lICS = lVC1
   For lICD = lCol1 To lColL           ' --- All elements of row
      vArrDst(lRow1, lICD) = sLstSrc(lICS)
      lICS = lICS + Cl01
   Next lICD

End Sub

