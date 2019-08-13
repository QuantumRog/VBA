Attribute VB_Name = "QRS_LibArr"
Option Explicit

' Module :  QRS_LibArr
' Project:  any
' Purpose:  Some very basic typed array utility VBA routines
'           This library supports five array data types "X"
'              Dates                                   ("D")
'              Double precision floating point values  ("F")
'              VBA Long (32 bit) integers              ("L")
'              Strings                                 ("S")
'              Variant                                 ("V")
'           Array references of unspecified type cannot be passed by reference
'           Array allocation and bound checking
'              ArrIsAllX               Returns true if array is allocated
'              ArrAllocX               Alloc if new, re-alloc if new size
'              ArrBoundX               Array bounds
'           Get operations extract data from a source array to a destination
'           The destination dimensions are determined by the source size and
'           extraction parameters and also work for data of size 1x1
'           Put operations output data from a source array to a destination
'           The destination start and end positions are specified and clipped
'           to avoid out of bound errors at either source or destination
' By     : QRS, Roger Strebel
' Date   : 18.02.2018
'          28.02.2018                  Versatile bound handling
'          04.03.2018                  GetRow and GetSub functions work
'          14.03.2018                  Serious improvements and 1D functions
'          15.03.2018                  Useful set of Get operations
'          17.03.2018                  Xtr functions added
'          18.03.2018                  GetSubDblVar added
'          19.03.2018                  Merge functions added
'          26.03.2018                  NdxGetXDS added for Put preparations
'          01.04.2018                  NdxGetXDS improved and thoroughly tested
'          02.04.2018                  PutCol operations added
'          04.07.2018                  Zip function added
'          15.08.2018                  Interface documentation updated
'          28.01.2019                  XtrEleVarCol bug fixed
'          13.02.2019                  ZipCol function added
' --- The public interface
'     ArrAllocD                        Allocate/sizeof dates Array   18.02.2018
'     ArrAllocF                        Allocate/sizeof double Array  18.02.2018
'     ArrAllocL                        Allocate/sizeof long Array    18.02.2018
'     ArrAllocS                        Allocate/sizeof string Array  14.03.2018
'     ArrAllocV                        Allocate/sizeof variant Array 18.02.2018
'     ArrBoundD                        Bounds of dates array         18.02.2018
'     ArrBoundF                        Bounds of double array        18.02.2018
'     ArrBoundL                        Bounds of long array          18.02.2018
'     ArrBoundS                        Bounds of string array        18.02.2018
'     ArrBoundV                        Bounds of variant array       18.02.2018
'     ArrIsAllD                        Is dates array allocated?     18.02.2018
'     ArrIsAllF                        Is double array allocated?    18.02.2018
'     ArrIsAllL                        Is long array allocated?      18.02.2018
'     ArrIsAllS                        Is string array allocated?    18.02.2018
'     ArrIsAllV                        Is variant array allocated?   18.02.2018
'     ArrMergeF                        Merge 2 real arrays into 1    22.03.2018
'     ArrMergeL                        Merge 2 long arrays into 1    23.03.2018
'     ArrMergeS                        Merge 2 string arrays to 1    23.03.2018
'     ArrMergeV                        Merge 2 variant arrays into 1 23.03.2018
'     GetColDblDbl                     Extract source columns        14.03.2018
'     GetColLonLon                     Extract source columns        14.03.2018
'     GetColVarDat                     Extract source columns (1D)   14.03.2018
'     GetColVarDbl                     Extract source columns        14.03.2018
'     GetColVarLon                     Extract source columns        14.03.2018
'     GetRowDblDbl                     Extract source rows           04.03.2018
'     GetRowLonLon                     Extract source rows           04.03.2018
'     GetRowStrStr                     Extract source rows           04.03.2018
'     GetRowVarDat                     Extract source rows (1D)      14.03.2018
'     GetRowVarDbl                     Extract source rows           04.03.2018
'     GetRowVarLon                     Extract source rows           04.03.2018
'     GetSubDblDbl                     Extract real sub from real    14.03.2018
'     GetSubDblVar                     Extract variant sub from real 18.03.2018
'     GetSubLonLon                     Extract long sub from long    14.03.2018
'     GetSubStrStr                     Extract long sub from string  14.03.2018
'     GetSubVarDbl                     Extract real sub-array        04.03.2018
'     GetSubVarLon                     Extract long sub-array        04.03.2018
'     GetSubVarStr                     Extract string sub-array      04.03.2018
'     GetSubVarVar                     Extract variant sub-array     04.03.2018
'     NdxClpSX1                        Clip destination start        28.02.2018
'     NdxClpSXL                        Clip destination end          28.02.2018
'     NdxGetSX1                        Get source start index        23.02.2018
'     NdxGetSXL                        Get source end   index        23.02.2018
'     NdxGetXDS                        Get dest and source indices   01.04.2018
'     NdxGetXSD                        Get source and dest indices   28.02.2018
'     NdxLenSXL                        Get source end from length    23.02.2018
'     PutColDatVar                     Output source columns         02.04.2018
'     PutColDblDbl                     Output source columns         02.04.2018
'     PutColLonLon                     Output source columns         02.04.2018
'     PutColStrStr                     Output source columns         02.04.2018
'     PutRowDblDbl                     Output source rows            02.04.2018
'     PutRowDblVar                     Output source rows            02.04.2018
'     PutRowLonLon                     Output source rows            02.04.2018
'     PutRowLonVar                     Output source rows            02.04.2018
'     PutRowStrVar                     Output source rows            02.04.2018
'     SetEleVarAbs                     Assign elements by abs index
'     XtrEleVarAbs                     Extract elements by abs index 17.03.2018
'     XtrEleVarCol                     Extract elements in columns   28.01.2019
'     XtrEleVarRel                     Extract elements by rel index
'     XtrEleVarRow                     Extract elements in rows      17.03.2018
'     ZipEleVarCol                     Insert elements to columns    13.02.2019
'     ZipEleVarRow                     Insert elements in rows
   
Public Sub GetColDblDbl(fSrc() As Double, fDst() As Double, _
                        Optional lCol1 As Long = 0, _
                        Optional lColL As Long = 0)

' Extract specific or all columns from a double source array to double array
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

   ArrBoundF fSrc(), lSR1, lSRL, lSC1, lSCL
   NdxGetXSD lSR1, lSRL, lRow1, lRowL, lDR1, lDRL, lUR1, lURL
   NdxGetXSD lSC1, lSCL, lCol1, lColL, lDC1, lDCL, lUC1, lUCL
   ArrAllocF fDst(), lURL, lUCL

   lICD = lUC1
   For lICS = lCol1 To lColL
      lIRD = lUR1
      For lIRS = lRow1 To lRowL
         fDst(lIRD, lICD) = fSrc(lIRS, lICS)
         lIRD = lIRD + Cl01
      Next lIRS
      lICD = lICD + Cl01
   Next lICS

End Sub

Public Sub GetColLonLon(lSrc() As Long, lDst() As Long, _
                        Optional lCol1 As Long = 0, _
                        Optional lColL As Long = 0)

' Extract specific or all columns from a long source array to long array
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

   ArrBoundL lSrc(), lSR1, lSRL, lSC1, lSCL
   NdxGetXSD lSR1, lSRL, lRow1, lRowL, lDR1, lDRL, lUR1, lURL
   NdxGetXSD lSC1, lSCL, lCol1, lColL, lDC1, lDCL, lUC1, lUCL
   ArrAllocL lDst(), lURL, lUCL

   lICD = lUC1
   For lICS = lCol1 To lColL
      lIRD = lUR1
      For lIRS = lRow1 To lRowL
         lDst(lIRD, lICD) = lSrc(lIRS, lICS)
         lIRD = lIRD + Cl01
      Next lIRS
      lICD = lICD + Cl01
   Next lICS

End Sub

Public Sub GetColVarDat(vSrc() As Variant, dDst() As Date, _
                        Optional lCol1 As Long = 0, _
                        Optional lColL As Long = 0, _
                        Optional bList As Boolean = True)

' Extract specific or all columns from a variant source array to dates array
' if bList is set, returns row lCol1 in a 1D list
' When lCol1=0, starts at first source column
' When lCol1>0, starts at source column number lCol1
' When lCol1<0, starts at source column Abs(lCol1) from the right

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lUR1 As Long, lURL As Long, lUC1 As Long, lUCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lRow1 As Long, lRowL As Long
   Dim bDoAll As Boolean

   ArrBoundV vSrc(), lSR1, lSRL, lSC1, lSCL
   NdxGetXSD lSR1, lSRL, lRow1, lRowL, lDR1, lDRL, lUR1, lURL
   NdxGetXSD lSC1, lSCL, lCol1, lColL, lDC1, lDCL, lUC1, lUCL

   If bList Then                       ' --- return 1D list
      bDoAll = Not ArrIsAllD(dDst())
      If Not bDoAll Then               ' --- Subset of QRS_LibLst.LstAllD
         bDoAll = Not LBound(dDst()) = lUR1 And UBound(dDst()) = lURL
      End If
      If bDoAll Then ReDim dDst(1 To lURL + Cl01 - lUR1)
      lIRD = lUR1
      For lIRS = lRow1 To lRowL
         dDst(lIRD) = vSrc(lIRS, lCol1)
         lIRD = lIRD + Cl01
      Next lIRS
   Else
      ArrAllocD dDst(), lURL, lUCL
      lICD = lUC1
      For lICS = lCol1 To lColL
         lIRD = lUR1
         For lIRS = lRow1 To lRowL
            dDst(lIRD, lICD) = vSrc(lIRS, lICS)
            lIRD = lIRD + Cl01
         Next lIRS
         lICD = lICD + Cl01
      Next lICS
   End If

End Sub

Public Sub GetColVarDbl(vSrc() As Variant, fDst() As Double, _
                        Optional lCol1 As Long = 0, _
                        Optional lColL As Long = 0)

' Extract specific or all columns from a variant source array to double array
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

   ArrBoundV vSrc(), lSR1, lSRL, lSC1, lSCL
   NdxGetXSD lSR1, lSRL, lRow1, lRowL, lDR1, lDRL, lUR1, lURL
   NdxGetXSD lSC1, lSCL, lCol1, lColL, lDC1, lDCL, lUC1, lUCL
   ArrAllocF fDst(), lURL, lUCL

   lICD = lUC1
   For lICS = lCol1 To lColL
      lIRD = lUR1
      For lIRS = lRow1 To lRowL
         fDst(lIRD, lICD) = vSrc(lIRS, lICS)
         lIRD = lIRD + Cl01
      Next lIRS
      lICD = lICD + Cl01
   Next lICS

End Sub

Public Sub GetColVarLon(vSrc() As Variant, lDst() As Long, _
                        Optional lCol1 As Long = 0, _
                        Optional lColL As Long = 0)

' Extract specific or all columns from a variant source array to long array
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

   ArrBoundV vSrc(), lSR1, lSRL, lSC1, lSCL
   NdxGetXSD lSR1, lSRL, lRow1, lRowL, lDR1, lDRL, lUR1, lURL
   NdxGetXSD lSC1, lSCL, lCol1, lColL, lDC1, lDCL, lUC1, lUCL
   ArrAllocL lDst(), lURL, lUCL

   lICD = lUC1
   For lICS = lCol1 To lColL
      lIRD = lUR1
      For lIRS = lRow1 To lRowL
         lDst(lIRD, lICD) = vSrc(lIRS, lICS)
         lIRD = lIRD + Cl01
      Next lIRS
      lICD = lICD + Cl01
   Next lICS

End Sub

Public Sub GetRowDblDbl(fSrc() As Double, fDst() As Double, _
                        Optional lRow1 As Long = 0, _
                        Optional lRowL As Long = 0)

' Extract specific or all rows from a double source array to double array
' When lRow1=0, starts at first source row
' When lRow1>0, starts at source row number lRow1
' When lRow1<0, starts at source row Abs(lRow1) from the bottom
' When lRowL=0, ends at last row

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lUR1 As Long, lURL As Long, lUC1 As Long, lUCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lCol1 As Long, lColL As Long

   ArrBoundF fSrc(), lSR1, lSRL, lSC1, lSCL
   NdxGetXSD lSR1, lSRL, lRow1, lRowL, lDR1, lDRL, lUR1, lURL
   NdxGetXSD lSC1, lSCL, lCol1, lColL, lDC1, lDCL, lUC1, lUCL
   ArrAllocF fDst(), lURL, lUCL

   lIRD = lUR1
   For lIRS = lRow1 To lRowL
      lICD = lUC1
      For lICS = lCol1 To lColL
         fDst(lIRD, lICD) = fSrc(lIRS, lICS)
         lICD = lICD + Cl01
      Next lICS
      lIRD = lIRD + Cl01
   Next lIRS

End Sub

Public Sub GetRowLonLon(lSrc() As Long, lDst() As Long, _
                        Optional lRow1 As Long = 0, _
                        Optional lRowL As Long = 0)

' Extract specific or all rows from a long source array to long array
' When lRow1=0, starts at first source row
' When lRow1>0, starts at source row number lRow1
' When lRow1<0, starts at source row Abs(lRow1) from the bottom
' When lRowL=0, ends at last row

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lUR1 As Long, lURL As Long, lUC1 As Long, lUCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lCol1 As Long, lColL As Long

   ArrBoundL lSrc(), lSR1, lSRL, lSC1, lSCL
   NdxGetXSD lSR1, lSRL, lRow1, lRowL, lDR1, lDRL, lUR1, lURL
   NdxGetXSD lSC1, lSCL, lCol1, lColL, lDC1, lDCL, lUC1, lUCL
   ArrAllocL lDst(), lURL, lUCL

   lIRD = lUR1
   For lIRS = lRow1 To lRowL
      lICD = lUC1
      For lICS = lCol1 To lColL
         lDst(lIRD, lICD) = lSrc(lIRS, lICS)
         lICD = lICD + Cl01
      Next lICS
      lIRD = lIRD + Cl01
   Next lIRS

End Sub

Public Sub GetRowStrStr(sSrc() As String, sDst() As String, _
                        Optional lRow1 As Long = 0, _
                        Optional lRowL As Long = 0)

' Extract specific or all rows from a string source array to string array
' When lRow1=0, starts at first source row
' When lRow1>0, starts at source row number lRow1
' When lRow1<0, starts at source row Abs(lRow1) from the bottom
' When lRowL=0, ends at last row

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lUR1 As Long, lURL As Long, lUC1 As Long, lUCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lCol1 As Long, lColL As Long

   ArrBoundS sSrc(), lSR1, lSRL, lSC1, lSCL
   NdxGetXSD lSR1, lSRL, lRow1, lRowL, lDR1, lDRL, lUR1, lURL
   NdxGetXSD lSC1, lSCL, lCol1, lColL, lDC1, lDCL, lUC1, lUCL
   ArrAllocS sDst(), lURL, lUCL

   lIRD = lUR1
   For lIRS = lRow1 To lRowL
      lICD = lUC1
      For lICS = lCol1 To lColL
         sDst(lIRD, lICD) = sSrc(lIRS, lICS)
         lICD = lICD + Cl01
      Next lICS
      lIRD = lIRD + Cl01
   Next lIRS

End Sub

Public Sub GetRowVarDat(vSrc() As Variant, dDst() As Date, _
                        Optional lRow1 As Long = 0, _
                        Optional lRowL As Long = 0, _
                        Optional bList As Boolean = True)

' Extract specific or all rows from a variant source array to dates array
' if bList is set, returns row lRow1 in a 1D list
' When lRow1=0, starts at first source row
' When lRow1>0, starts at source row number lRow1
' When lRow1<0, starts at source row Abs(lRow1) from the bottom

   Const Cl01 As Long = 1
   Dim bDoAll As Boolean

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lUR1 As Long, lURL As Long, lUC1 As Long, lUCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lCol1 As Long, lColL As Long

   ArrBoundV vSrc(), lSR1, lSRL, lSC1, lSCL
   NdxGetXSD lSR1, lSRL, lRow1, lRowL, lDR1, lDRL, lUR1, lURL
   NdxGetXSD lSC1, lSCL, lCol1, lColL, lDC1, lDCL, lUC1, lUCL

   If bList Then                       ' --- Return 1D list
      bDoAll = Not ArrIsAllD(dDst())
      If Not bDoAll Then               ' --- Subset of QRS_LibLst.LstAllD
         bDoAll = Not LBound(dDst()) = lUC1 And UBound(dDst()) = lUCL
      End If
      If bDoAll Then ReDim dDst(1 To lUCL + Cl01 - lUC1)
      lICD = lUC1
      For lICS = lCol1 To lColL
         dDst(lICD) = vSrc(lRow1, lICS)
         lICD = lICD + Cl01
      Next lICS
   Else
      ArrAllocD dDst(), lURL, lUCL
      lIRD = lUR1
      For lIRS = lRow1 To lRowL
         lICD = lUC1
         For lICS = lCol1 To lColL
            dDst(lIRD, lICD) = vSrc(lIRS, lICS)
            lICD = lICD + Cl01
         Next lICS
         lIRD = lIRD + Cl01
      Next lIRS
   End If

End Sub

Public Sub GetRowVarDbl(vSrc() As Variant, fDst() As Double, _
                        Optional lRow1 As Long = 0, _
                        Optional lRowL As Long = 0)

' Extract specific or all rows from a variant source array to long array
' When lRow1=0, starts at first source row
' When lRow1>0, starts at source row number lRow1
' When lRow1<0, starts at source row Abs(lRow1) from the bottom

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lUR1 As Long, lURL As Long, lUC1 As Long, lUCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lCol1 As Long, lColL As Long

   ArrBoundV vSrc(), lSR1, lSRL, lSC1, lSCL
   NdxGetXSD lSR1, lSRL, lRow1, lRowL, lDR1, lDRL, lUR1, lURL
   NdxGetXSD lSC1, lSCL, 0, 0, lDC1, lDCL, lUC1, lUCL
   ArrAllocF fDst(), lURL, lUCL

   lIRD = lUR1
   For lIRS = lRow1 To lRowL
      lICD = lUC1
      For lICS = lCol1 To lColL
         fDst(lIRD, lICD) = vSrc(lIRS, lICS)
         lICD = lICD + Cl01
      Next lICS
      lIRD = lIRD + Cl01
   Next lIRS

End Sub

Public Sub GetRowVarLon(vSrc() As Variant, lDst() As Long, _
                        Optional lRow1 As Long = 0, _
                        Optional lRowL As Long = 0)

' Extract one specific row or all rows from a variant source array
' When lRow1=0, starts at first source row
' When lRow1>0, starts at source row number lRow1
' When lRow1<0, starts at source row Abs(lRow1) from the bottom

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lUR1 As Long, lURL As Long, lUC1 As Long, lUCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lCol1 As Long, lColL As Long

   ArrBoundV vSrc(), lSR1, lSRL, lSC1, lSCL
   NdxGetXSD lSR1, lSRL, lRow1, lRowL, lDR1, lDRL, lUR1, lURL
   NdxGetXSD lSC1, lSCL, 0, 0, lDC1, lDCL, lUC1, lUCL
   ArrAllocL lDst(), lURL, lUCL

   lIRD = lUR1
   For lIRS = lRow1 To lRowL
      lICD = lUC1
      For lICS = lCol1 To lColL
         lDst(lIRD, lICD) = vSrc(lIRS, lICS)
         lICD = lICD + Cl01
      Next lICS
      lIRD = lIRD + Cl01
   Next lIRS

End Sub

Public Sub GetSubDblDbl(fSrc() As Double, fDst() As Double, _
                        Optional lRow1 As Long = 0, _
                        Optional lRowL As Long = 0, _
                        Optional lCol1 As Long = 0, _
                        Optional lColL As Long = 0)

' Extract double sub-array from a double source array
' When lRow1=0, starts at first source row
' When lRow1>0, starts at source row number lRow1
' When lRow1<0, starts at source row Abs(lRow1) from the bottom
' Columns work by analogy

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lUR1 As Long, lURL As Long, lUC1 As Long, lUCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long

   ArrBoundF fSrc(), lSR1, lSRL, lSC1, lSCL
   NdxGetXSD lSR1, lSRL, lRow1, lRowL, lDR1, lDRL, lUR1, lURL
   NdxGetXSD lSC1, lSCL, lCol1, lColL, lDC1, lDCL, lUC1, lUCL
   ArrAllocF fDst(), lURL, lUCL

   lIRD = lUR1
   For lIRS = lRow1 To lRowL
      lICD = lUC1
      For lICS = lCol1 To lColL
         fDst(lIRD, lICD) = fSrc(lIRS, lICS)
         lICD = lICD + Cl01
      Next lICS
      lIRD = lIRD + Cl01
   Next lIRS

End Sub

Public Sub GetSubDblVar(fSrc() As Double, vDst() As Variant, _
                        Optional lRow1 As Long = 0, _
                        Optional lRowL As Long = 0, _
                        Optional lCol1 As Long = 0, _
                        Optional lColL As Long = 0)

' Extract double sub-array from a double source array
' When lRow1=0, starts at first source row
' When lRow1>0, starts at source row number lRow1
' When lRow1<0, starts at source row Abs(lRow1) from the bottom
' Columns work by analogy

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lUR1 As Long, lURL As Long, lUC1 As Long, lUCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long

   ArrBoundF fSrc(), lSR1, lSRL, lSC1, lSCL
   NdxGetXSD lSR1, lSRL, lRow1, lRowL, lDR1, lDRL, lUR1, lURL
   NdxGetXSD lSC1, lSCL, lCol1, lColL, lDC1, lDCL, lUC1, lUCL
   ArrAllocV vDst(), lURL, lUCL

   lIRD = lUR1
   For lIRS = lRow1 To lRowL
      lICD = lUC1
      For lICS = lCol1 To lColL
         vDst(lIRD, lICD) = fSrc(lIRS, lICS)
         lICD = lICD + Cl01
      Next lICS
      lIRD = lIRD + Cl01
   Next lIRS

End Sub

Public Sub GetSubLonLon(lSrc() As Long, lDst() As Long, _
                        Optional lRow1 As Long = 0, _
                        Optional lRowL As Long = 0, _
                        Optional lCol1 As Long = 0, _
                        Optional lColL As Long = 0)

' Extract long integer sub-array from a long integer source array
' When lRow1=0, starts at first source row
' When lRow1>0, starts at source row number lRow1
' When lRow1<0, starts at source row Abs(lRow1) from the bottom
' Columns work by analogy

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lUR1 As Long, lURL As Long, lUC1 As Long, lUCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long

   ArrBoundL lSrc(), lSR1, lSRL, lSC1, lSCL
   NdxGetXSD lSR1, lSRL, lRow1, lRowL, lDR1, lDRL, lUR1, lURL
   NdxGetXSD lSC1, lSCL, lCol1, lColL, lDC1, lDCL, lUC1, lUCL
   ArrAllocL lDst(), lURL, lUCL

   lIRD = lUR1
   For lIRS = lRow1 To lRowL
      lICD = lUC1
      For lICS = lCol1 To lColL
         lDst(lIRD, lICD) = lSrc(lIRS, lICS)
         lICD = lICD + Cl01
      Next lICS
      lIRD = lIRD + Cl01
   Next lIRS

End Sub

Public Sub GetSubStrStr(sSrc() As String, sDst() As String, _
                        Optional lRow1 As Long = 0, _
                        Optional lRowL As Long = 0, _
                        Optional lCol1 As Long = 0, _
                        Optional lColL As Long = 0)

' Extract string sub-array from a string source array
' When lRow1=0, starts at first source row
' When lRow1>0, starts at source row number lRow1
' When lRow1<0, starts at source row Abs(lRow1) from the bottom
' Columns work by analogy

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lUR1 As Long, lURL As Long, lUC1 As Long, lUCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long

   ArrBoundS sSrc(), lSR1, lSRL, lSC1, lSCL
   NdxGetXSD lSR1, lSRL, lRow1, lRowL, lDR1, lDRL, lUR1, lURL
   NdxGetXSD lSC1, lSCL, lCol1, lColL, lDC1, lDCL, lUC1, lUCL
   ArrAllocS sDst(), lURL, lUCL

   lIRD = lUR1
   For lIRS = lRow1 To lRowL
      lICD = lUC1
      For lICS = lCol1 To lColL
         sDst(lIRD, lICD) = sSrc(lIRS, lICS)
         lICD = lICD + Cl01
      Next lICS
      lIRD = lIRD + Cl01
   Next lIRS

End Sub

Public Sub GetSubVarDbl(vSrc() As Variant, fDst() As Double, _
                        Optional lRow1 As Long = 0, _
                        Optional lRowL As Long = 0, _
                        Optional lCol1 As Long = 0, _
                        Optional lColL As Long = 0)

' Extract double sub-array from a variant source array
' When lRow1=0, starts at first source row
' When lRow1>0, starts at source row number lRow1
' When lRow1<0, starts at source row Abs(lRow1) from the bottom
' Columns work by analogy

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lUR1 As Long, lURL As Long, lUC1 As Long, lUCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long

   ArrBoundV vSrc(), lSR1, lSRL, lSC1, lSCL
   NdxGetXSD lSR1, lSRL, lRow1, lRowL, lDR1, lDRL, lUR1, lURL
   NdxGetXSD lSC1, lSCL, lCol1, lColL, lDC1, lDCL, lUC1, lUCL
   ArrAllocF fDst(), lURL, lUCL

   lIRD = lUR1
   For lIRS = lRow1 To lRowL
      lICD = lUC1
      For lICS = lCol1 To lColL
         fDst(lIRD, lICD) = vSrc(lIRS, lICS)
         lICD = lICD + Cl01
      Next lICS
      lIRD = lIRD + Cl01
   Next lIRS

End Sub

Public Sub GetSubVarLon(vSrc() As Variant, lDst() As Long, _
                        Optional lRow1 As Long = 0, _
                        Optional lRowL As Long = 0, _
                        Optional lCol1 As Long = 0, _
                        Optional lColL As Long = 0)

' Extract long integer sub-array from a variant source array
' When lRow1=0, starts at first source row
' When lRow1>0, starts at source row number lRow1
' When lRow1<0, starts at source row Abs(lRow1) from the bottom
' Columns work by analogy

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lUR1 As Long, lURL As Long, lUC1 As Long, lUCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long

   ArrBoundV vSrc(), lSR1, lSRL, lSC1, lSCL
   NdxGetXSD lSR1, lSRL, lRow1, lRowL, lDR1, lDRL, lUR1, lURL
   NdxGetXSD lSC1, lSCL, lCol1, lColL, lDC1, lDCL, lUC1, lUCL
   ArrAllocL lDst(), lURL, lUCL

   lIRD = lUR1
   For lIRS = lRow1 To lRowL
      lICD = lUC1
      For lICS = lCol1 To lColL
         lDst(lIRD, lICD) = vSrc(lIRS, lICS)
         lICD = lICD + Cl01
      Next lICS
      lIRD = lIRD + Cl01
   Next lIRS

End Sub

Public Sub GetSubVarStr(vSrc() As Variant, sDst() As String, _
                        Optional lRow1 As Long = 0, _
                        Optional lRowL As Long = 0, _
                        Optional lCol1 As Long = 0, _
                        Optional lColL As Long = 0)

' Extract string sub-array from a variant source array
' When lRow1=0, starts at first source row
' When lRow1>0, starts at source row number lRow1
' When lRow1<0, starts at source row Abs(lRow1) from the bottom
' Columns work by analogy

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lUR1 As Long, lURL As Long, lUC1 As Long, lUCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long

   ArrBoundV vSrc(), lSR1, lSRL, lSC1, lSCL
   NdxGetXSD lSR1, lSRL, lRow1, lRowL, lDR1, lDRL, lUR1, lURL
   NdxGetXSD lSC1, lSCL, lCol1, lColL, lDC1, lDCL, lUC1, lUCL
   ArrAllocS sDst(), lURL, lUCL

   lIRD = lUR1
   For lIRS = lRow1 To lRowL
      lICD = lUC1
      For lICS = lCol1 To lColL
         sDst(lIRD, lICD) = vSrc(lIRS, lICS)
         lICD = lICD + Cl01
      Next lICS
      lIRD = lIRD + Cl01
   Next lIRS

End Sub

Public Sub GetSubVarVar(vSrc() As Variant, vDst() As Variant, _
                        Optional lRow1 As Long = 0, _
                        Optional lRowL As Long = 0, _
                        Optional lCol1 As Long = 0, _
                        Optional lColL As Long = 0)

' Extract variant sub-array from a variant source array
' When lRow1=0, starts at first source row
' When lRow1>0, starts at source row number lRow1
' When lRow1<0, starts at source row Abs(lRow1) from the bottom
' Columns work by analogy

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lUR1 As Long, lURL As Long, lUC1 As Long, lUCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long

   ArrBoundV vSrc(), lSR1, lSRL, lSC1, lSCL
   NdxGetXSD lSR1, lSRL, lRow1, lRowL, lDR1, lDRL, lUR1, lURL
   NdxGetXSD lSC1, lSCL, lCol1, lColL, lDC1, lDCL, lUC1, lUCL
   ArrAllocV vDst(), lURL, lUCL

   lIRD = lUR1
   For lIRS = lRow1 To lRowL
      lICD = lUC1
      For lICS = lCol1 To lColL
         vDst(lIRD, lICD) = vSrc(lIRS, lICS)
         lICD = lICD + Cl01
      Next lICS
      lIRD = lIRD + Cl01
   Next lIRS

End Sub

Public Sub PutColDatVar(dSrc() As Date, vDst() As Variant, _
                        Optional lCol1 As Long = 0, _
                        Optional lColL As Long = 0)

' Output a specific or all columns from a date source array to variant array
' When lCol1=0, output starts et first destination column
' When lCol1>0, output starts at destination column lCol1
' When lCol1<0, output starts at destination column Abs(lCol1) from the right
' When lColL=0, output ends at last column
' dSrc is a 2D array. For 1D lists, use the QRS_LibA2L.PutColDatVar() routine

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lVR1 As Long, lVRL As Long, lVC1 As Long, lVCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lRow1 As Long, lRowL As Long

   ArrBoundD dSrc(), lSR1, lSRL, lSC1, lSCL
   ArrBoundV vDst(), lDR1, lDRL, lDC1, lDCL
   NdxGetXDS lDC1, lDCL, lSC1, lSCL, lCol1, lColL, lVC1, lVCL
   NdxGetXDS lDR1, lDRL, lSR1, lSRL, lRow1, lRowL, lVR1, lVRL

   lICS = lVC1
   For lICD = lCol1 To lColL
      lIRS = lVR1
      For lIRD = lRow1 To lRowL
         vDst(lIRD, lICD) = dSrc(lIRS, lICS)
         lIRS = lIRS + Cl01
      Next lIRD
      lICS = lICS + Cl01
   Next lICD

End Sub

Public Sub PutColDblDbl(fSrc() As Double, fDst() As Double, _
                        Optional lCol1 As Long = 0, _
                        Optional lColL As Long = 0)

' Output a specific or all columns from a double source array to double array
' When lCol1=0, output starts et first destination column
' When lCol1>0, output starts at destination column lCol1
' When lCol1<0, output starts at destination column Abs(lCol1) from the right
' When lColL=0, output ends at last column

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lVR1 As Long, lVRL As Long, lVC1 As Long, lVCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lRow1 As Long, lRowL As Long

   ArrBoundF fSrc(), lSR1, lSRL, lSC1, lSCL
   ArrBoundF fDst(), lDR1, lDRL, lDC1, lDCL
   NdxGetXDS lDC1, lDCL, lSC1, lSCL, lCol1, lColL, lVC1, lVCL
   NdxGetXDS lDR1, lDRL, lSR1, lSRL, lRow1, lRowL, lVR1, lVRL

   lICS = lVC1
   For lICD = lCol1 To lColL
      lIRS = lVR1
      For lIRD = lRow1 To lRowL
         fDst(lIRD, lICD) = fSrc(lIRS, lICS)
         lIRS = lIRS + Cl01
      Next lIRD
      lICS = lICS + Cl01
   Next lICD

End Sub

Public Sub PutRowDblVar(fSrc() As Double, vDst() As Variant, _
                        Optional lRow1 As Long = 0, _
                        Optional lRowL As Long = 0)

' Output a specific or all columns from a double source array to variant array
' When lCol1=0, output starts et first destination column
' When lCol1>0, output starts at destination column lCol1
' When lCol1<0, output starts at destination column Abs(lCol1) from the right
' When lColL=0, output ends at last column

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lVR1 As Long, lVRL As Long, lVC1 As Long, lVCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lCol1 As Long, lColL As Long

   ArrBoundF fSrc(), lSR1, lSRL, lSC1, lSCL
   ArrBoundV vDst(), lDR1, lDRL, lDC1, lDCL
   NdxGetXDS lDC1, lDCL, lSC1, lSCL, lCol1, lColL, lVC1, lVCL
   NdxGetXDS lDR1, lDRL, lSR1, lSRL, lRow1, lRowL, lVR1, lVRL

   lIRS = lVR1
   For lIRD = lRow1 To lRowL
      lICS = lVC1
      For lICD = lCol1 To lColL
         vDst(lIRD, lICD) = fSrc(lIRS, lICS)
         lICS = lICS + Cl01
      Next lICD
      lIRS = lIRS + Cl01
   Next lIRD

End Sub

Public Sub PutColLonLon(lSrc() As Long, lDst() As Long, _
                        Optional lCol1 As Long = 0, _
                        Optional lColL As Long = 0)

' Output a specific or all columns from a double source array to double array
' When lCol1=0, output starts et first destination column
' When lCol1>0, output starts at destination column lCol1
' When lCol1<0, output starts at destination column Abs(lCol1) from the right
' When lColL=0, output ends at last column

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lVR1 As Long, lVRL As Long, lVC1 As Long, lVCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lRow1 As Long, lRowL As Long

   ArrBoundL lSrc(), lSR1, lSRL, lSC1, lSCL
   ArrBoundL lDst(), lDR1, lDRL, lDC1, lDCL
   NdxGetXDS lDC1, lDCL, lSC1, lSCL, lCol1, lColL, lVC1, lVCL
   NdxGetXDS lDR1, lDRL, lSR1, lSRL, lRow1, lRowL, lVR1, lVRL

   lICS = lVC1
   For lICD = lCol1 To lColL
      lIRS = lVR1
      For lIRD = lRow1 To lRowL
         lDst(lIRD, lICD) = lSrc(lIRS, lICS)
         lIRS = lIRS + Cl01
      Next lIRD
      lICS = lICS + Cl01
   Next lICD

End Sub

Public Sub PutColStrStr(sSrc() As String, sDst() As String, _
                        Optional lCol1 As Long = 0, _
                        Optional lColL As Long = 0)

' Output a specific or all columns from a double source array to double array
' When lCol1=0, output starts et first destination column
' When lCol1>0, output starts at destination column lCol1
' When lCol1<0, output starts at destination column Abs(lCol1) from the right
' When lColL=0, output ends at last column

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lVR1 As Long, lVRL As Long, lVC1 As Long, lVCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lRow1 As Long, lRowL As Long

   ArrBoundS sSrc(), lSR1, lSRL, lSC1, lSCL
   ArrBoundS sDst(), lDR1, lDRL, lDC1, lDCL
   NdxGetXDS lDC1, lDCL, lSC1, lSCL, lCol1, lColL, lVC1, lVCL
   NdxGetXDS lDR1, lDRL, lSR1, lSRL, lRow1, lRowL, lVR1, lVRL

   lICS = lVC1
   For lICD = lCol1 To lColL
      lIRS = lVR1
      For lIRD = lRow1 To lRowL
         sDst(lIRD, lICD) = sSrc(lIRS, lICS)
         lIRS = lIRS + Cl01
      Next lIRD
      lICS = lICS + Cl01
   Next lICD

End Sub

Public Sub PutRowDblDbl(fSrc() As Double, fDst() As Double, _
                        Optional lRow1 As Long = 0, _
                        Optional lRowL As Long = 0)

' Output a specific or all columns from a double source array to double array
' When lCol1=0, output starts et first destination column
' When lCol1>0, output starts at destination column lCol1
' When lCol1<0, output starts at destination column Abs(lCol1) from the right
' When lColL=0, output ends at last column

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lVR1 As Long, lVRL As Long, lVC1 As Long, lVCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lCol1 As Long, lColL As Long

   ArrBoundF fSrc(), lSR1, lSRL, lSC1, lSCL
   ArrBoundF fDst(), lDR1, lDRL, lDC1, lDCL
   NdxGetXDS lDC1, lDCL, lSC1, lSCL, lCol1, lColL, lVC1, lVCL
   NdxGetXDS lDR1, lDRL, lSR1, lSRL, lRow1, lRowL, lVR1, lVRL

   lIRS = lVR1
   For lIRD = lRow1 To lRowL
      lICS = lVC1
      For lICD = lCol1 To lColL
         fDst(lIRD, lICD) = fSrc(lIRS, lICS)
         lICS = lICS + Cl01
      Next lICD
      lIRS = lIRS + Cl01
   Next lIRD

End Sub

Public Sub PutRowLonLon(lSrc() As Long, lDst() As Long, _
                        Optional lRow1 As Long = 0, _
                        Optional lRowL As Long = 0)

' Output a specific or all columns from a long source array to long array
' When lCol1=0, output starts et first destination column
' When lCol1>0, output starts at destination column lCol1
' When lCol1<0, output starts at destination column Abs(lCol1) from the right
' When lColL=0, output ends at last column

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lVR1 As Long, lVRL As Long, lVC1 As Long, lVCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lCol1 As Long, lColL As Long

   ArrBoundL lSrc(), lSR1, lSRL, lSC1, lSCL
   ArrBoundL lDst(), lDR1, lDRL, lDC1, lDCL
   NdxGetXDS lDC1, lDCL, lSC1, lSCL, lCol1, lColL, lVC1, lVCL
   NdxGetXDS lDR1, lDRL, lSR1, lSRL, lRow1, lRowL, lVR1, lVRL

   lIRS = lVR1
   For lIRD = lRow1 To lRowL
      lICS = lVC1
      For lICD = lCol1 To lColL
         lDst(lIRD, lICD) = lSrc(lIRS, lICS)
         lICS = lICS + Cl01
      Next lICD
      lIRS = lIRS + Cl01
   Next lIRD

End Sub

Public Sub PutRowLonVar(lSrc() As Long, vDst() As Variant, _
                        Optional lRow1 As Long = 0, _
                        Optional lRowL As Long = 0)

' Output a specific or all columns from a long source array to variant array
' When lCol1=0, output starts et first destination column
' When lCol1>0, output starts at destination column lCol1
' When lCol1<0, output starts at destination column Abs(lCol1) from the right
' When lColL=0, output ends at last column

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lVR1 As Long, lVRL As Long, lVC1 As Long, lVCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lCol1 As Long, lColL As Long

   ArrBoundL lSrc(), lSR1, lSRL, lSC1, lSCL
   ArrBoundV vDst(), lDR1, lDRL, lDC1, lDCL
   NdxGetXDS lDC1, lDCL, lSC1, lSCL, lCol1, lColL, lVC1, lVCL
   NdxGetXDS lDR1, lDRL, lSR1, lSRL, lRow1, lRowL, lVR1, lVRL

   lIRS = lVR1
   For lIRD = lRow1 To lRowL
      lICS = lVC1
      For lICD = lCol1 To lColL
         vDst(lIRD, lICD) = lSrc(lIRS, lICS)
         lICS = lICS + Cl01
      Next lICD
      lIRS = lIRS + Cl01
   Next lIRD

End Sub

Public Sub PutRowStrVar(sSrc() As String, vDst() As Variant, _
                        Optional lRow1 As Long = 0, _
                        Optional lRowL As Long = 0)

' Output a specific or all columns from a string source array to variant array
' When lCol1=0, output starts et first destination column
' When lCol1>0, output starts at destination column lCol1
' When lCol1<0, output starts at destination column Abs(lCol1) from the right
' When lColL=0, output ends at last column

   Const Cl01 As Long = 1

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lDR1 As Long, lDRL As Long, lDC1 As Long, lDCL As Long
   Dim lVR1 As Long, lVRL As Long, lVC1 As Long, lVCL As Long
   Dim lIRS As Long, lICS As Long, lIRD As Long, lICD As Long
   Dim lCol1 As Long, lColL As Long

   ArrBoundS sSrc(), lSR1, lSRL, lSC1, lSCL
   ArrBoundV vDst(), lDR1, lDRL, lDC1, lDCL
   NdxGetXDS lDC1, lDCL, lSC1, lSCL, lCol1, lColL, lVC1, lVCL
   NdxGetXDS lDR1, lDRL, lSR1, lSRL, lRow1, lRowL, lVR1, lVRL

   lIRS = lVR1
   For lIRD = lRow1 To lRowL
      lICS = lVC1
      For lICD = lCol1 To lColL
         vDst(lIRD, lICD) = sSrc(lIRS, lICS)
         lICS = lICS + Cl01
      Next lICD
      lIRS = lIRS + Cl01
   Next lIRD

End Sub

Public Function ArrMergeF(fArrA() As Double, fArrB() As Double, _
                          fArrC() As Double, _
                          Optional fMulA As Double = 1, _
                          Optional fMulB As Double = 1) As Boolean

' Merges values of fArrA() and fArrB() into fArrC() by linear operation
' on each element fArrC(r,c) = fArrA(r,c) * fMulA + fArrB(r,c) * fMulB
' If bounds of fArrA() and fArrB() do not match, aborts and returns true

   Dim lAR1 As Long, lARL As Long, lAC1 As Long, lACL As Long
   Dim lBR1 As Long, lBRL As Long, lBC1 As Long, lBCL As Long
   Dim lIR As Long, lIC As Long
   Dim bFail As Boolean

   bFail = Not (ArrIsAllF(fArrA()) And ArrIsAllF(fArrB()))
   If bFail Then GoTo Ende

   ArrBoundF fArrA(), lAR1, lARL, lAC1, lACL
   ArrBoundF fArrB(), lBR1, lBRL, lBC1, lBCL
   bFail = Not (lAR1 = lBR1 And lARL = lBRL And lAC1 = lBC1 And lACL = lBCL)
   If bFail Then GoTo Ende

   ArrAllocF fArrC(), lARL, lACL

   For lIR = lAR1 To lARL
      For lIC = lAC1 To lACL
         fArrC(lIR, lIC) = fArrA(lIR, lIC) * fMulA + fArrB(lIR, lIC) * fMulB
      Next lIC
   Next lIR

Ende:

   ArrMergeF = bFail

End Function

Public Function ArrMergeL(lArrA() As Long, lArrB() As Long, _
                          lArrC() As Long, _
                          Optional lMulA As Long = 1, _
                          Optional lMulB As Long = 1) As Boolean

' Merges values of lArrA() and lArrB() into lArrC() by linear operation
' on each element lArrC(r,c) = lArrA(r,c) * lMulA + lArrB(r,c) * lMulB
' If bounds of lArrA() and lArrB() do not match, aborts and returns true

   Dim lAR1 As Long, lARL As Long, lAC1 As Long, lACL As Long
   Dim lBR1 As Long, lBRL As Long, lBC1 As Long, lBCL As Long
   Dim lIR As Long, lIC As Long
   Dim bFail As Boolean

   bFail = Not (ArrIsAllL(lArrA()) And ArrIsAllL(lArrB()))
   If bFail Then GoTo Ende

   ArrBoundL lArrA(), lAR1, lARL, lAC1, lACL
   ArrBoundL lArrB(), lBR1, lBRL, lBC1, lBCL
   bFail = Not (lAR1 = lBR1 And lARL = lBRL And lAC1 = lBC1 And lACL = lBCL)
   If bFail Then GoTo Ende

   ArrAllocL lArrC(), lARL, lACL

   For lIR = lAR1 To lARL
      For lIC = lAC1 To lACL
         lArrC(lIR, lIC) = lArrA(lIR, lIC) * lMulA + lArrB(lIR, lIC) * lMulB
      Next lIC
   Next lIR

Ende:

   ArrMergeL = bFail

End Function

Public Function ArrMergeS(sArrA() As String, sArrB() As String, _
                          sArrC() As String, _
                          Optional bUseA As Boolean = True, _
                          Optional bAppB As Boolean = True) As Boolean

' Merges values of sArrA() and sArrB() into sArrC() by string appending
' on each element sArrC(r,c) = if bUseA sArrA(r,c) & if bAppA sArrB(r,c)
' if both bUseA and bUseB are false, clears C, which has an effect when
' sArrC is the same as either A or B
' If bounds of sArrA() and sArrB() do not match, aborts and returns true

   Dim lAR1 As Long, lARL As Long, lAC1 As Long, lACL As Long
   Dim lBR1 As Long, lBRL As Long, lBC1 As Long, lBCL As Long
   Dim lIR As Long, lIC As Long
   Dim bFail As Boolean

   bFail = Not (ArrIsAllS(sArrA()) And ArrIsAllS(sArrB()))
   If bFail Then GoTo Ende

   ArrBoundS sArrA(), lAR1, lARL, lAC1, lACL
   ArrBoundS sArrB(), lBR1, lBRL, lBC1, lBCL
   bFail = Not (lAR1 = lBR1 And lARL = lBRL And lAC1 = lBC1 And lACL = lBCL)
   If bFail Then GoTo Ende

   ArrAllocS sArrC(), lARL, lACL

   If bUseA Then                       ' --- Keep A
      If bAppB Then                    ' --- A & B
         For lIR = lAR1 To lARL
            For lIC = lAC1 To lACL
               sArrC(lIR, lIC) = sArrA(lIR, lIC) & sArrB(lIR, lIC)
            Next lIC
         Next lIR
      Else                             '     Just A
         For lIR = lAR1 To lARL
            For lIC = lAC1 To lACL
               sArrC(lIR, lIC) = sArrA(lIR, lIC)
            Next lIC
         Next lIR
      End If
   Else                                ' --- Ignore A
      If bAppB Then                    ' --- Just B
         For lIR = lAR1 To lARL
            For lIC = lAC1 To lACL
               sArrC(lIR, lIC) = sArrB(lIR, lIC)
            Next lIC
         Next lIR
      Else                             '     Clear
         For lIR = lAR1 To lARL
            For lIC = lAC1 To lACL
               sArrC(lIR, lIC) = ""
            Next lIC
         Next lIR
      End If
   End If

Ende:

   ArrMergeS = bFail

End Function

Public Function ArrMergeV(vArrA() As Variant, vArrB() As Variant, _
                          vArrC() As Variant, _
                          Optional fMulA As Double = 1, _
                          Optional fMulB As Double = 1) As Boolean

' Merges values of vArrA() and vArrB() into vArrC() by operations depending
' on the variable type. It is assumed that for all corresponding elements
' the variable types are the same.
'    for dates and if fMulA and fMulB are bot non-zero, vArrB is used
'    for numeric data, the linear operations are applied
'    for string data, the logical equivalent is fMulA <> 0
' If bounds of sArrA() and sArrB() do not match, aborts and returns true

   Dim lAR1 As Long, lARL As Long, lAC1 As Long, lACL As Long
   Dim lBR1 As Long, lBRL As Long, lBC1 As Long, lBCL As Long
   Dim lIR As Long, lIC As Long
   Dim bUA As Boolean, bAB As Boolean
   Dim lVT As VbVarType
   Dim bFail As Boolean

   bFail = Not (ArrIsAllV(vArrA()) And ArrIsAllV(vArrB()))
   If bFail Then GoTo Ende

   ArrBoundV vArrA(), lAR1, lARL, lAC1, lACL
   ArrBoundV vArrB(), lBR1, lBRL, lBC1, lBCL
   bFail = Not (lAR1 = lBR1 And lARL = lBRL And lAC1 = lBC1 And lACL = lBCL)
   If bFail Then GoTo Ende

   ArrAllocV vArrC(), lARL, lACL

   bUA = Not fMulA = 0
   bAB = Not fMulB = 0
   For lIR = lAR1 To lARL
      For lIC = lAC1 To lACL
         lVT = VarType(vArrB(lIR, lIC))
         Select Case lVT
         Case vbDouble, vbLong         ' --- Numbers
            vArrC(lIR, lIC) = vArrA(lIR, lIC) * fMulA + vArrB(lIR, lIC) * fMulB
         Case vbString                 ' --- String
            If bUA Then                '     A
               If bAB Then             '     A & B
                  vArrC(lIR, lIC) = vArrA(lIR, lIC) & vArrB(lIR, lIC)
               Else                    '     A
                  vArrC(lIR, lIC) = vArrA(lIR, lIC)
               End If
            Else                       '     B
               If bAB Then             '     B
                  vArrC(lIR, lIC) = vArrB(lIR, lIC)
               Else                    '     None
                  vArrC(lIR, lIC) = ""
               End If
            End If
         Case vbDate                   ' --- Dates: Always B
            If bAB Then vArrC(lIR, lIC) = vArrB(lIR, lIC)
         End Select
      Next lIC
   Next lIR

Ende:

   ArrMergeV = bFail

End Function

Public Function XtrEleVarAbs(vArr(), _
                             Optional lRow1 As Long, Optional lCol1 As Long, _
                             Optional A1, _
                             Optional lRow2 As Long, Optional lCol2 As Long, _
                             Optional A2, _
                             Optional lRow3 As Long, Optional lCol3 As Long, _
                             Optional A3, _
                             Optional lRow4 As Long, Optional lCol4 As Long, _
                             Optional A4) As Boolean

' Extract up to 4 elements from an array by individual row and column
' row and col may be specified using the same rules as for the Get functions
' Returns true if any index is out of bounds

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lSRX As Long, lSCX As Long
   Dim bErr As Boolean, bFail As Boolean

   ArrBoundV vArr(), lSR1, lSRL, lSC1, lSCL

   If Not IsMissing(A1) Then
      lSRX = NdxGetSX1(lSR1, lSRL, lRow1)
      lSCX = NdxGetSX1(lSC1, lSCL, lCol1)
      bErr = lSRX < lSR1 Or lSRX > lSRL Or lSCX < lSC1 Or lSCX > lSCL
      bFail = bFail Or bErr
      If Not bErr Then A1 = vArr(lSRX, lSCX)
   End If

   If Not IsMissing(A2) Then
      lSRX = NdxGetSX1(lSR1, lSRL, lRow2)
      lSCX = NdxGetSX1(lSC1, lSCL, lCol2)
      bErr = lSRX < lSR1 Or lSRX > lSRL Or lSCX < lSC1 Or lSCX > lSCL
      bFail = bFail Or bErr
      If Not bErr Then A2 = vArr(lSRX, lSCX)
   End If

   If Not IsMissing(A3) Then
      lSRX = NdxGetSX1(lSR1, lSRL, lRow3)
      lSCX = NdxGetSX1(lSC1, lSCL, lCol3)
      bErr = lSRX < lSR1 Or lSRX > lSRL Or lSCX < lSC1 Or lSCX > lSCL
      bFail = bFail Or bErr
      If Not bErr Then A3 = vArr(lSRX, lSCX)
   End If

   If Not IsMissing(A4) Then
      lSRX = NdxGetSX1(lSR1, lSRL, lRow4)
      lSCX = NdxGetSX1(lSC1, lSCL, lCol4)
      bErr = lSRX < lSR1 Or lSRX > lSRL Or lSCX < lSC1 Or lSCX > lSCL
      bFail = bFail Or bErr
      If Not bErr Then A4 = vArr(lSRX, lSCX)
   End If

Ende:

   XtrEleVarAbs = bFail

End Function

Public Function XtrEleVarRel(vArr(), _
                             lRow1 As Long, lCol1 As Long, _
                             lIncR As Long, lIncC As Long, _
                             Optional A1, Optional A2, _
                             Optional A3, Optional A4) As Boolean

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lSRX As Long, lSCX As Long
   Dim bErr As Boolean, bFail As Boolean

   ArrBoundV vArr(), lSR1, lSRL, lSC1, lSCL
   lSRX = NdxGetSX1(lSR1, lSRL, lRow1)
   lSCX = NdxGetSX1(lSC1, lSCL, lCol1)

   If Not IsMissing(A1) Then
      bErr = lSRX < lSR1 Or lSRX > lSRL Or lSCX < lSC1 Or lSCX > lSCL
      bFail = bFail Or bErr
      If Not bErr Then A1 = vArr(lSRX, lSCX)
   End If

   If Not IsMissing(A2) Then
      lSRX = lSRX + lIncR
      lSCX = lSCX + lIncC
      bErr = lSRX < lSR1 Or lSRX > lSRL Or lSCX < lSC1 Or lSCX > lSCL
      bFail = bFail Or bErr
      If Not bErr Then A2 = vArr(lSRX, lSCX)
   End If

   If Not IsMissing(A3) Then
      lSRX = lSRX + lIncR
      lSCX = lSCX + lIncC
      bErr = lSRX < lSR1 Or lSRX > lSRL Or lSCX < lSC1 Or lSCX > lSCL
      bFail = bFail Or bErr
      If Not bErr Then A3 = vArr(lSRX, lSCX)
   End If

   If Not IsMissing(A4) Then
      lSRX = lSRX + lIncR
      lSCX = lSCX + lIncC
      bErr = lSRX < lSR1 Or lSRX > lSRL Or lSCX < lSC1 Or lSCX > lSCL
      bFail = bFail Or bErr
      If Not bErr Then A4 = vArr(lSRX, lSCX)
   End If

Ende:

   XtrEleVarRel = bFail

End Function

Public Function XtrEleVarCol(vArr(), lCol1 As Long, _
                             Optional lRow1 As Long = 1, _
                             Optional A1, Optional A2, Optional A3, _
                             Optional A4, Optional A5, Optional A6) As Boolean

' Extract elements from a column of a variant array
' Col1 and row1 may be specified using the same rules as for the Get functions
' Returns true if any of the indices is out of bounds

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lSRX As Long
   Dim bErr As Boolean, bFail As Boolean

   ArrBoundV vArr(), lSR1, lSRL, lSC1, lSCL
   lRow1 = NdxGetSX1(lSR1, lSRL, lRow1)
   lCol1 = NdxGetSX1(lSC1, lSCL, lCol1)

   bFail = lCol1 < lSC1 Or lCol1 > lSCL
   If bFail Then GoTo Ende

   If Not IsMissing(A1) Then
      lSRX = lRow1
      bErr = lSRX < lSR1 Or lSRX > lSRL
      bFail = bFail Or bErr
      If Not bErr Then A1 = vArr(lSRX, lCol1)
   End If

   If Not IsMissing(A2) Then
      lSRX = lRow1 + 1
      bErr = lSRX < lSR1 Or lSRX > lSRL
      bFail = bFail Or bErr
      If Not bErr Then A2 = vArr(lSRX, lCol1)
   End If

   If Not IsMissing(A3) Then
      lSRX = lRow1 + 2
      bErr = lSRX < lSR1 Or lSRX > lSRL
      bFail = bFail Or bErr
      If Not bErr Then A3 = vArr(lSRX, lCol1)
   End If

   If Not IsMissing(A4) Then
      lSRX = lRow1 + 3
      bErr = lSRX < lSR1 Or lSRX > lSRL
      bFail = bFail Or bErr
      If Not bErr Then A4 = vArr(lSRX, lCol1)
   End If

   If Not IsMissing(A5) Then
      lSRX = lRow1 + 4
      bErr = lSRX < lSR1 Or lSRX > lSRL
      bFail = bFail Or bErr
      If Not bErr Then A5 = vArr(lSRX, lCol1)
   End If

   If Not IsMissing(A6) Then
      lSRX = lRow1 + 5
      bErr = lSRX < lSR1 Or lSRX > lSRL
      bFail = bFail Or bErr
      If Not bErr Then A6 = vArr(lSRX, lCol1)
   End If

Ende:

   XtrEleVarCol = bFail

End Function

Public Function XtrEleVarRow(vArr(), lRow1 As Long, _
                             Optional lCol1 As Long = 1, _
                             Optional A1, Optional A2, Optional A3, _
                             Optional A4, Optional A5, Optional A6) As Boolean

' Extract elements from a row of a variant array
' Row1 and Col1 may be specified using the same rules as for the Get function
' Returns true if any of the indices is out of bounds

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lSCX As Long
   Dim bErr As Boolean, bFail As Boolean

   ArrBoundV vArr(), lSR1, lSRL, lSC1, lSCL
   lRow1 = NdxGetSX1(lSR1, lSRL, lRow1)
   lCol1 = NdxGetSX1(lSC1, lSCL, lCol1)

   bFail = lRow1 < lSR1 Or lRow1 > lSRL
   If bFail Then GoTo Ende

   If Not IsMissing(A1) Then
      lSCX = lCol1
      bErr = lSCX < lSC1 Or lSCX > lSCL
      bFail = bFail Or bErr
      If Not bFail Then A1 = vArr(lRow1, lSCX)
   End If

   If Not IsMissing(A2) Then
      lSCX = lCol1 + 1
      bFail = lSCX < lSC1 Or lSCX > lSCL
      bFail = bFail Or bErr
      If Not bFail Then A2 = vArr(lRow1, lSCX)
   End If

   If Not IsMissing(A3) Then
      lSCX = lCol1 + 2
      bFail = lSCX < lSC1 Or lSCX > lSCL
      bFail = bFail Or bErr
      If Not bFail Then A3 = vArr(lRow1, lSCX)
   End If

   If Not IsMissing(A4) Then
      lSCX = lCol1 + 3
      bFail = lSCX < lSC1 Or lSCX > lSCL
      bFail = bFail Or bErr
      If Not bFail Then A4 = vArr(lRow1, lSCX)
   End If

   If Not IsMissing(A5) Then
      lSCX = lCol1 + 4
      bFail = lSCX < lSC1 Or lSCX > lSCL
      bFail = bFail Or bErr
      If Not bFail Then A5 = vArr(lRow1, lSCX)
   End If

   If Not IsMissing(A6) Then
      lSCX = lCol1 + 5
      bFail = lSCX < lSC1 Or lSCX > lSCL
      bFail = bFail Or bErr
      If Not bFail Then A6 = vArr(lRow1, lSCX)
   End If

Ende:

   XtrEleVarRow = bFail

End Function

Public Function ZipEleVarCol(vArr(), lCol1 As Long, _
                             Optional lRow1 As Long = 1, _
                             Optional A1, Optional A2, Optional A3, _
                             Optional A4, Optional A5, Optional A6) As Boolean

' Insert elements into a row of a variant array
' Row1 and Col1 may be specified using the same rules as for the Get function
' Returns true if any of the indices is out of bounds

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lSRX As Long
   Dim bErr As Boolean, bFail As Boolean

   ArrBoundV vArr(), lSR1, lSRL, lSC1, lSCL
   lRow1 = NdxGetSX1(lSR1, lSRL, lRow1)
   lCol1 = NdxGetSX1(lSC1, lSCL, lCol1)

   bFail = lCol1 < lSC1 Or lCol1 > lSCL
   If bFail Then GoTo Ende

   If Not IsMissing(A1) Then
      lSRX = lRow1
      bErr = lSRX < lSR1 Or lSRX > lSRL
      bFail = bFail Or bErr
      If Not bFail Then vArr(lSRX, lCol1) = A1
   End If

   If Not IsMissing(A2) Then
      lSRX = lRow1 + 1
      bErr = lSRX < lSR1 Or lSRX > lSRL
      bFail = bFail Or bErr
      If Not bFail Then vArr(lSRX, lCol1) = A2
   End If

   If Not IsMissing(A3) Then
      lSRX = lRow1 + 2
      bErr = lSRX < lSR1 Or lSRX > lSRL
      bFail = bFail Or bErr
      If Not bFail Then vArr(lSRX, lCol1) = A3
   End If

   If Not IsMissing(A4) Then
      lSRX = lRow1 + 3
      bErr = lSRX < lSR1 Or lSRX > lSRL
      bFail = bFail Or bErr
      If Not bFail Then vArr(lSRX, lCol1) = A4
   End If

   If Not IsMissing(A5) Then
      lSRX = lRow1 + 4
      bErr = lSRX < lSR1 Or lSRX > lSRL
      bFail = bFail Or bErr
      If Not bFail Then vArr(lSRX, lCol1) = A5
   End If

   If Not IsMissing(A6) Then
      lSRX = lRow1 + 5
      bErr = lSRX < lSR1 Or lSRX > lSRL
      bFail = bFail Or bErr
      If Not bFail Then vArr(lSRX, lCol1) = A6
   End If

Ende:

   ZipEleVarCol = bFail

End Function

Public Function ZipEleVarRow(vArr(), lRow1 As Long, _
                             Optional lCol1 As Long = 1, _
                             Optional A1, Optional A2, Optional A3, _
                             Optional A4, Optional A5, Optional A6) As Boolean

' Insert elements into a row of a variant array
' Row1 and Col1 may be specified using the same rules as for the Get function
' Returns true if any of the indices is out of bounds

   Dim lSR1 As Long, lSRL As Long, lSC1 As Long, lSCL As Long
   Dim lSCX As Long
   Dim bErr As Boolean, bFail As Boolean

   ArrBoundV vArr(), lSR1, lSRL, lSC1, lSCL
   lRow1 = NdxGetSX1(lSR1, lSRL, lRow1)
   lCol1 = NdxGetSX1(lSC1, lSCL, lCol1)

   bFail = lRow1 < lSR1 Or lRow1 > lSRL
   If bFail Then GoTo Ende

   If Not IsMissing(A1) Then
      lSCX = lCol1
      bErr = lSCX < lSC1 Or lSCX > lSCL
      bFail = bFail Or bErr
      If Not bFail Then vArr(lRow1, lSCX) = A1
   End If

   If Not IsMissing(A2) Then
      lSCX = lCol1 + 1
      bFail = lSCX < lSC1 Or lSCX > lSCL
      bFail = bFail Or bErr
      If Not bFail Then vArr(lRow1, lSCX) = A2
   End If

   If Not IsMissing(A3) Then
      lSCX = lCol1 + 2
      bFail = lSCX < lSC1 Or lSCX > lSCL
      bFail = bFail Or bErr
      If Not bFail Then vArr(lRow1, lSCX) = A3
   End If

   If Not IsMissing(A4) Then
      lSCX = lCol1 + 3
      bFail = lSCX < lSC1 Or lSCX > lSCL
      bFail = bFail Or bErr
      If Not bFail Then vArr(lRow1, lSCX) = A4
   End If

   If Not IsMissing(A5) Then
      lSCX = lCol1 + 4
      bFail = lSCX < lSC1 Or lSCX > lSCL
      bFail = bFail Or bErr
      If Not bFail Then vArr(lRow1, lSCX) = A5
   End If

   If Not IsMissing(A6) Then
      lSCX = lCol1 + 5
      bFail = lSCX < lSC1 Or lSCX > lSCL
      bFail = bFail Or bErr
      If Not bFail Then vArr(lRow1, lSCX) = A6
   End If

Ende:

   ZipEleVarRow = bFail

End Function

Public Function ArrIsAllD(dArr() As Date) As Boolean

' Returns true if date array has been allocated

   ArrIsAllD = Not Not dArr()

End Function

Public Function ArrIsAllF(fArr() As Double) As Boolean

' Returns true if double array has been allocated

   ArrIsAllF = Not Not fArr()

End Function

Public Function ArrIsAllL(lArr() As Long) As Boolean

' Returns true if long array has been allocated

   ArrIsAllL = Not Not lArr()

End Function

Public Function ArrIsAllS(sArr() As String) As Boolean

' Returns true if string array has been allocated

   ArrIsAllS = Not Not sArr()

End Function

Public Function ArrIsAllV(vArr() As Variant) As Boolean

' Returns true if variant array has been allocated

   ArrIsAllV = Not Not vArr()

End Function

Public Sub ArrAllocD(dArr() As Date, _
                     Optional lNRow As Long = 0, _
                     Optional lNCol As Long = 0)

' Allocates or re-allocates array if necessary
' Allocation is necessary if unallocated
' Re-allocation is necessary if size does not match
' if lNRow=0 and lNCol=0 then returns array size

   Const Cl01 As Long = 1

   Dim lARow As Long, lACol As Long
   Dim bDoAll As Boolean
   
   bDoAll = Not ArrIsAllD(dArr())
   If bDoAll Then                      ' --- not allocated
      bDoAll = Not (lNRow = 0 And lNCol = 0)
   Else                                ' --- is allocated
      lARow = UBound(dArr(), 1) + Cl01 - LBound(dArr(), 1)
      lACol = UBound(dArr(), 2) + Cl01 - LBound(dArr(), 2)
      If lNRow = 0 And lNCol = 0 Then
         lNRow = lARow
         lNCol = lACol
      End If
      bDoAll = Not (lNRow = lARow And lNCol = lACol)
   End If
   If bDoAll Then ReDim dArr(1 To lNRow, 1 To lNCol)

End Sub

Public Sub ArrAllocF(fArr() As Double, _
                     Optional lNRow As Long = 0, _
                     Optional lNCol As Long = 0)

' Allocates or re-allocates array if necessary
' Allocation is necessary if unallocated
' Re-allocation is necessary if size does not match
' if lNRow=0 and lNCol=0 then returns array size

   Const Cl01 As Long = 1

   Dim lARow As Long, lACol As Long
   Dim bDoAll As Boolean

   bDoAll = Not ArrIsAllF(fArr())
   If bDoAll Then                      ' --- not allocated
      bDoAll = Not (lNRow = 0 And lNCol = 0)
   Else                                ' --- is allocated
      lARow = UBound(fArr(), 1) + Cl01 - LBound(fArr(), 1)
      lACol = UBound(fArr(), 2) + Cl01 - LBound(fArr(), 2)
      If lNRow = 0 Or lNCol = 0 Then
         lNRow = lARow
         lNCol = lACol
      End If
      bDoAll = Not (lNRow = lARow And lNCol = lACol)
   End If
   If bDoAll Then ReDim fArr(1 To lNRow, 1 To lNCol)

End Sub

Public Sub ArrAllocL(lArr() As Long, _
                     Optional lNRow As Long = 0, _
                     Optional lNCol As Long = 0)

' Allocates or re-allocates array if necessary
' Allocation is necessary if unallocated
' Re-allocation is necessary if size does not match
' if lNRow=0 and lNCol=0 then returns array size

   Const Cl01 As Long = 1

   Dim lARow As Long, lACol As Long
   Dim bDoAll As Boolean

   bDoAll = Not ArrIsAllL(lArr())
   If bDoAll Then                      ' --- not allocated
      bDoAll = Not (lNRow = 0 Or lNCol = 0)
   Else                                ' --- is allocated
      lARow = UBound(lArr(), 1) + Cl01 - LBound(lArr(), 1)
      lACol = UBound(lArr(), 2) + Cl01 - LBound(lArr(), 2)
      If lNRow = 0 Or lNCol = 0 Then
         lNRow = lARow
         lNCol = lACol
      End If
      bDoAll = Not (lNRow = lARow And lNCol = lACol)
   End If
   If bDoAll Then ReDim lArr(1 To lNRow, 1 To lNCol)

End Sub

Public Sub ArrAllocS(sArr() As String, _
                     Optional lNRow As Long = 0, _
                     Optional lNCol As Long = 0)

' Allocates or re-allocates array if necessary
' Allocation is necessary if unallocated
' Re-allocation is necessary if size does not match
' if lNRow=0 and lNCol=0 then returns array size

   Const Cl01 As Long = 1

   Dim lARow As Long, lACol As Long
   Dim bDoAll As Boolean

   bDoAll = Not ArrIsAllS(sArr())
   If bDoAll Then                      ' --- not allocated
      bDoAll = Not (lNRow = 0 Or lNCol = 0)
   Else                                ' --- is allocated
      lARow = UBound(sArr(), 1) + Cl01 - LBound(sArr(), 1)
      lACol = UBound(sArr(), 2) + Cl01 - LBound(sArr(), 2)
      If lNRow = 0 Or lNCol = 0 Then
         lNRow = lARow
         lNCol = lACol
      End If
      bDoAll = Not (lNRow = lARow And lNCol = lACol)
   End If
   If bDoAll Then ReDim sArr(1 To lNRow, 1 To lNCol)

End Sub

Public Sub ArrAllocV(vArr() As Variant, _
                     Optional lNRow As Long = 0, _
                     Optional lNCol As Long = 0)

' Allocates or re-allocates array if necessary
' Allocation is necessary if unallocated
' Re-allocation is necessary if size does not match
' if lNRow=0 and lNCol=0 then returns array size

   Const Cl01 As Long = 1

   Dim lARow As Long, lACol As Long
   Dim bDoAll As Boolean

   bDoAll = Not ArrIsAllV(vArr())
   If bDoAll Then                      ' --- not allocated
      bDoAll = Not (lNRow = 0 Or lNCol = 0)
   Else                                ' --- is allocated
      lARow = UBound(vArr(), 1) + Cl01 - LBound(vArr(), 1)
      lACol = UBound(vArr(), 2) + Cl01 - LBound(vArr(), 2)
      If lNRow = 0 Or lNCol = 0 Then   ' --- Just query size
         lNRow = lARow
         lNCol = lACol
      End If
      bDoAll = Not (lNRow = lARow And lNCol = lACol)
   End If
   If bDoAll Then ReDim vArr(1 To lNRow, 1 To lNCol)

End Sub

Public Sub ArrBoundD(dArr() As Date, _
                     Optional lR1 As Long = 0, Optional lRL As Long = 0, _
                     Optional lC1 As Long = 0, Optional lCL As Long = 0)

' Returns the bounds of the array in optional return arguments

   If ArrIsAllD(dArr()) Then
      lR1 = LBound(dArr(), 1)
      lRL = UBound(dArr(), 1)
      lC1 = LBound(dArr(), 2)
      lCL = UBound(dArr(), 2)
   End If

End Sub

Public Sub ArrBoundF(fArr() As Double, _
                     Optional lR1 As Long = 0, Optional lRL As Long = 0, _
                     Optional lC1 As Long = 0, Optional lCL As Long = 0)

' Returns the bounds of the array in optional return arguments

   If ArrIsAllF(fArr()) Then
      lR1 = LBound(fArr(), 1)
      lRL = UBound(fArr(), 1)
      lC1 = LBound(fArr(), 2)
      lCL = UBound(fArr(), 2)
   End If

End Sub

Public Sub ArrBoundL(lArr() As Long, _
                     Optional lR1 As Long = 0, Optional lRL As Long = 0, _
                     Optional lC1 As Long = 0, Optional lCL As Long = 0)

' Returns the bounds of the array in optional return arguments

   If ArrIsAllL(lArr()) Then
      lR1 = LBound(lArr(), 1)
      lRL = UBound(lArr(), 1)
      lC1 = LBound(lArr(), 2)
      lCL = UBound(lArr(), 2)
   End If

End Sub

Public Sub ArrBoundS(sArr() As String, _
                     Optional lR1 As Long = 0, Optional lRL As Long = 0, _
                     Optional lC1 As Long = 0, Optional lCL As Long = 0)

' Returns the bounds of the array in optional return arguments

   If ArrIsAllS(sArr()) Then
      lR1 = LBound(sArr(), 1)
      lRL = UBound(sArr(), 1)
      lC1 = LBound(sArr(), 2)
      lCL = UBound(sArr(), 2)
   End If

End Sub

Public Sub ArrBoundV(vArr() As Variant, _
                     Optional lR1 As Long = 0, Optional lRL As Long = 0, _
                     Optional lC1 As Long = 0, Optional lCL As Long = 0)

' Returns the bounds of the array in optional return arguments

   If ArrIsAllV(vArr()) Then
      lR1 = LBound(vArr(), 1)
      lRL = UBound(vArr(), 1)
      lC1 = LBound(vArr(), 2)
      lCL = UBound(vArr(), 2)
   End If

End Sub

Public Sub NdxGetXDS(lDst1 As Long, lDstL As Long, _
                     lSrc1 As Long, lSrcL As Long, _
                     lPut1 As Long, lPutL As Long, _
                     Optional lSsV1 As Long = 0, _
                     Optional lSsVL As Long = 0)

' An all-in-one routine for determining output bounds
' from destination bounds and output request indices
' returns valid Source bounds in lSsV1 and lSsVL

   Const Cl01 As Long = 1

   Dim lP As Long
                                       ' --- Decode meaning of output start
   If lPut1 > 0 Then lP = lPut1 + Cl01 - lDst1  ' Put1>0: from lower dest bound
   If lPut1 < 0 Then lP = lDstL + Cl01 + lPut1  ' Put1<0: from upper dest bound
   If lPut1 = 0 Then lP = lDst1                 ' Put1=0: 1st dest element
   lPut1 = lP
                                       ' --- Decode meaning of output end
   If lPutL > 0 Then lP = lPutL + Cl01 - lSrc1  ' PutL>0: from lower dest bound
   If lPutL < 0 Then lP = lSrcL + Cl01 + lPutL  ' PutL<0: from upper dest bound
   If lPutL = 0 Then lP = lDstL                 ' PutL=0: last dest element
   lPutL = lP

   lSsV1 = Cl01                        '     source start bound
   lSsVL = lPutL + lSsV1 - lPut1       '     source  end  bound
   If lPutL > lDstL Then lPutL = lDstL ' --- Clip at upper dest bound

   If lPut1 < lDst1 Then
      lSsV1 = lSrc1 + lDst1 - lPut1    '     Set source start index
      lPut1 = lDst1
   End If

   If lSsVL > lSrcL Then               ' --- Clip at upper source bound
      lPutL = lPutL + lSrcL - lSsVL    '     last dest element
      lSsVL = lSrcL                    '     last src element
   End If

End Sub

Public Sub NdxGetXSD(lSrc1 As Long, lSrcL As Long, _
                     lXtr1 As Long, lXtrL As Long, _
                     lDst1 As Long, lDstL As Long, _
                     Optional lDsU1 As Long = 0, _
                     Optional lDsUL As Long = 0)

' An all-in-one routine for determining extraction bounds from
' source bounds and extraction request indices
' returns unclipped destination bounds in lDsU1 and lDsUL

   Const Cl01 As Long = 1

   Dim lX As Long
                                       ' --- Decode meaning of extraction start
   If lXtr1 > 0 Then lX = lXtr1 + Cl01 - lSrc1
   If lXtr1 < 0 Then lX = lSrcL + Cl01 + lXtr1
   If lXtr1 = 0 Then lX = lSrc1
   lXtr1 = lX
                                       ' --- Decode meaning of extraction end
   If lXtrL > 0 Then lX = lXtrL + Cl01 - lSrc1
   If lXtrL < 0 Then lX = lSrcL + Cl01 + lXtrL
   If lXtrL = 0 Then lX = lSrcL
   lXtrL = lX

   lDsU1 = Cl01                        '     Unclipped dest start bound
   lDsUL = lXtrL + lDsU1 - lXtr1       '     Unclipped dest  end  bound
   If lXtrL > lSrcL Then lXtrL = lSrcL ' --- Clip upper bound
   lDstL = lXtrL + Cl01 - lXtr1        '     Set destination end index

   lDst1 = Cl01                        ' --- Clip lower bound
   If lXtr1 < lSrc1 Then
      lDst1 = lDst1 + lSrc1 - lXtr1
      lXtr1 = lSrc1
   End If

End Sub

Public Function NdxGetSX1(lSrc1 As Long, lSrcL As Long, lXtr1 As Long) As Long

' Returns the extraction start source position
' lSrc1 and lSrcL are the source upper and lower bounds
' lXtr1           specifies the first source position to extract
'                 >0: First position from the left  ( 1: leftmost )
'                 <0: First position from the right (-1: rightmost)
'                 =0: First position
' If lXtr1 > lSrcL, is returned as such (nothing extracted)
' if lXtr1 < lSrc1, is returned as such (clipped extraction)

   Const Cl01 As Long = 1

   Dim lSX1 As Long

   If lXtr1 > 0 Then
      lSX1 = lXtr1 + Cl01 - lSrc1
   End If
   If lXtr1 < 0 Then
      lSX1 = lSrcL + Cl01 + lXtr1
   End If
   If lXtr1 = 0 Then
      lSX1 = lSrc1
   End If

   NdxGetSX1 = lSX1

End Function

Public Function NdxGetSXL(lSrc1 As Long, lSrcL As Long, lXtrL As Long) As Long

' Returns the extraction end source position
' lSrc1 and lSrcL are the source upper and lower bounds
' lXtrL           specifies the last source position to extract
'                 >0: Last position from the left ( 1: leftmost)
'                 <0: Last position from the right (-1: rightmost)
'                 =0: Last position
' If lXtrL > lSrcL, is returned as such (clipped extraction)
' if lXtrL < lSrc1, is returned as such (nothing extracted)

   Const Cl01 As Long = 1

   Dim lSXL As Long

   If lXtrL > 0 Then
      lSXL = lXtrL + Cl01 - lSrc1
   End If
   If lXtrL < 0 Then
      lSXL = lSrcL + Cl01 + lXtrL
   End If
   If lXtrL = 0 Then
      lSXL = lSrcL
   End If

   NdxGetSXL = lSXL

End Function

Public Function NdxLenSXL(lXtr1 As Long, lXtrL As Long, lXtrN As Long) As Long

' Returns the extraction end source position from
' lXtr1: extraction start source position
' lXtrL: extraction end source position
' lXtrN: extraction length
' lXtr1 must have been converted to effective index by NdxGetSX1 before
' if lXtrN = 0, lXtrL is used as such
' if lXtrN > 0, lXtrL is determined by lXtr1 and lXtrN

   Const Cl01 As Long = 1

   Dim lSXL As Long

   If lXtrN > 0 Then
      lSXL = lXtr1 + lXtrN - Cl01
   Else
      lSXL = lXtrL
   End If

   NdxLenSXL = lSXL

End Function

Public Function NdxClpSX1(lSrc1 As Long, lXtr1 As Long) As Long

' Returns the clipped destination start index when the source start position
' is out of source  bounds and corrects lXtr1 to be within the source bounds:
' When lXtr1 < lSrc1, returns lSrc1+1-lXtr1 and sets lXtr1 = lSrc1

   Const Cl01 As Long = 1

   Dim lDst1 As Long

   lDst1 = Cl01
   If lXtr1 < lSrc1 Then
      lDst1 = lDst1 + lSrc1 - lXtr1
      lXtr1 = lSrc1
   End If

   NdxClpSX1 = lDst1

End Function

Public Function NdxClpSXL(lSrcL As Long, lXtr1 As Long, lXtrL As Long) As Long

' Returns the clipped destination end index when the source end position
' is out of source bounds and corrects lXtrL to be within the source bounds:
' When lXtrL>lSrcL, returns .... and sets lXtrL=lSrcL

   Const Cl01 As Long = 1

   Dim lDstL As Long

   If lXtrL > lSrcL Then               ' --- Extract past upper bound
      lXtrL = lSrcL                    '     Clip source end to bound
   End If                              ' --- Set destination end index
   lDstL = lXtrL + Cl01 - lXtr1

   NdxClpSXL = lDstL

End Function

