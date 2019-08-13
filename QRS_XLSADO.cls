VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QRS_XLSADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Module :
' Project:
' Purpose: A class for simple update of database data in tables
'          upon edition in the Excel output range
' concept: Specify connection, database table
'          Specify ID field, Field list to display
'          Specify condition for data retrieving
'          The retrieve query is executed and outputs data
'          When a cell in the output range is left (SelectionChange)
'          the field value is edited
'          When the row is left, the record is updated
'          When the row insert input range is left, the row is appended
'          Connection is closed upon class termination
'          Connection data are perennalized on a range specified
' use A  : Specification of a table, field list for output to Excel ranges
'           1. instantiate this class
'           2. specify Excel output ranges using either
'              SetXLSRef  with reference strings or
'              SetXLSRng  with output range objects
'           3. specify the ADO to use (opened before)
'              using SetCnxADO
'           4. Query and output the table using QryTblXls
'              specifying the table name and field list
' Date   : 07.02.2019                     All members for use A work
' Refs   : ADO 2.8
'          QRS_LibADO
' --- The public interface
' Properties
'     QryTblXLS                           Query table and output to ranges
' Methods:
'     SetCnxADO                           Specify database connection
'     SetXLSRef                           Specify output ranges by ref strings
'     SetXLSRng                           Specify output ranges as ranges

Private MbShowHdrs As Boolean
Private MbShowData As Boolean

Private MaRngColHdr As Range              ' --- Column header (field names) range
Private MaRngTbData As Range              ' --- Table Data output range
Private MaRngRowIns As Range              ' --- Row insert input range

Private MoCnx As ADODB.Connection         ' --- Foreign ADO connection
Private MoRcs As ADODB.Recordset          ' --- Internal ADO recordset

Private MaLstObj As ListObject

Public Sub SetXLSRef(sRefHdr As String, sRefTbData As String, _
                     Optional bShowHdr As Boolean = False, _
                     Optional bShowDat As Boolean = True)

   Dim aRefXL As QRS_LibXL.tRefXL
   Dim bFail As Boolean

   aRefXL.sRef = sRefHdr
   bFail = QRS_LibXL.RefXLSetObj(aRefXL)
   If Not bFail Then Set MaRngColHdr = aRefXL.aCl

   aRefXL.sRef = sRefTbData
   bFail = QRS_LibXL.RefXLSetObj(aRefXL)
   If Not bFail Then Set MaRngTbData = aRefXL.aCl

   MbShowHdrs = bShowHdr
   MbShowData = bShowDat

End Sub

Public Sub SetXLSRng(aRngHdr As Range, aRngTbData As Range, _
                     Optional bShowHdr As Boolean = False, _
                     Optional bShowDat As Boolean = True)

   Set MaRngColHdr = aRngHdr
   Set MaRngTbData = aRngTbData

   MbShowHdrs = bShowHdr
   MbShowData = bShowDat

End Sub

Public Sub SetCnxADO(oCnx As ADODB.Connection)

   Set MoCnx = oCnx

End Sub

Public Function QryTblXLS(sTbl As String, sFld As String) As Boolean

' A method to query fields from a table and putting the values
' out to the specified excel range, with or without headers

   Dim s() As String
   Dim bFail As Boolean

   If MbShowHdrs Then                  ' --- Headers from sFld
      QRS_LibStr.ArrStrArr "[" & sFld, s(), , ","
      QRS_LibXL.RngPutArrS MaRngColHdr, s()
   End If

   QRS_LibADO.TblGetRcS MoCnx.ConnectionString, sTbl, MoCnx, MoRcs
   bFail = Not MoRcs.State = adStateOpen
   If bFail Then GoTo QryTblXLS_Ende
   bFail = MoRcs.BOF Or MoRcs.EOF
   If bFail Then GoTo QryTblXLS_Ende

   If MbShowData Then

      Dim vF(), vV(), vX()

      If bFail Then Exit Function
      QRS_LibADO.RcSFldStrLst sFld, vF()
      QRS_LibADO.RcSGetCol MoRcs, vV(), vF(), bFail
'      QRS_LibADO.VarArrX01 vV(), vX()
      QRS_LibXL.RngPutArrV MaRngTbData, vV()

   End If

QryTblXLS_Ende:

   QryTblXLS = bFail

End Function

Private Sub Class_Terminate()

   QRS_LibADO.ByeRcS MoRcs
   QRS_LibADO.ByeCnx MoCnx

End Sub
