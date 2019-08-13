VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QRS_TblSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Module : QRS_TblSel
' Purpose: Maintain a selection drop-down list
'          for a foreign-key field
'          A view may return a table with joined
'          base table values by foreign key of the
'          main table. This class manages a list of
'          the joined values and may return the
'          base table ID to be used as the foreign key
'          Some base tables may also use other keys
'          for filtering
' Use    : Specify the base table
'          Specify the list field
'          Specify the identity field
'          Specify up to three other keys
'          Specify the cell linked to the list
'
' --- public interface
' Properties
'     ValIsUnique                      Check if all values are unique
' Methods
'     TblQry
'     TblSet


Private MsTblBase As String            ' --- Base table name
Private MsFldShow As String            ' --- Field shown in join output
Private MsFldFK As String              ' --- Foreign key field in join
Private MsFldID As String              ' --- Base table ID field
Private MsFldSel As String             ' --- Base table select field

Public McQLU As QRS_LookUp             ' --- The quick look-up table

Private MoRS As ADODB.Recordset
Private MsLU() As String                ' --- local string array

Public Sub TblQry(oCnx As ADODB.Connection)

   TblInit oCnx

End Sub

Public Sub TblSet(sTblBase As String, sFldShow As String, _
                  sFldFK As String, sFldBaseID As String, _
                  sFldSel As String)

   MsTblBase = sTblBase                ' --- base table for query
   MsFldShow = sFldShow                ' --- join show field for list display
   MsFldFK = sFldFK                    ' --- Foreign key join field for ref
   MsFldID = sFldBaseID                ' --- Base id field for join
   MsFldSel = sFldSel                  ' --- Base field for selection list

End Sub

Private Sub TblInit(oCnx As ADODB.Connection)

' --- Initialize the Base table data from the database
'     It is assumed that transferring the entire table
'     and extracting the two columns locally is faster
'     and safer than executing a string command which
'     would return only the two columns

   Dim vV(), vF()
   Dim sCnx As String
   Dim bNoData As Boolean

   If Not IsSetQLU_OK Then Exit Sub
   If oCnx Is Nothing Then Exit Sub

   QRS_LibLst.LstAllocV vF(), 2        ' --- Col 1: ID, Col 2: Selection
   vF(1) = MsFldID
   vF(2) = MsFldSel

   sCnx = oCnx.ConnectionString        ' --- Reuse connection
   QRS_LibADO.TblGetRcS sCnx, MsTblBase, oCnx, MoRS
   QRS_LibADO.RcSGetCol MoRS, vV(), vF(), bNoData
   If bNoData Then Exit Sub

   QRS_LibArr.GetSubVarStr vV(), MsLU()    ' Copy to string

'   McQLU.SetTbl sV(), 1, 2             ' --- Set up list

End Sub

Private Function IsSetQLU_OK() As Boolean

' Returns true if the settings are OK to obtain the look-up table
' Base table name, ID field and field for selection list

   IsSetQLU_OK = Not (MsTblBase = "" Or MsFldID = "" Or MsFldSel = "")

End Function
