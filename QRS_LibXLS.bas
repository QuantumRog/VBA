Attribute VB_Name = "QRS_LibXLS"
Option Explicit

' Module : QRS_LibXLS
' Project: any
' Purpose: Utility functions with Excel objects
'          Workbooks, Worksheets, Ranges
'          Connections, ListObjects, QueryTables
'          DB Connections: Importing data from a database table
'          instantiates the follwing objects:
'           1) A Workbook.Connection of Type xlConnectionTypeOLEDB
'           2) If the Workbook.Connection connection type is OLEDB
'              the OLEDBConnection object contains the following:
'              a) The OLEDB connection text
'              b) The command passed
'              c) The command type
'           3) The Workbook.Connection object contains a handle to
'              the range object representing the data output range
'              which allow to identify the workbook connection from
'              a cell within the range
'           4) The range object contains a handle to the ListObject
'              which contains the ListColumn collection representing
'              the table fields
'           5) The parent object of the range is the output worksheet
'
' By     : QRS, Roger Strebel
' Date   : 28.06.2018                  IsOpenWB, ExistsWB
'          29.07.2018                  Smart OLEDB connection reuse
'          02.10.2018                  RangeBounds improved
'          28.01.2019                  VarToBoolean added            28.01.2019
'          06.02.2019                  OLECnxStr improved            06.02.2019
' --- The public interface
'     ExistsCnxWkB                     Exists workbook connection?
'     ExistsWB                         Exists workbook in folder?    28.06.2018
'     ExistsWS                         Exists worksheet in book?     03.07.2018
'     IsCellInRng                      Is specified cell in range?   28.06.2018
'     IsOpenWB                         Is specified workbook open?
'     LstObjColLst                     List object columns list      29.07.2018
'     LstObjOrderN                     List object multi-sort        02.10.2018
'     LstObjGetRow                     List object row values
'     OLECnxStr                        Get OLEDB connection string   06.02.2019
'     RangeBounds                      Row and Column range bounds   02.10.2018
'     RngCol_Width                     Range column width (twips)
'     RngLocOffRng                     Range location within range   31.07.2018
'     SetRefXLS                        Set Excel reference objects   03.07.2018
'     TxtLsOSrcTyp                     ListObject SourceType text    28.06.2018
'     TxtOLECmdTyp                     OLEDB command type text       28.06.2018
'     TxtOLERobust                     OLEDB Robust connection mode  29.07.2018
'     TxtOLESvrCre                     OLEDB server credent. method  29.07.2018
'     TxtWbkCnxTbl                     Workbook connection table     03.07.2108
'     TxtWbkCnxTyp                     Workbook connection type      03.07.2018
'     VarToBoolean                     Variant to Boolean            28.01.2019

Public Sub SetRefXLS(sPath As String, sWb As String, _
                     sWs As String, sRng As String, _
                     aWb As Workbook, aWs As Worksheet, aRng As Range)

' Returns the objects referred by names
' Some rules for simplifications
' - If the workbook name is empty, path is ignored
' - The workbook name may contain abreviations
'     "." means "ThisWorkbook"
'     "@" means "ActiveWorkbook"
'     ""  means no reassignment
' - If the worksheet name is not specified, the
'   worksheet object is not reassigned. If the
'   worksheet object hast not been assigned before,
'   the range is ignored
' - If the range begins with "#", it is considered a name
'   In this case, the names lists are browsed
'   first for the worksheet, then for the workbook, the for the application

   Dim b As Boolean

   If sWb = "." Then
      b = True
      Set aWb = ThisWorkbook
   End If
   If sWb = "@" Then
      b = True
      Set aWb = Application.ActiveWorkbook
   End If
   If Not (b Or sWb = "") Then
      b = IsOpenWB(sWb, aWb)
      If Not b Then b = ExistsWB(sPath, sWb, aWb)
   End If

   If Not b Then Exit Sub
   b = False

   If Not sWs = "" Then
      b = ExistsWS(aWb, sWs, aWs)
   End If

   If Not b Then Exit Sub

   Set aRng = aWs.Range(sRng)

End Sub

Sub testado()

   Dim sX As String, sT As String, sf() As String
   Dim v() As Variant
   Dim b As Boolean

   sT = "t_Test1"
   sX = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=QRS_Playground;Data Source=ROGSHP;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=ROGSHP;Use Encryption for Data=False;Tag with column collation when possible=False"

   QRS_LibArr.ArrAllocV v(), 2, 11
   b = QRS_LibArr.ZipEleVarRow(v(), 1, 1, "AB", "CD", -12, -34, 3.1, 0.014)
   b = QRS_LibArr.ZipEleVarRow(v(), 1, 7, False, Date, Date - 2, "Kommentar")
   
   b = QRS_LibArr.ZipEleVarRow(v(), 2, 1, "ZY", "XW", 12, 34, 2.7, 0.018)
   b = QRS_LibArr.ZipEleVarRow(v(), 2, 7, True, Date + 2, Date, "Zweite")

'   b = QRS_LibADO.TblUpdArr(v(), sf(), sX, sT)

End Sub

Sub TestCnx()

   Dim sW As String, sH As String
   Dim aW As Workbook, aH As Worksheet, aC As Range, aD As Range
   Dim aL As ListObject
   Dim aX As WorkbookConnection
'   Dim oRS As ADODB.Recordset
   Dim lR1 As Long, lRL As Long, lC1 As Long, lCL As Long
   Dim b As Boolean

' ---------------------------------------- Workbookconnection object

   SetRefXLS "", ".", "Test", "C7", aW, aH, aC

   b = ExistsCnxWkB(aC, sH, aX)
   If Not aX.Type = xlConnectionTypeOLEDB Then Exit Sub

   sW = TxtWbkCnxTbl(aX.OLEDBConnection.CommandText)

'   Set oRS = New ADODB.Recordset
'   oRS.Open Source:=sW, ActiveConnection:=aX.OLEDBConnection, CursorType:=adOpenDynamic, LockType:=adLockBatchOptimistic

   With aL.ListColumns
      For lR1 = 1 To .Count
Debug.Print "Name  = " & .Item(lR1).Name
Debug.Print "Index = " & .Item(lR1).Index
      Next lR1
   End With


Stop
   If aW.Connections.Count = 0 Then
Debug.Print "Workbook " & aW.Name & " contains no connections"
      Exit Sub
   End If
   lCL = 1
   With ThisWorkbook.Connections(lCL)
Debug.Print "Connection " & lCL & " info:"
Debug.Print "   Name: " & .Name
Debug.Print "   Descr: " & .Description
Debug.Print "   Type " & TxtWbkCnxTyp(.Type)
      If .Type = xlConnectionTypeOLEDB Then
         With .OLEDBConnection
Debug.Print "   OLEDB connection info: "
Debug.Print "      Command: " & .CommandText
Debug.Print "      CmdType: " & TxtOLECmdTyp(.CommandType)
Debug.Print "      Connect: " & .Connection
         End With
Debug.Print "   Range count: " & .Ranges.Count
         If .Ranges.Count > 0 Then
            lC1 = 1
            With .Ranges(lC1)
Debug.Print "      Range " & lC1 & " address: " & .Parent.Name & "!" & .Address
               Set aL = .ListObject
            End With
Debug.Print "   Listobject name: " & aL.Name
Debug.Print "      Comment: " & aL.Comment
Debug.Print "      Source type: " & TxtLsOSrcTyp(aL.SourceType)
            If Not aL.HeaderRowRange Is Nothing Then
Debug.Print "      Header row range: " & aL.HeaderRowRange.Address
            End If
            If Not aL.InsertRowRange Is Nothing Then
Debug.Print "      Insert row range: " & aL.InsertRowRange.Address
            End If
Debug.Print "      List columns count "; aL.ListColumns.Count
Debug.Print "      List column 1: " & aL.ListColumns(1).Name
Debug.Print "      Query table info:"
Debug.Print "         Edit enabled? " & aL.QueryTable.EnableEditing
Stop
Debug.Print "         Parameter count: " & aL.QueryTable.Parameters.Count
Debug.Print "   " & aL.QueryTable.Connection
Debug.Print "   " & .OLEDBConnection.Connection
         End If
      End If
   End With

End Sub

Public Function ExistsCnxWkB(aCell As Range, sCnxName As String, _
                             aCnxWkB As WorkbookConnection) As Boolean

' Returns true if the specified workbook connection exists
' if sCnxName="", uses aCell to determine the workbook, sheet and range
'                 to see whether the cell is contained in the output range
' else,           Checks the parent workbook of aCell for
'                 connections with the specified name

   Dim aWb As Workbook, aWs As Worksheet, aRg As Range
   Dim bNo As Boolean

   If aCell Is Nothing Then Exit Function

   Set aWb = aCell.Parent.Parent

   If sCnxName = "" Then               ' --- Connection search by
      For Each aCnxWkB In aWb.Connections  ' output range
         With aCnxWkB
            For Each aRg In .Ranges
               If IsCellInRng(aCell, aRg) Then Exit For
            Next aRg
         End With
         If Not aRg Is Nothing Then Exit For
      Next aCnxWkB
      bNo = aCnxWkB Is Nothing
      If Not bNo Then
         sCnxName = aCnxWkB.Name
      End If
   Else
      For Each aCnxWkB In aWb.Connections
         If aCnxWkB.Name = sCnxName Then Exit For
      Next aCnxWkB
      bNo = aCnxWkB Is Nothing
      If Not bNo Then                  ' --- connection identified
         With aCnxWkB
            If .Ranges.Count > 0 Then Set aCell = aCnxWkB.Ranges(1)
         End With
      End If
   End If

   ExistsCnxWkB = Not bNo

End Function

Public Function ExistsWB(sPath As String, sWb As String, _
                         Optional aWb As Workbook, _
                         Optional bReadOnly As Boolean = False, _
                         Optional bUpdLinks As Boolean = False) As Boolean

' Returns true if the specified workbook in the specified path exists

   Dim sApp As String

   sApp = QRS_Lib0.PathApp(sPath, sWb)
   If QRS_LibDOS.FileExists(sApp) Then
      ExistsWB = True
      Set aWb = Application.Workbooks.Open(sApp, bUpdLinks, bReadOnly)
   End If

End Function

Public Function ExistsWS(aWb As Workbook, sWs As String, _
                         aWs As Worksheet) As Boolean

' Returns true if the worksheet specified by sWs exists in the
' workbok aWb and also returns the sheet in aWs

   Dim lLen As Long
   
   lLen = Len(sWs)
   If lLen = 0 Then Exit Function

   For Each aWs In aWb.Worksheets
      If StrComp(Left(aWs.Name, lLen), sWs, vbTextCompare) = 0 Then Exit For
   Next aWs

   ExistsWS = Not aWs Is Nothing

End Function

Public Function IsOpenWB(sWb As String, _
                         Optional aWb As Workbook) As Boolean

' Returns true if the specified workbook is open
' Returns workbook reference in the optional aWb output argument

   Dim lLen As Long

   lLen = Len(sWb)
   If lLen = 0 Then Exit Function

   For Each aWb In Application.Workbooks
      If StrComp(Left(aWb.Name, lLen), sWb, vbTextCompare) = 0 Then Exit For
   Next aWb

   IsOpenWB = Not aWb Is Nothing

End Function

Public Function IsCellInRng(aCell As Range, aRng As Range) As Boolean

' Returns true if aCell is contained in aRng

   Dim lRngRow1 As Long, lRngRowL As Long
   Dim lRngCol1 As Long, lRngColL As Long
   Dim lCellRow As Long, lCellCol As Long

   RangeBounds aRng, lRngRow1, lRngRowL, lRngCol1, lRngColL
   RangeBounds aCell, lCellRow, , lCellCol
   IsCellInRng = QRS_Lib0.WithinL(lCellRow, lRngRow1, lRngRowL) And _
                 QRS_Lib0.WithinL(lCellCol, lRngCol1, lRngColL)

End Function

Public Function OLECnxStr(aWkbCnx As WorkbookConnection) As String

' Returns the OLEDB connection string if the WorkbookConnection provided
' is of type OLEDB
' Returns an empty string if th OLEDB connection is not initialized

   Dim sOLE As String, sCnx As String

   If aWkbCnx.Type = xlConnectionTypeOLEDB Then
      On Error Resume Next
      sOLE = aWkbCnx.OLEDBConnection.Connection
      On Error GoTo 0
      QRS_LibStr.StrSplit2 sOLE, ";", , sCnx
   End If

   OLECnxStr = sCnx

End Function

Public Function LstObjColLst(oLstObj As ListObject, sFld() As String) As Long

' Returns the columns list of the list object
' Goes through the ListColumns collection and
' is independent of the header row show/hide.

   Dim oLstCol As ListColumn
   Dim lICol As Long

   If oLstObj Is Nothing Then
      lICol = -1
   Else
      ReDim sFld(1 To oLstObj.ListColumns.Count)
   
      For Each oLstCol In oLstObj.ListColumns
         lICol = lICol + 1
         sFld(lICol) = oLstCol.Name
      Next oLstCol
   End If

   LstObjColLst = lICol

End Function

Public Sub LstObjOrderN(oLstObj As ListObject, sFldLst As String, _
                        Optional sLstSep As String = ",")

' Applies multi-level ordering on list-object columns
' The list-object is supposed to be linked to external data
' and to have a header row in which the fields are identified
' The sort level oder is determined by the field sequence
' The first field is the top level

   Dim aCol As Range
   Dim v()
   Dim sColHdr() As String, sColFld() As String
   Dim lIFld As Long, lNFld As Long, lXCol As Long
                                       ' --- Field string to list
   QRS_LibStr.LstStrLst sFldLst, sColFld(), "", "", ","
   lNFld = UBound(sColFld())

   With oLstObj
      v() = .HeaderRowRange.Value
      QRS_LibA2L.GetRowStrVar v(), sColHdr()
      .Sort.SortFields.Clear           ' --- Remove previous selection
      For lIFld = 1 To lNFld
         lXCol = QRS_LibLst.LstFind_S(sColHdr(), sColFld(lIFld))
         If lXCol > 0 Then
            Set aCol = .Range.Columns(lXCol)
            .Sort.SortFields.Add aCol, xlSortOnValues, xlAscending
         End If
      Next lIFld
      .Sort.Apply
   End With


End Sub

Public Sub LstObjGetRow(oLstObj As ListObject, lDataRngRow As Long, vLstRow())

' Returns the list of value in DataBodyRange row lDataRngRow in vLstRow

   Dim lICol As Long, lNCol As Long
   Dim v()

   With oLstObj.DataBodyRange
      lNCol = .Columns.Count
      QRS_LibLst.LstAllocV vLstRow(), lNCol
      If lNCol = 1 Then
         vLstRow(1) = .Cells(lDataRngRow, 1).Value
      Else
         v() = .Rows(lDataRngRow).Value
         For lICol = 1 To lNCol        ' --- Flatten to list
            vLstRow(lICol) = v(1, lICol)
         Next lICol
      End If
   End With

End Sub

Public Sub RangeBounds(aRng As Range, _
                       Optional lRow1 As Long = 0, _
                       Optional lRowL As Long = 0, _
                       Optional lCol1 As Long = 0, _
                       Optional lColL As Long = 0)

' Returns the range bounds in row and column numbers

   If aRng Is Nothing Then
      lRow1 = 0: lRowL = 0: lCol1 = 0: lColL = 0
   Else
      With aRng
         lRow1 = .Row
         lCol1 = .Column
         lRowL = .Rows(.Rows.Count).Row
         lColL = .Columns(.Columns.Count).Column
      End With
   End If

End Sub

Public Function RngCol_Width(aRngIn As Range) As Double

' Returns the width of the column of arngIn (top left cell)
' in Twips (resolution-independent Excel width figure)

   RngCol_Width = aRngIn.ColumnWidth

End Function

Public Function RngLocOffRng(aRngIn As Range, aRngOf As Range, _
                             Optional bTopLeftOnly As Boolean = True, _
                             Optional lRowOf As Long = 0, _
                             Optional lColOf As Long = 0) As Boolean

' Reurns the offset of aRngIn within aRngOf
' if bTopLeftOnly is true, the function returns true if no part of aRngIn
' is out of the bounds of aRngOf
' if bTopLeftOnly is false, the function checks only for the top left cell
' of aRngIn

   Const Cl01 As Long = 1

   Dim lRI1 As Long, lRIL As Long, lCI1 As Long, lCIL As Long
   Dim lRO1 As Long, lROL As Long, lCO1 As Long, lCol As Long
   Dim bOut As Boolean

   RangeBounds aRngIn, lRI1, lRIL, lCI1, lCIL
   RangeBounds aRngOf, lRO1, lROL, lCO1, lCol
   lRowOf = lRI1 + Cl01 - lRO1
   lColOf = lCI1 + Cl01 - lCO1
   If bTopLeftOnly Then
      lCIL = lCI1
      lRI1 = lRIL
   End If

   bOut = lCI1 < lCO1 Or lRI1 < lRO1 Or lCIL > lCol Or lRIL > lROL

   RngLocOffRng = bOut

End Function

Public Function TxtWbkCnxTbl(sCnx As String) As String

' WorkbookConnection string table name (section after last dot)

   Dim sX As String, s() As String

   QRS_LibStr.LstStrLst sCnx, s(), "."

   sX = s(UBound(s()))
   TxtWbkCnxTbl = QRS_LibStr.StrRpl(sX, Chr(34), "")

End Function

Public Function TxtWbkCnxTyp(eCnxTyp As XlConnectionType)

' Workbook Connection type enumeration

   Dim s As String

   Select Case eCnxTyp
   Case XlConnectionType.xlConnectionTypeODBC
      s = "ODBC"
   Case XlConnectionType.xlConnectionTypeOLEDB
      s = "OLEDB"
   Case XlConnectionType.xlConnectionTypeTEXT
      s = "TEXT"
   Case XlConnectionType.xlConnectionTypeWEB
      s = "WEB"
   Case XlConnectionType.xlConnectionTypeXMLMAP
      s = "XML"
   End Select

   TxtWbkCnxTyp = s

End Function

Public Function TxtLsOSrcTyp(eSrcTyp As XlSourceType) As String

' Returns the ListObject XlSourceType enumeration value as clear text

   Dim s As String

   Select Case eSrcTyp
   Case xlSourceAutoFilter
      s = "autofilter"
   Case xlSourceChart
      s = "chart"
   Case xlSourcePivotTable
      s = "pivot table"
   Case xlSourcePrintArea
      s = "print area"
   Case xlSourceQuery
      s = "query"
   Case xlSourceRange
      s = "range"
   Case xlSourceSheet
      s = "worksheet"
   Case xlSourceWorkbook
      s = "workbook"
   End Select

   TxtLsOSrcTyp = s

End Function

Public Function TxtOLECmdTyp(eCmdTyp As XlCmdType) As String

' OLEDB Connection command type enumeration

   Dim s As String

   Select Case eCmdTyp
   Case xlCmdCube
      s = "cube"
   Case xlCmdDefault
      s = "default"
   Case xlCmdList
      s = "list"
   Case xlCmdSql
      s = "SQL"
   Case xlCmdTable
      s = "table"
   End Select

   TxtOLECmdTyp = s

End Function

Public Function TxtOLERobust(eCnxRobust As XlRobustConnect) As String

' OLEDB Robust connection mode

   Dim s As String

   Select Case eCnxRobust
   Case xlAsRequired
      s = "As required"
   Case xlAlways
      s = "Always"
   Case xlNever
      s = "Never"
   End Select

   TxtOLERobust = s

End Function

Public Function TxtOLESvrCre(eSvrCre As XlCredentialsMethod) As String

' Returns the OLEDB connection server credentials method clear text

   Dim s As String

   Select Case eSvrCre
   Case xlCredentialsMethodIntegrated
      s = "Integrated"
   Case xlCredentialsMethodNone
      s = "None"
   Case xlCredentialsMethodStored
      s = "Stored"
   End Select

   TxtOLESvrCre = s

End Function

Public Function VarToBoolean(v As Variant) As Boolean

   Dim b As Boolean

   VarToBoolean = v

End Function
