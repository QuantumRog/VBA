Attribute VB_Name = "QRS_LibADO"
Option Explicit

' Module : LibADO
' Project: Any project involving ADO database operations
' Purpose: Some intermediate-level data processing utilites
'          Application intention
'            1) Optain a recordset from a command
'               with optional parameters
'            2) Execute a command with optional parameters
'            3) Insert values to a table
' By     : QRS, Roger Strebel
' Date   : 24.06.2018                  Carmen Spycher's Birthday
'          04.07.2018                  Really useful routines OK
'          07.07.2018                  RcSFldLs4Upd hardened
'          08.07.2018                  Minor improvements on TblUpdArr
'          23.07.2018                  RcSGetRow added
'          07.02.2019                  RcSGetCol, VarArr/LstX01 added
' --- The public interface
'     ByeCmd                           Command smooth release        04.07.2018
'     ByeCnx                           Connection smooth release     04.07.2018
'     ByeRcS                           Recordset smooth release      04.07.2018
'     DBCmdExec                        Execute command with return   04.07.2018
'     DBCmdIni                         Initialize command object     04.07.2018
'     DBCmdSetPrm                      Set parameters and values     04.07.2018
'     DBConnect                                                      04.07.2018
'     RcsFldLs4Upd                     Recordset update fields list  04.07.2018
'     RcsFldLstStr                     Recordset fields list         04.07.2018
'     RcSFldNdx_ID                     Recordset identifier field
'     RcSFldStrLst                     Recordset fields from string
'     RcsGetCol                        Recordset column extraction   07.02.2019
'     RcsGetRow                        Recordset values of one row   07.02.2019
'     TblGetRcS                        Recordset from one table      04.07.2018
'     TblRow4RSUpd                     Recordset update row values   07.07.2018
'     TblUpdArr                        Update table from array       08.07.2018
'     TxtBmkMode                       Recordset bookmark modes      24.06.2018
'     TxtCmdTyp                        Command types                 24.06.2018
'     TxtCnxMode                       Connection modes              24.06.2018
'     TxtCnxOpt                        Connection options            24.06.2018
'     TxtCurLoc                        Recordset cursor location     24.06.2018
'     TxtCurOpt                        Recordset cursor options      24.06.2018
'     TxtCurTyp                        Recordset cursor type         24.06.2018
'     TxtDatTyp                        Field Data types              24.06.2018
'     TxtEdiMode                       Recordset edit modes          24.06.2018
'     TxtFldPrp                        Field properties debug
'     TxtLckTyp                        Lock types                    24.06.2018
'     TxtPrmAtt                        Command parameter attributes  24.06.2018
'     TxtPrmDir                        Command parameter directions  24.06.2018
'     TstRcsPrp                        Recordset properties debug
'     TxtRecStat                       Record status                 24.06.2018
'     TxtRecTyp                        Record type                   24.06.2018
'     TxtXecOpt                        Command execution modes       24.06.2018
'     VarArrX01                        Transpose-shift var array     07.02.2019
'     VarLstX01                        Shift-expand field list       07.02.2019
' --- The private sphere

Private Const MCsSep As String = ","

Public Sub DBCmdExec(sCmd As String, _
                     oCnx As ADODB.Connection, oCmd As ADODB.Command, _
                     oRcS As ADODB.Recordset, _
                     eTyp As ADODB.CommandTypeEnum, _
                     Optional sPrm As String = "", _
                     Optional lRet As Long = 0)

' Passes SQL to the database and returns records from dataase
' SQL may either be unspecified, unknown
' an    existing       table name (eTyp = adCmdTable)
' or an existing stored procedure (eTyp = adCmdStoredProc)
' or a  valid       SQL statement (eTyp = adCmdText)
' For a call to a stored procedure, the parameter names and values
' are taken from the optional sPrm (format see CmdSetPrm) argument
' If the stored procedure returns a return code, it is contained in
' the first "@RETURN_VALUE". If present, its value is returned in lRet

   Const sRet As String = "@RETURN_VALUE"

   If oCnx Is Nothing Then Exit Sub

   DBCmdIni oCnx, oCmd                 ' --- Initialize command object

   oCmd.CommandText = sCmd             ' --- Set command text
   oCmd.CommandType = eTyp             ' --- Set command type
   If eTyp = adCmdStoredProc Then DBCmdSetPrm sPrm, oCmd
   Set oRcS = oCmd.Execute()           ' --- Set parameter values

   If eTyp = adCmdStoredProc Then      ' --- Extract return value
      If Not IsEmpty(oCmd.Parameters(sRet).Value) Then
         lRet = oCmd.Parameters(sRet).Value
      End If
   End If

' Idee hier: weitere Output-Parameter zurücklesen:
'            Parameterliste aus String splitten,
'            Schauen, ob Parameter auch Output sind (AND adParamOutput)
'            falls ja, String zusammensetzen @Parm=Rückgabewert

End Sub

Public Sub DBCmdSetPrm(sPrm As String, oCmd As ADODB.Command, _
                       Optional sDel = ",")

' Sets the parameter values of the command object
' It looks as if the parameter collection was set to contain all parameters
' of the stored procedure specified by the oCmd.CommmandText property value
' This routine sets the values for the named input parameters in sPrm
' Paremeter names in sPrm may or may not contain the At-Sign. If they
' don't the At-sign is prepended.
'    PrmName1=PrmValue1,PrmName2=PrmValue2

   Const sAt As String = "@"

   Dim lE1 As Long, lEL As Long, lEI As Long, lL As Long
   Dim sP() As String, sV() As String, sN As String
                                       ' --- Split string to list
   sP() = Split(sPrm, sDel, , vbTextCompare)
   QRS_LibLst.LstBoundS sP(), lE1, lEL
   For lEI = lE1 To lEL                ' --- All parameters
      sV() = Split(sP(lEI), "=", , vbTextCompare)
      lL = LBound(sV())
      sN = sV(lL)                      ' --- Prepend At if missing
      If Not Left(sN, Len(sAt)) = sAt Then sN = sAt & sN
      If UBound(sV()) > lL Then oCmd.Parameters(sN).Value = sV(lL + 1)
   Next lEI

End Sub

Public Sub DBConnect(sCnx As String, sUsr As String, sPwd As String, _
                     oCnx As ADODB.Connection)

' Establishes a connection to a database using an ADO connection string
' The connection is then available through the ADODB.oCnx object
' Also instantiates a command object for passing database commands
' of specific type (table, stored procedures, sql command text)

   Dim bNew As Boolean

   If oCnx Is Nothing Then
      Set oCnx = New ADODB.Connection
   End If
   If Not oCnx.State = adStateClosed Then
      If Not oCnx.ConnectionString = sCnx Then oCnx.Close
   End If
   If Not oCnx.State = adStateOpen Then
      bNew = True
      oCnx.Open sCnx, sUsr, sPwd
   End If

End Sub

Public Sub DBCmdIni(oCnx As ADODB.Connection, oCmd As ADODB.Command)

' Initializes a Command object on a given connection

   If oCnx.State = adStateOpen Then
      If Not oCmd Is Nothing Then Set oCmd = Nothing
      If oCmd Is Nothing Then
         Set oCmd = New ADODB.Command  ' --- Create, assign connection
      End If
      Set oCmd.ActiveConnection = oCnx
   End If

End Sub

Public Function TblUpdArr(v(), sFld() As String, _
                          sCnx As String, sTbl As String, _
                          Optional lRowID As Long = 0) As Boolean

' Update a database table from a variant array
' The field order can be specified in sFld()
' if sFld() is not allocated, the field list
'    is generated ordered by field index
' The array columns are associated to the field
' with name at the respective sFld() position
' If a large number is to be updated, then the
' recordset is temporarily disconnected, appended,
' reconnected and updated in one batch.
' Returs true if something failed

   Dim oCnx As ADODB.Connection
   Dim oRcS As ADODB.Recordset
   Dim vF() As Variant, vV() As Variant
   Dim sOvr As String
   Dim lR1 As Long, lRL As Long, lRI As Long
   Dim b As Boolean, bBig As Boolean

   b = TblGetRcS(sCnx, sTbl, oCnx, oRcS, adOpenDynamic, adUseClient, _
                 adLockBatchOptimistic)
   If b Then GoTo TblUpdArr_Ende

   QRS_LibArr.ArrBoundV v(), lR1, lRL  ' --- Update row count
   sOvr = "/" & lRL                    '     String "over"

   bBig = lRL - lR1 > 600000           ' --- Big list
   If bBig Then                        ' --- Remove connection
      Set oRcS.ActiveConnection = Nothing
   End If

   b = QRS_LibLst.LstIsAllS(sFld())    ' --- Field list specified?
   If Not b Then RcSFldLstStr oRcS, sFld(), lRowID
   b = False                           '     No? -> Generate

   RcSFldLs4Upd sFld(), vF()           ' --- Prepare field list for .Update

   For lRI = lR1 To lRL
      TblRow4RSUpd sFld(), v(), lRI, vV()  ' Prepare row values list
      oRcS.AddNew vF(), vV()           ' --- Append from field and value lists
      If lRI Mod 37 = 0 Then           ' --- Set status bar message
         Application.StatusBar = "Adding " & lRI & sOvr
         DoEvents
      End If
      If Not bBig Then oRcS.UpdateBatch adAffectCurrent
   Next lRI

   If bBig Then                        ' --- Big list -> Reconnect
      Application.StatusBar = "Updating..."
      DoEvents
      Set oRcS.ActiveConnection = oCnx ' --- Reconnect recordset
      oRcS.UpdateBatch adAffectAll     ' --- Update current record in batch
   End If

   Application.StatusBar = False       ' --- Reset status bar
   DoEvents

TblUpdArr_Ende:

   ByeRcS oRcS
   ByeCnx oCnx

   TblUpdArr = b

End Function

Public Function TblGetRcS(sCnx As String, sTbl As String, _
                          oCnx As ADODB.Connection, oRcS As ADODB.Recordset, _
                Optional eCurTyp As ADODB.CursorTypeEnum = adOpenStatic, _
                Optional eCurLoc As ADODB.CursorLocationEnum = adUseClient, _
                Optional eLckTyp As ADODB.LockTypeEnum = adLockReadOnly) _
                          As Boolean

' Obtains a recordset representing a table
' Returns true if something failed

   DBConnect sCnx, "", "", oCnx

   Set oRcS = New ADODB.Recordset
   oRcS.CursorLocation = eCurLoc       ' --- Specify cursor location
   oRcS.Open sTbl, oCnx, eCurTyp, eLckTyp

   TblGetRcS = Not (oCnx.State = adStateOpen And oRcS.State = adStateOpen)

End Function

Public Function FcnGetRcS(sCnx As String, sFcn As String, sPrmVal As String, _
                          oCnx As ADODB.Connection, oRcS As ADODB.Recordset)
         
' Obtains a recordset represeting a table
' returned from by a table-valued function
' This version uses no parameter names, just the values

   Dim oCmd As ADODB.Command

   DBConnect sCnx, "", "", oCnx
   DBCmdIni oCnx, oCmd

   oCmd.CommandType = adCmdTable
   oCmd.CommandText = sFcn & "(" & sPrmVal & ")"

   Set oRcS = oCmd.Execute

End Function

Public Sub ByeRcS(oRcS As ADODB.Recordset)

   If oRcS Is Nothing Then Exit Sub

   If oRcS.State = adStateOpen Then oRcS.Close
   Set oRcS = Nothing

End Sub

Public Sub ByeCnx(oCnx As ADODB.Connection)

   If oCnx Is Nothing Then Exit Sub

   If oCnx.State = adStateOpen Then oCnx.Close
   Set oCnx = Nothing

End Sub

Public Sub ByeCmd(oCmd As ADODB.Command)

   If Not oCmd Is Nothing Then
      Set oCmd = Nothing
   End If

End Sub

Public Function TxtCurLoc(eCurLoc As ADODB.CursorLocationEnum) As String

' Texts to recordset cursor location enumeration values

   Dim s As String

   Select Case eCurLoc
   Case adUseClient
      s = "client"
   Case adUseClientBatch
      s = "client batch"
   Case adUseNone
      s = "none"
   Case adUseServer
      s = "server"
   End Select

   TxtCurLoc = s

End Function

Public Function TxtCurOpt(eCurOpt As ADODB.CursorOptionEnum) As String

' Texts to recordset cursor option enumeration values

   Dim s As String

   If eCurOpt And adAddNew = adAddNew Then
      s = s & "add New"
   End If
   If eCurOpt And adApproxPosition = adApproxPosition Then
      If Not s = "" Then s = s & MCsSep
      s = s & "approximate Position"
   End If
   If eCurOpt And adBookmark = adBookmark Then
      If Not s = "" Then s = s & MCsSep
      s = s & "bookmark"
   End If
   If eCurOpt And adFind = adFind Then
      If Not s = "" Then s = s & MCsSep
      s = s & "find"
   End If
   If eCurOpt And adHoldRecords = adHoldRecords Then
      If Not s = "" Then s = s & MCsSep
      s = s & "hold records"
   End If
   If eCurOpt And adIndex = adIndex Then
      If Not s = "" Then s = s & MCsSep
      s = s & "index"
   End If
   If eCurOpt And adMovePrevious = adMovePrevious Then
      If Not s = "" Then s = s & MCsSep
      s = s & "move previous"
   End If
   If eCurOpt And adNotify = adNotify Then
      If Not s = "" Then s = s & MCsSep
      s = s & "notify"
   End If
   If eCurOpt And adResync = adResync Then
      If Not s = "" Then s = s & MCsSep
      s = s & "resync"
   End If
   If eCurOpt And adSeek = adSeek Then
      If Not s = "" Then s = s & MCsSep
      s = s & "seek"
   End If
   If eCurOpt And adUpdate = adUpdate Then
      If Not s = "" Then s = s & MCsSep
      s = s & "update"
   End If
   If eCurOpt And adUpdateBatch = adUpdateBatch Then
      If Not s = "" Then s = s & MCsSep
      s = s & "update batch"
   End If

   TxtCurOpt = s

End Function

Public Function TxtCurTyp(eCurTyp As ADODB.CursorTypeEnum) As String

' Texts to recordset cursor type enumeration values

   Dim s As String

   Select Case eCurTyp
   Case adOpenDynamic
      s = "dynamic"
   Case adOpenForwardOnly
      s = "forward only"
   Case adOpenKeyset
      s = "keyset"
   Case adOpenStatic
      s = "static"
   Case adOpenUnspecified
      s = "unspecified2"
   End Select

   TxtCurTyp = s
End Function

Public Function TxtCmdTyp(eCmdTyp As ADODB.CommandTypeEnum) As String

' Texts to Connection Command type enumeration values

   Dim s As String

   Select Case eCmdTyp
   Case adCmdFile
      s = "command file"
   Case adCmdStoredProc
      s = "stored procedure"
   Case adCmdTable
      s = "table"
   Case adCmdTableDirect
      s = "direct table"
   Case adCmdText
      s = "command text"
   Case adCmdUnknown
      s = "unknown"
   Case adCmdUnspecified
      s = "unspecified"
   End Select

   TxtCmdTyp = s

End Function

Public Function TxtCnxMode(eCnxMode As ADODB.ConnectModeEnum)

' Texts to connection mode enumeration values

   Dim s As String

   If eCnxMode And adModeRead = adModeRead Then
      s = "read"
   End If
   If eCnxMode And adModeReadWrite = adModeReadWrite Then
      If Not s = "" Then s = s & MCsSep
      s = "read/write"
   End If
   If eCnxMode And adModeRecursive = adModeRecursive Then
      If Not s = "" Then s = s & MCsSep
      s = "recursive"
   End If
   If eCnxMode And adModeShareDenyNone = adModeShareDenyNone Then
      If Not s = "" Then s = s & MCsSep
      s = "share and deny none"
   End If
   If eCnxMode And adModeShareDenyRead = adModeShareDenyRead Then
      If Not s = "" Then s = s & MCsSep
      s = "share and deny read"
   End If
   If eCnxMode And adModeShareDenyWrite = adModeShareDenyWrite Then
      If Not s = "" Then s = s & MCsSep
      s = "share and deny write"
   End If
   If eCnxMode And adModeShareExclusive = adModeShareExclusive Then
      If Not s = "" Then s = s & MCsSep
      s = "share exclusive"
   End If
   If eCnxMode And adModeUnknown = adModeUnknown Then
      If Not s = "" Then s = s & MCsSep
      s = "unknown"
   End If
   If eCnxMode And adModeWrite = adModeWrite Then
      If Not s = "" Then s = s & MCsSep
      s = "write"
   End If

   TxtCnxMode = s

End Function

Public Function TxtCnxOpt(eCnxOpt As ADODB.ConnectOptionEnum)

' Texts to connection option enumeration values

   Dim s As String
   
   Select Case eCnxOpt
   Case adAsyncConnect
      s = "asynchronous"
   Case adConnectUnspecified
      s = "unspecified"
   End Select

   TxtCnxOpt = s
   
End Function

Public Function TxtBmkMode(eBmkMode As ADODB.BookmarkEnum)

' Texts to Recordset bookmark mode enumeration values

   Dim s As String

   Select Case eBmkMode
   Case adBookmarkCurrent
      s = "current"
   Case adBookmarkFirst
      s = "first"
   Case adBookmarkLast
      s = "last"
   End Select

   TxtBmkMode = s

End Function

Public Function TxtDatTyp(eDatTyp As ADODB.DataTypeEnum)

' Texts to record fields data type enumeration values

   Dim s As String

   If eDatTyp And adArray = adArray Then
      If Not s = "" Then s = s & MCsSep
      s = "array"
   End If
   If eDatTyp And adBigInt = adBigInt Then
      If Not s = "" Then s = s & MCsSep
      s = "big integer"
   End If
   If eDatTyp And adBinary = adBinary Then
      If Not s = "" Then s = s & MCsSep
      s = "binary"
   End If
   If eDatTyp And adBoolean = adBoolean Then
      If Not s = "" Then s = s & MCsSep
      s = "boolean"
   End If
   If eDatTyp And adBSTR = adBSTR Then
      If Not s = "" Then s = s & MCsSep
      s = "b string"
   End If
   If eDatTyp And adChapter = adChapter Then
      If Not s = "" Then s = s & MCsSep
      s = "chapter"
   End If
   If eDatTyp And adChar = adChar Then
      If Not s = "" Then s = s & MCsSep
      s = "character"
   End If
   If eDatTyp And adCurrency = adCurrency Then
      If Not s = "" Then s = s & MCsSep
      s = "currency"
   End If
   If eDatTyp And adDate = adDate Then
      If Not s = "" Then s = s & MCsSep
      s = "date"
   End If
   If eDatTyp And adDBDate = adDBDate Then
      If Not s = "" Then s = s & MCsSep
      s = "db date"
   End If
   If eDatTyp And adDBTime = adDBTime Then
      If Not s = "" Then s = s & MCsSep
      s = "db time"
   End If
   If eDatTyp And adDBTimeStamp = adDBTimeStamp Then
      If Not s = "" Then s = s & MCsSep
      s = "db timestamp"
   End If
   If eDatTyp And adDecimal = adDecimal Then
      If Not s = "" Then s = s & MCsSep
      s = "decimal (real)"
   End If
   If eDatTyp And adDouble = adDouble Then
      If Not s = "" Then s = s & MCsSep
      s = "double (real)"
   End If
   If eDatTyp And adEmpty = adEmpty Then
      If Not s = "" Then s = s & MCsSep
      s = "empty"
   End If
   If eDatTyp And adError = adError Then
      If Not s = "" Then s = s & MCsSep
      s = "error"
   End If
   If eDatTyp And adFileTime = adFileTime Then
      If Not s = "" Then s = s & MCsSep
      s = "file time"
   End If
   If eDatTyp And adGUID = adGUID Then
      If Not s = "" Then s = s & MCsSep
      s = "guid"
   End If
   If eDatTyp And adInteger = adInteger Then
      If Not s = "" Then s = s & MCsSep
      s = "integer"
   End If
   If eDatTyp And adLongVarBinary = adLongVarBinary Then
      If Not s = "" Then s = s & MCsSep
      s = "long variant binary"
   End If
   If eDatTyp And adLongVarChar = adLongVarChar Then
      If Not s = "" Then s = s & MCsSep
      s = "long variant character"
   End If
   If eDatTyp And adLongVarWChar = adLongVarWChar Then
      If Not s = "" Then s = s & MCsSep
      s = "long variant WChar"
   End If
   If eDatTyp And adNumeric = adNumeric Then
      If Not s = "" Then s = s & MCsSep
      s = "Numeric"
   End If
   If eDatTyp And adSingle = adSingle Then
      If Not s = "" Then s = s & MCsSep
      s = "single (real)"
   End If
   If eDatTyp And adSmallInt = adSmallInt Then
      If Not s = "" Then s = s & MCsSep
      s = "small integer"
   End If
   If eDatTyp And adTinyInt = adTinyInt Then
      If Not s = "" Then s = s & MCsSep
      s = "tiny integer"
   End If
   If eDatTyp And adUnsignedBigInt = adUnsignedBigInt Then
      If Not s = "" Then s = s & MCsSep
      s = "unsigned big integer"
   End If
   If eDatTyp And adUnsignedInt = adUnsignedInt Then
      If Not s = "" Then s = s & MCsSep
      s = "unsigned integer"
   End If
   If eDatTyp And adUnsignedSmallInt = adUnsignedSmallInt Then
      If Not s = "" Then s = s & MCsSep
      s = "unsigned small integer"
   End If
   If eDatTyp And adUnsignedTinyInt = adUnsignedTinyInt Then
      If Not s = "" Then s = s & MCsSep
      s = "unsigned tiny integer"
   End If
   If eDatTyp And adUserDefined = adUserDefined Then
      If Not s = "" Then s = s & MCsSep
      s = "user defined"
   End If
   If eDatTyp And adVarBinary = adVarBinary Then
      If Not s = "" Then s = s & MCsSep
      s = "variant binary"
   End If
   If eDatTyp And adVarChar = adVarChar Then
      If Not s = "" Then s = s & MCsSep
      s = "variant character"
   End If
   If eDatTyp And adVariant = adVariant Then
      If Not s = "" Then s = s & MCsSep
      s = "variant"
   End If
   If eDatTyp And adVarNumeric = adVarNumeric Then
      If Not s = "" Then s = s & MCsSep
      s = "numeric"
   End If
   If eDatTyp And adVarWChar = adVarWChar Then
      If Not s = "" Then s = s & MCsSep
      s = "variant WChar"
   End If
   If eDatTyp And adWChar = adWChar Then
      If Not s = "" Then s = s & MCsSep
      s = "WChar"
   End If
   
   TxtDatTyp = s

End Function

Public Function TxtEdiMode(eEdiMode As ADODB.EditModeEnum)

' Texts to recordset edit mode enumeration values

   Dim s As String

   If eEdiMode And adEditAdd = adEditAdd Then
      If Not s = "" Then s = s & MCsSep
      s = "add"
   End If
   If eEdiMode And adEditDelete = adEditDelete Then
      If Not s = "" Then s = s & MCsSep
      s = "delete"
   End If
   If eEdiMode And adEditInProgress = adEditInProgress Then
      If Not s = "" Then s = s & MCsSep
      s = "edit in progress"
   End If
   If eEdiMode And adEditNone = adEditNone Then
      If Not s = "" Then s = s & MCsSep
      s = "none"
   End If

   TxtEdiMode = s

End Function

Public Sub TxtFldPrp(oFld As ADODB.Field, bDebug As Boolean)

' Outputs the property names of a field

   Dim oPrp As ADODB.Property
   Dim sPrp As String

   If bDebug Then
      sPrp = QRS_LibStr.StrPad("Property", 20)
      sPrp = sPrp & QRS_LibStr.StrPad("Data type", 15)
      sPrp = sPrp & QRS_LibStr.StrPad("Value", 20)
      sPrp = sPrp & QRS_LibStr.StrPad("Attr", 10)
Debug.Print sPrp
      For Each oPrp In oFld.Properties
         sPrp = QRS_LibStr.StrPad(CStr(oPrp.Name), 20)
         sPrp = sPrp & QRS_LibStr.StrPad(TxtDatTyp(oPrp.Type), 15)
         If IsNull(oPrp.Value) Then sPrp = sPrp & String(20, " ") _
         Else sPrp = sPrp & QRS_LibStr.StrPad(CStr(oPrp.Value), 20)
         sPrp = sPrp & QRS_LibStr.StrPad(CStr(oPrp.Attributes), 10)
Debug.Print sPrp
      Next oPrp
   End If

End Sub

Public Sub TxtRcsPrp(oRcS As ADODB.Recordset, bDebug As Boolean)

' Outputs the property names of a field

   Dim oPrp As ADODB.Property
   Dim sPrp As String

   If bDebug Then
      sPrp = QRS_LibStr.StrPad("Property", 20)
      sPrp = sPrp & QRS_LibStr.StrPad("Data type", 15)
      sPrp = sPrp & QRS_LibStr.StrPad("Value", 20)
      sPrp = sPrp & QRS_LibStr.StrPad("Attr", 10)
Debug.Print sPrp
      For Each oPrp In oRcS.Properties
         sPrp = QRS_LibStr.StrPad(CStr(oPrp.Name), 20)
         sPrp = sPrp & QRS_LibStr.StrPad(TxtDatTyp(oPrp.Type), 15)
         If IsNull(oPrp.Value) Then sPrp = sPrp & String(20, " ") _
         Else sPrp = sPrp & QRS_LibStr.StrPad(CStr(oPrp.Value), 20)
         sPrp = sPrp & QRS_LibStr.StrPad(CStr(oPrp.Attributes), 10)
Debug.Print sPrp
      Next oPrp
   End If

End Sub

Public Function TxtXecOpt(eXecOpt As ADODB.ExecuteOptionEnum)

' Texts to command execution mode enumeration values

   Dim s As String

   Select Case eXecOpt
   Case adAsyncExecute
      s = "asynchronous execution"
   Case adAsyncFetch
      s = "asynchronous fetch"
   Case adAsyncFetchNonBlocking
      s = "asynchronous fetch non blocking"
   Case adExecuteNoRecords
      s = "execution returns no records"
   Case adExecuteRecord
      s = "execution returns records"
   Case adExecuteStream
      s = "execution streams"
   Case adOptionUnspecified
      s = "unspecified"
   End Select

   TxtXecOpt = s

End Function

Public Function TxtLckTyp(eLckTyp As ADODB.LockTypeEnum)

' Texts to recordset lock type enumeration values

   Dim s As String

   Select Case eLckTyp
   Case adLockBatchOptimistic
      s = "batch optimistic"
   Case adLockOptimistic
      s = "optimistic"
   Case adLockPessimistic
      s = "pessimistic"
   Case adLockReadOnly
      s = "read only"
   Case adLockUnspecified
      s = "unspecified"
   End Select
   
   TxtLckTyp = s

End Function

Public Function TxtPrmAtt(ePrmAtt As ADODB.ParameterAttributesEnum)

' Texts to command parameter attributes enumeration values

   Dim s As String

   If ePrmAtt And adParamLong = adParamLong Then
      s = "long"
   End If
   If ePrmAtt And adParamNullable = adParamNullable Then
      If Not s = "" Then s = s & MCsSep
      s = "nullable"
   End If
   If ePrmAtt And adParamSigned = adParamSigned Then
      If Not s = "" Then s = s & MCsSep
      s = "signed"
   End If

   TxtPrmAtt = s

End Function

Public Function TxtPrmDir(ePrmDir As ADODB.ParameterDirectionEnum)

' Texts to command parameter direction enumeration values

   Dim s As String

   Select Case ePrmDir
   Case adParamInput
      s = "input"
   Case adParamInputOutput
      s = "input/output"
   Case adParamOutput
      s = "output"
   Case adParamReturnValue
      s = "return value"
   Case adParamUnknown
      s = "unknown"
   End Select

   TxtPrmDir = s

End Function

Public Function TxtRecStat(eRecStat As ADODB.RecordStatusEnum)

' Texts to record status enumeration values

   Dim s As String

   Select Case eRecStat
   Case adRecCanceled
      s = "canceled"
   Case adRecCantRelease
      s = "can't release"
   Case adRecConcurrencyViolation
      s = "concurrency violation"
   Case adRecDBDeleted
      s = "database deleted"
   Case adRecDeleted
      s = "deleted"
   Case adRecIntegrityViolation
      s = "integrity violation"
   Case adRecInvalid
      s = "invalid"
   Case adRecMaxChangesExceeded
      s = "max changes exceeded"
   Case adRecModified
      s = "modified"
   Case adRecMultipleChanges
      s = "multiple changes"
   Case adRecNew
      s = "new"
   Case adRecObjectOpen
      s = "object open"
   Case adRecOK
      s = "record ok"
   Case adRecOutOfMemory
      s = "out of memory"
   Case adRecPendingChanges
      s = "pending changes"
   Case adRecPermissionDenied
      s = "permission denied"
   Case adRecSchemaViolation
      s = "schema violation"
   Case adRecUnmodified
      s = "unmodified"
   End Select

   TxtRecStat = s

End Function

Public Function TxtRecTyp(eRecTyp As ADODB.RecordTypeEnum)

' Texts to record type enumeration valueus

   Dim s As String

   Select Case eRecTyp
   Case adCollectionRecord
      s = "record collection"
   Case adSimpleRecord
      s = "simple record"
   Case adStructDoc
      s = "structured document"
   End Select

   TxtRecTyp = s

End Function

Public Sub RcSFldLstStr(oRcS As ADODB.Recordset, sFld() As String, _
                        Optional lFldID_Base1 As Long = 0)

' Returns a list of field names in oRcS ordered by the field index
' if lFldID_Base1 is > 0, this column is omitted in the list
' Tested and works OK along with RcsFldLs4Upd on 04.07.2018

   Const Cl01 As Long = 1, ClM1 As Long = -1

   Dim lI As Long, lJ As Long, lN As Long

   If oRcS Is Nothing Then Exit Sub

   With oRcS.Fields
      lN = .Count
      If lFldID_Base1 > 0 Then lN = lN + ClM1
      QRS_LibLst.LstAllocS sFld(), lN  ' --- Omit ID column if specified
      For lI = 1 To lN
         If lI = lFldID_Base1 Then lJ = lJ + Cl01
         With .Item(lJ)                ' --- ADODB is zero-based
            sFld(lI) = .Name           ' --- Field name
         End With
         lJ = lJ + Cl01
      Next lI
   End With

End Sub

Public Sub RcSFldStrLst(sFldLst As String, vF(), _
                        Optional sLstDel As String = ",")
                        
' Fills the variant field list array from the field list string

   Dim sTmp() As String
   Dim lE1 As Long, lEL As Long, lEI As Long

   QRS_LibStr.LstStrLst sFldLst, sTmp(), sLstDel
   QRS_LibLst.LstBoundS sTmp(), lE1, lEL
   ReDim vF(lE1 - 1 To lEL - 1)
   For lEI = lE1 To lEL
      vF(lEI - 1) = sTmp(lEI)
   Next lEI
                        
End Sub

Public Sub RcSFldNdx_ID(oRcS As ADODB.Recordset, sFld_ID As String, _
                        Optional bQrySysSQLSrv As Boolean = False)

' Finds the identifier field in the recordset fields list
' The only hint is the "AUTOINCREMENT" Property set to true

   Dim oFld As ADODB.Field

   For Each oFld In oRcS.Fields
      If oFld.Properties.Item("ISAUTOINCREMENT").Value Then Exit For
   Next oFld

   If Not oFld Is Nothing Then sFld_ID = oFld.Name Else sFld_ID = ""

End Sub

Public Sub RcSFldLs4Upd(sFld() As String, vFld() As Variant)

' Returns a zero-based variant field list for the update method

   Dim lE1 As Long, lEL As Long, lEI As Long

   QRS_LibLst.LstBoundS sFld(), lE1, lEL
   lEI = lEL - lE1

   If Not QRS_LibLst.LstIsAllV(vFld()) Then
      ReDim vFld(0 To lEI)             ' --- Not allocated : reallocate
   Else                                ' --- Different size: reallocate
      If Not (LBound(vFld()) = 0 And UBound(vFld()) = lEI) Then
         ReDim vFld(0 To lEI)
      End If
   End If

   For lEI = lE1 To lEL
      vFld(lEI - lE1) = sFld(lEI)
   Next lEI

End Sub

Public Sub RcSGetCol(oRcS As ADODB.Recordset, vV(), vF(), _
                     bNoR As Boolean)

' Extracts the values of all rows of the specified fields
' if vF() is not allocated, all fields are extracted,
' else   the fields refered to in vF() are extracted
' If the recordset contains no date, bNoR returns true

   Dim vT()
   Dim lNFld As Long
   Dim bSuppBM As Boolean
   Dim vBM

   bNoR = oRcS.EOF Or oRcS.BOF         ' --- No data
   If bNoR Then
      If Not Not vF() Then             ' --- All fields
         lNFld = oRcS.Fields.Count     '     Use recordset field count
      Else                             ' --- Specific fields -> List length
         lNFld = UBound(vF()) + 1 - LBound(vF())
      End If
      If lNFld > 0 Then ReDim vV(1 To 1, 1 To lNFld)
      Exit Sub
   End If

   bSuppBM = oRcS.Supports(adBookmark)
   If bSuppBM Then vBM = oRcS.Bookmark ' --- Bookmark current record

   If Not Not vF() Then                ' --- Field list specified    -> use
      vT() = oRcS.GetRows(adGetRowsRest, adBookmarkFirst, vF())
   Else                                ' --- No field list specified -> all
      vT() = oRcS.GetRows(adGetRowsRest, adBookmarkFirst)
   End If

   VarArrX01 vT(), vV(), True

   If bSuppBM Then oRcS.Bookmark = vBM ' --- Ensure to be at current reocrd

End Sub

Public Sub RcSGetRow(oRcS As ADODB.Recordset, vV(), vF(), _
                     bNoR As Boolean)

' Extracts the current row data from the specified recordset
' if vF() is not allocated, all fields are extracted,
' else  the fields referred to in vF() are extracted
' If the recordset contains no data, bNoR returns true
' In this case an empty values table is returned that
' is ready for use by the merge operation, but for
' updating the database, the AddNew method must be used

   Dim vT()
   Dim lNFld As Long
   Dim bSuppBM As Boolean
   Dim vBM

   bNoR = oRcS.EOF Or oRcS.BOF         ' --- No data so far?
   If bNoR Then
      If Not Not vF() Then             ' --- All fields
         lNFld = oRcS.Fields.Count     '     Use recordset field count
      Else                             ' --- Specific fields -> List length
         lNFld = UBound(vF()) + 1 - LBound(vF())
      End If                           ' --- Empty return array
      If lNFld > 0 Then ReDim vV(1 To 1, 1 To lNFld)
      Exit Sub
   End If

   bSuppBM = oRcS.Supports(adBookmark)
   If bSuppBM Then vBM = oRcS.Bookmark ' --- Bookmark current record

   If Not Not vF() Then                ' --- Field list specified?
      vT() = oRcS.GetRows(1, , vF())   '     Extract specific fields
   Else                                ' --- No field list specified?
      vT() = oRcS.GetRows(1)           '     Extract all fields
   End If
Stop
   VarArrX01 vT(), vV(), True

   If bSuppBM Then oRcS.Bookmark = vBM ' --- Ensure to be at current reocrd

End Sub

Public Sub TblRow4RSUpd(sFld() As String, _
                        v() As Variant, lRow As Long, _
                        vVal() As Variant)

' Returns vVal, a zero-based list of values
' for the .Update method along with vFld()
' vVal() is allocated only on the first call
' sFld may contain less fields than v() counts columns
' v() contains the values source table
' lRow  specifies the row to be prepared

   Dim lF1 As Long, lFL As Long, lFI As Long
                                       ' --- Field list width
   QRS_LibLst.LstBoundS sFld(), lF1, lFL

   If Not QRS_LibLst.LstIsAllV(vVal()) Then
      ReDim vVal(0 To lFL - lF1)       ' --- zero-based value list
   End If

   For lFI = lF1 To lFL                ' --- All fields
      vVal(lFI - lF1) = v(lRow, lFI)
   Next lFI

End Sub

Public Sub VarLstX01(sFld() As String, vCol())

' Inserts the 0-based field list sFld() into the
' 1-based field list row array vCol() for header output

   Const Cl01 As Long = 1

   If Not QRS_LibArr.ArrIsAllS(sFld()) Then Exit Sub

   Dim lE0 As Long, lEL As Long, lEI As Long

   QRS_LibLst.LstBoundS sFld(), lE0, lEL
   QRS_LibArr.ArrAllocV vCol(), Cl01, lEL + Cl01 - lE0

   For lEI = lE0 To lEL
      vCol(Cl01, lEI + Cl01) = sFld(lEI)
   Next lEI

End Sub

Public Sub VarArrX01(vRcs(), vArr(), Optional bClrRcs As Boolean = True)

' Transposes-shifts the vRcs() array obtained from a GetRows() recordset call
' to a QRS-usual array with 1-index base and a row in Dimension 1

   Const Cl01 As Long = 1

   Dim lR0 As Long, lRL As Long, lRI As Long
   Dim lC0 As Long, lCL As Long, lCI As Long

   If Not QRS_LibArr.ArrIsAllV(vRcs()) Then Exit Sub

   QRS_LibArr.ArrBoundV vRcs(), lC0, lCL, lR0, lRL
   QRS_LibArr.ArrAllocV vArr(), lRL + Cl01 - lR0, lCL + Cl01 - lC0

   For lCI = lC0 To lCL
      For lRI = lR0 To lRL
         vArr(lRI + Cl01, lCI + Cl01) = vRcs(lCI, lRI)
      Next lRI
   Next lCI

   If bClrRcs Then Erase vRcs()

End Sub

