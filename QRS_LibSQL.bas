Attribute VB_Name = "QRS_LibSQL"
Option Explicit

' Module : QRS_LibSQL
' Purpose: Some specialized SQL string assembly procedures
'          Contains a routine for SQL value formatting according
'          to the SQL Server data type identifier number
' By     : QRS, Roger Strebel
' Date   : 09.08.2018                  First working version of StrSQLValFmt
'          14.08.2018                  StrSQLValFmt evolved
'          15.08.2018                  StrSQLInsRow works
'          06.02.2019                  StrSQLDelRow works
' --- The public interface
'     StrSQLDelRow                     Delete a table row            06.02.2019
'     StrSQLInsRow                     Insert values of a table row  15.08.2018
'     StrSQLValFmt                     Format values for SQL         14.08.2018
'     StrSQLWhrStr                     Where clause string
'     StrVBAValFmt                     An attempt to format values for VBA

Public Function StrSQLInsRow(sTbl As String, sFld As String, _
                             sVal As String) As String

' IDEE: die Werte (für INSERT, UPDATE oder WHERE) sollen im Module ModQRS_ADX
'       unter Zuhilfenahme der Field Infos korrekt formatiert werden
'       und anschliessend in einen "Delimited String" vereint. Diese Delimited
'       strings sind als Argumente für folgende Funktionen sinnvoll:
'          StrSQLWhrStr (sCnd)
'          StrSQLWhrLst (sCnd), derzeit noch in ModQRS_ADX mit vVal()
'          StrSQLInsRow (sVal)
'       Die Funktion StrSQLValFmt konvertiert jeweils einen Wert
'       Fas Modul ModQRS_ADX muss eine Funktion erhalten, die die Field Infos
'       für jeden Eintrag einer Feldliste an StrSQLValFmt übergibt und die
'       Delimited Strings zusammenstellt

   Dim sCmd As String

   sCmd = "INSERT INTO <Tbl> (<Fld>) VALUES (<Val>)"
   sCmd = QRS_LibStr.StrRpl(sCmd, "<Tbl>", sTbl)
   sCmd = QRS_LibStr.StrRpl(sCmd, "<Fld>", sFld)
   sCmd = QRS_LibStr.StrRpl(sCmd, "<Val>", sVal)

   StrSQLInsRow = sCmd

End Function

Public Function StrSQLDelRow(sTbl As String, _
                             sFID As String, sVID As String) As String

' Assembles the DELETE SQL statement using the ID field name provided
' Attention: If either or both sFID and sVID are empty, the returned
'            command will delete all rows

   Dim sCmd As String

   sCmd = "DELETE FROM <Tbl>"
   sCmd = QRS_LibStr.StrRpl(sCmd, "<Tbl>", sTbl)
   If Not (sFID = "" Or sVID = "") Then
      sCmd = sCmd & " WHERE " & StrSQLWhrStr(sFID, "=", sVID, "")
   End If

   StrSQLDelRow = sCmd

End Function

Public Function StrSQLWhrStr(sFld As String, sCmp As String, sCnd As String, _
                             sOps As String, _
                             Optional sSepFld As String = ",") As String

' Assemble a WHERE clause from a field designation string
' containing one or more field names against conditions in sCnd

   Const sTpl As String = "<Fld> <Cmp> <Val>"

   Dim sCmd As String, sVal As String
   Dim lP As Long

   lP = InStr(1, sFld, sSepFld, vbTextCompare)
   If lP = 0 Then                      ' --- Just one field
      sCmd = QRS_LibStr.StrRpl(sTpl, "<Fld>", sFld)
      sCmd = QRS_LibStr.StrRpl(sCmd, "<Cmp>", sCmp)
      sCmd = QRS_LibStr.StrRpl(sCmd, "<Val>", sVal)
   Else

   End If

End Function

Public Function StrVBAValFmt(v, _
                Optional eValTyp As VbVarType = vbVariant) As String

' Returns a string in the format for SQL for several frequently used
' VBA datatypes:
'    String: in simple quotes
'    Numeric: without quotes
'    Date   : in single quotes, formatted '<YYYY>-<MM>-<DD>'

   Dim sVal As String

   If eValTyp = vbVariant Then         ' --- Variant - unspecified
      If IsNumeric(v) Then             '     Is numeric? -> assume double
         eValTyp = vbDouble
      Else                             '     Not numeric?
         If IsDate(v) Then             '     Is date?
            eValTyp = vbDate
         Else                          '     Assume string
            sVal = vbString
         End If
      End If
   End If

   Select Case eValTyp
   Case vbString                       ' --- String: enclose in quotes
      sVal = "'" & v & "'"
   Case vbDouble                       ' --- Double: double with point
      sVal = CDbl(v)
   Case vbLong                         ' --- Long  : long integer
      sVal = CLng(v)
   Case vbDate                         ' --- Date  : Y-M-D with quotes
      sVal = Format(v, "'YYYY-MM-DD'")
   Case vbBoolean                      ' --- Bool  : True = 1
      sVal = IIf(v, "1", "0")
   Case Else                           ' --- Other : no quotes
      sVal = v
   End Select

   StrVBAValFmt = sVal

End Function

Public Function StrSQLValFmt(v, lDataType As Long, lMaxLen As Long, _
                             Optional bFail As Boolean) As String

' Recognized data types are:
'    type_id   name
'      56      int
'     127      bigint
'      52      smallint
'      48      tinyint
'      62      float
'     106      decimal
'     104      bit
'      40      date
'      42      datetime2         YYYY-MM-DD hh:mm:ss[.fracsec]
'      43      datetimeoffset    YYYY-MM-DD hh:mm:ss[.nnnnnnn] [{+|-}hh:mm]
'      61      datetime          no default format, max 3 digits fracsec
'     167      varchar           ANSI string
'     231      nvarchar          unicode string
' this lookup is obtained by joining the columns and the types tables:
'     SELECT TS0.user_type_id, TS2.name FROM sys.all_columns
'     LEFT JOIN sys.types AS TS2 ON TS2.user_type_id = TS0.user_type_id
' MaxLen operates on the varchar and nvarchar types to cut excess length

   Const CsFmtDTm As String = "YYYY-MM-DD HH:MM:SS"
   Const CsNul As String = "NULL"

   Dim sVal As String

   Select Case lDataType
   Case 56, 127, 52, 48                ' --- int
      bFail = Not IsNumeric(v)
      If Not bFail Then sVal = v
   Case 62                             ' --- float
      bFail = Not IsNumeric(v)
      If Not bFail Then sVal = v
   Case 106                            ' --- decimal
   Case 40                             ' --- date
      If v = "" Then sVal = CsNul Else _
      sVal = "'" & Format(v, "YYYY-MM-DD") & "'"
   Case 42                             ' --- datetime2
   Case 43                             ' --- datetimeoffset
      sVal = "'" & Format(v, CsFmtDTm)
   Case 61                             ' --- datetime
   Case 104                            ' --- Boolean (0, 1)
      If v Then sVal = "1" Else sVal = "0"
   Case 167                            ' --- varchar
      sVal = "'" & Left(v, lMaxLen) & "'"
   Case 231                            ' --- nvarchar
      sVal = "'" & Left(v, lMaxLen) & "'"
   End Select

   If sVal = "" Then sVal = CsNul      ' --- Empty non-string -> NULL

   StrSQLValFmt = sVal

End Function


