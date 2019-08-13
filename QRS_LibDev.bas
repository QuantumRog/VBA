Attribute VB_Name = "QRS_LibDev"
Option Explicit

' Module : QRS_LibDev
' Project: any
' Purpose: Some utility VBA routines for developpment and debugging purposes
' By     : QRS, Roger Strebel
' Date   : 04.03.3018                  StrPad added and tested
'          14.03.2018                  Default boolean argument, LstDebugD
'          25.03.2018                  LstDebugF added
'          30.03.2018                  Binary and Hex function added
'          02.04.2018                  LstDebugL added
'          29.07.2018                  LstDebugS added
'          07.02.2019                  LstDebugV, ClrImmedi added
' --- The public interface
'     ArrDebugF                        Output real array elements    04.03.2018
'     ArrDebugL                        Output long array elements    04.03.2018
'     ArrDebugS                        Output string array elements  04.03.2018
'     ArrDebugV                        Output variant array elements 04.03.2018
'     BooleanFT                        Output boolean "F" or "T"
'     ClrImmedi                        Clear immediate window        07.02.2019
'     Lon2BinS                         Binary string of long         30.03.2018
'     Lon2HexS                         Hexadecimal string of long    30.03.2018
'     LonFBinS                         Long from string binary       30.03.2018
'     LonFHexS                         Long from string hexadecimal  30.03.2018
'     LstDebugD                        Output date list elements     14.03.2018
'     LstDebugF                        Output real list elements     25.03.2018
'     LstDebugL                        Output real list elements     02.04.2018
'     LstDebugS                        Output text list elemeents    29.07.2018
'     LstDebugV                        Output variant list elements  07.02.2019
'     WbShNames                        Output all names lists        14.03.2018
' --- The private sphere
'     NamesDebg                        Outputs names of collection

Public Function ClrImmedi()

' Clears the immediate window
' Credits to
' https://stackoverflow.com/questions/10203349/
'use-vba-to-clear-immediate-window
' Note: Does put the focus on the immmediate window
'       and may sometimes suppress consecutive output

'   Application.SendKeys "^g^a", True
'   DoEvents
'   Application.SendKeys "^g{DEL}", True
'   DoEvents

Debug.Print String(12, vbNewLine)      ' --- Variant avoiding SendKeys

End Function

Public Function Lon2HexS(l As Long) As String

' Returns hexadecimal string representation of a 32-bit integer value
' VBA provides the Hex() function which returns the hex representation
' of a 16-bit integer number
' Note: VBA has no unsigned integer type. When the most
'       significant bit is set, the numer is "negative"

   Const ClW As Long = 65535           ' --- 16 least significant bits
   Const Cl1 As Long = 65536

   Dim lLSW As Long, lMSW As Long
   Dim s As String

   lLSW = l And ClW                    ' --- Mask out least significant word
   lMSW = (l - lLSW) / Cl1 And ClW     ' --- Mask out most significant word
                                       '     And shift right by 16 digits
   s = Hex(lLSW)                       ' --- No leading zeros
   If Not lMSW = 0 Then                '     More thant 16 bits
      s = QRS_LibStr.StrPad(s, 4, "R", "0")
      s = Hex(lMSW) & s                '     Pad MSW to 4 hex digits
   End If
   Lon2HexS = s

End Function

Public Function Lon2BinS(l As Long) As String

' Returns binary string representation of a 32-bit integer value
' The bits are tested for 0...30. Bit 31 is tested by the sign

   Const Cl01 As Long = 1, Cl02 As Long = 2
   Const Cs0 As String = "0", Cs1 As String = "1"

   Dim lMsk As Long, lBit As Long
   Dim s As String

   lMsk = Cl01
   If (l And lMsk) > 0 Then s = Cs1 Else s = Cs0
   For lBit = 1 To 30
      lMsk = lMsk * Cl02
      If (l And lMsk) > 0 Then s = Cs1 & s Else s = Cs0 & s
   Next lBit
   If l < 0 Then s = Cs1 & s Else s = Cs0 & s
   Lon2BinS = s

End Function

Public Function LonFBinS(sBin As String) As Long

' Encapsulates the function of the string library

   LonFBinS = QRS_LibStr.StrBinLon(sBin)

End Function


Public Function LonFHexS(sHex As String) As Long

' Encapsulates the function of the string libraray

   LonFHexS = QRS_LibStr.StrHexLon(sHex)

End Function

Public Sub LstDebugD(d() As Date, Optional bSize As Boolean = True, _
                     Optional sFmt As String = "")

   Dim sMsg As String
   Dim lE As Long, lE1 As Long, lEL As Long

   If Not QRS_LibLst.LstIsAllD(d()) Then Exit Sub
   lE1 = LBound(d(), 1): lEL = UBound(d(), 1)

   If bSize Then
      sMsg = "List bounds " & lE1 & " to " & lEL & " elements"
Debug.Print sMsg
   End If

   If sFmt = "" Then
      For lE = lE1 To lEL
Debug.Print d(lE)
      Next lE
   Else
      For lE = lE1 To lEL
Debug.Print Format(d(lE), sFmt)
      Next lE
   End If

End Sub

Public Sub LstDebugF(f() As Double, Optional bSize As Boolean = True, _
                     Optional sFmt As String = "")

   Dim sMsg As String
   Dim lE As Long, lE1 As Long, lEL As Long

   If Not QRS_LibLst.LstIsAllF(f()) Then Exit Sub
   lE1 = LBound(f(), 1): lEL = UBound(f(), 1)

   If bSize Then
      sMsg = "List bounds " & lE1 & " to " & lEL & " elements"
Debug.Print sMsg
   End If

   If sFmt = "" Then
      For lE = lE1 To lEL
Debug.Print f(lE)
      Next lE
   Else
      For lE = lE1 To lEL
Debug.Print Format(f(lE), sFmt)
      Next lE
   End If

End Sub

Public Sub LstDebugL(l() As Long, Optional bSize As Boolean = True, _
                     Optional sFmt As String = "")

   Dim sMsg As String
   Dim lE As Long, lE1 As Long, lEL As Long

   If Not QRS_LibLst.LstIsAllL(l()) Then Exit Sub
   lE1 = LBound(l(), 1): lEL = UBound(l(), 1)

   If bSize Then
      sMsg = "List bounds " & lE1 & " to " & lEL & " elements"
Debug.Print sMsg
   End If

   If sFmt = "" Then
      For lE = lE1 To lEL
Debug.Print l(lE)
      Next lE
   Else
      For lE = lE1 To lEL
Debug.Print Format(l(lE), sFmt)
      Next lE
   End If

End Sub

Public Sub LstDebugS(s() As String, Optional bSize As Boolean = True, _
                                    Optional bIndex As Boolean = True)

   Dim sMsg As String
   Dim lE As Long, lE1 As Long, lEL As Long

   If Not QRS_LibLst.LstIsAllS(s()) Then Exit Sub
   lE1 = LBound(s(), 1): lEL = UBound(s(), 1)

   If bSize Then
      sMsg = "List bounds " & lE1 & " to " & lEL & " elements"
Debug.Print sMsg
   End If

   If bIndex Then
      For lE = lE1 To lEL
Debug.Print Format(lE, "### ") & s(lE)
      Next lE
   Else
      For lE = lE1 To lEL
Debug.Print s(lE)
      Next lE
   End If

End Sub

Public Sub LstDebugV(v(), Optional bSize As Boolean = True, _
                          Optional bIndex As Boolean = True)

   Dim sMsg As String
   Dim lE As Long, lE1 As Long, lEL As Long

   If Not QRS_LibLst.LstIsAllV(v()) Then Exit Sub
   lE1 = LBound(v(), 1): lEL = UBound(v(), 1)

   If bSize Then
      sMsg = "List bounds " & lE1 & " to " & lEL & " elements"
Debug.Print sMsg
   End If

   For lE = lE1 To lEL
      sMsg = ""
      If IsMissing(v(lE)) Then sMsg = "#missing"
      If IsEmpty(v(lE)) Then sMsg = "#empty"
      If IsNull(v(lE)) Then sMsg = "#null"
      If sMsg = "" Then sMsg = v(lE)
      If bIndex Then sMsg = Format(lE, "### ") & sMsg
Debug.Print sMsg
   Next lE

End Sub

Public Sub ArrDebugF(f() As Double, Optional bSize As Boolean = True, _
                     Optional sFmt As String = "")

' Outputs array elements to the "immediate" window
' if bSize is true, outputs array bounds first
' sFmt may contain a formatting string for numeric output

   Const Cl01 As Long = 1
   Const CsSep As String = " "

   Dim sMsg As String
   Dim lR As Long, lR1 As Long, lRL As Long
   Dim lC As Long, lC1 As Long, lCL As Long

   If Not QRS_LibArr.ArrIsAllF(f()) Then Exit Sub
   lR1 = LBound(f(), 1): lRL = UBound(f(), 1)
   lC1 = LBound(f(), 2): lCL = UBound(f(), 2)

   If bSize Then
      sMsg = "Array bounds "
      sMsg = sMsg & lR1 & " to " & lRL & " rows, "
      sMsg = sMsg & lC1 & " to " & lCL & " cols"
Debug.Print sMsg
   End If

   If sFmt = "" Then
      For lR = lR1 To lRL
         sMsg = f(lR, lC1)
         For lC = lC1 + Cl01 To lCL
            sMsg = sMsg & CsSep & f(lR, lC)
         Next lC
Debug.Print sMsg
      Next lR
   Else
      For lR = lR1 To lRL
         sMsg = Format(f(lR, lC1), sFmt)
         For lC = lC1 + Cl01 To lCL
            sMsg = sMsg & CsSep & Format(f(lR, lC), sFmt)
         Next lC
Debug.Print sMsg
      Next lR
   End If

End Sub

Public Sub ArrDebugL(l() As Long, Optional bSize As Boolean = True, _
                     Optional sFmt As String = "")

' Outputs array elements to the "immediate" window
' if bSize is true, outputs array bounds first
' sFmt may contain a formatting string for numeric output

   Const Cl01 As Long = 1
   Const CsSep As String = " "

   Dim sMsg As String
   Dim lR As Long, lR1 As Long, lRL As Long
   Dim lC As Long, lC1 As Long, lCL As Long

   If Not QRS_LibArr.ArrIsAllL(l()) Then Exit Sub
   lR1 = LBound(l(), 1): lRL = UBound(l(), 1)
   lC1 = LBound(l(), 2): lCL = UBound(l(), 2)

   If bSize Then
      sMsg = "Array bounds "
      sMsg = sMsg & lR1 & " to " & lRL & " rows, "
      sMsg = sMsg & lC1 & " to " & lCL & " cols"
Debug.Print sMsg
   End If

   If sFmt = "" Then
      For lR = lR1 To lRL
         sMsg = l(lR, lC1)
         For lC = lC1 + Cl01 To lCL
            sMsg = sMsg & CsSep & l(lR, lC)
         Next lC
Debug.Print sMsg
      Next lR
   Else
      For lR = lR1 To lRL
         sMsg = Format(l(lR, lC1), sFmt)
         For lC = lC1 + Cl01 To lCL
            sMsg = sMsg & CsSep & Format(l(lR, lC), sFmt)
         Next lC
Debug.Print sMsg
      Next lR
   End If

End Sub

Public Sub ArrDebugS(s() As String, Optional bSize As Boolean = True, _
                     Optional lWdth As Long = 0)

' Outputs array elements to the "immediate" window
' if bSize is true, outputs array bounds first
' lWdth contains an optional column width, if = 0, full values are output

   Const Cl01 As Long = 1
   Const CsSep As String = " "

   Dim sMsg As String
   Dim lR As Long, lR1 As Long, lRL As Long
   Dim lC As Long, lC1 As Long, lCL As Long

   If Not QRS_LibArr.ArrIsAllS(s()) Then Exit Sub
   lR1 = LBound(s(), 1): lRL = UBound(s(), 1)
   lC1 = LBound(s(), 2): lCL = UBound(s(), 2)

   If bSize Then
      sMsg = "Array bounds "
      sMsg = sMsg & lR1 & " to " & lRL & " rows, "
      sMsg = sMsg & lC1 & " to " & lCL & " cols"
Debug.Print sMsg
   End If

   If lWdth = 0 Then
      For lR = lR1 To lRL
         sMsg = s(lR, lC1)
         For lC = lC1 + Cl01 To lCL
            sMsg = sMsg & CsSep & s(lR, lC)
         Next lC
Debug.Print sMsg
      Next lR
   Else
      For lR = lR1 To lRL
         sMsg = QRS_LibStr.StrPad(s(lR, lC), lWdth)
         For lC = lC1 + Cl01 To lCL
            sMsg = sMsg & CsSep & QRS_LibStr.StrPad(s(lR, lC), lWdth)
         Next lC
Debug.Print sMsg
      Next lR
   End If

End Sub

Public Sub ArrDebugV(v() As Variant, Optional bSize As Boolean = True, _
                     Optional sFmt As String = "")

' Outputs array elements to the "immediate" window
' if bSize is true, outputs array bounds first
' sFmt may contain a formatting string for numeric output

   Const Cl01 As Long = 1
   Const CsSep As String = " "

   Dim sMsg As String
   Dim lR As Long, lR1 As Long, lRL As Long
   Dim lC As Long, lC1 As Long, lCL As Long

   If Not QRS_LibArr.ArrIsAllV(v()) Then Exit Sub
   lR1 = LBound(v(), 1): lRL = UBound(v(), 1)
   lC1 = LBound(v(), 2): lCL = UBound(v(), 2)

   If bSize Then
      sMsg = "Array bounds "
      sMsg = sMsg & lR1 & " to " & lRL & " rows, "
      sMsg = sMsg & lC1 & " to " & lCL & " cols"
Debug.Print sMsg
   End If

   If sFmt = "" Then
      For lR = lR1 To lRL
         sMsg = v(lR, lC1)
         For lC = lC1 + Cl01 To lCL
            sMsg = sMsg & CsSep & v(lR, lC)
         Next lC
Debug.Print sMsg
      Next lR
   Else
      For lR = lR1 To lRL
         sMsg = Format(v(lR, lC1), sFmt)
         For lC = lC1 + Cl01 To lCL
            sMsg = sMsg & CsSep & Format(v(lR, lC), sFmt)
         Next lC
Debug.Print sMsg
      Next lR
   End If

End Sub

Public Sub BooleanFT(b As Boolean)

   Dim s As String

   If b Then s = "T" Else s = "F"

Debug.Print s

End Sub

Public Sub WbShNames(aWb As Workbook, _
                     Optional bWb As Boolean = True, _
                     Optional bSh As Boolean = True)

' Outputs all names in the workbook and its sheets
' if bWb is set true, outputs the workbook names
' if bSh is set true, outputs the sheet names

   Dim aSh As Worksheet

   If bWb Then
Debug.Print "Workbook: " & aWb.Name
      NamesDebg aWb.Names
   End If
   If bSh Then
      For Each aSh In aWb.Worksheets
Debug.Print "Worksheet: " & aSh.Name
         NamesDebg aSh.Names
      Next aSh
   End If

End Sub

Private Sub NamesDebg(aNames As Names)

   Const lPad As Long = 25

   Dim aName As Name

   If aNames.Count = 0 Then
Debug.Print "(none)"
   Else
      For Each aName In aNames
Debug.Print QRS_LibStr.StrPad(aName.Name, lPad) & aName.RefersTo & aName.RefersToRange
      Next aName
   End If

End Sub
