Attribute VB_Name = "QRS_LibCnv"
Option Explicit

' Purpose: Some very basic typed array utility VBA routines
' Note   : SQL server distinguishes between unicode and ASCII data
'          for  ASCII  data, use the  VARCHAR data type
'          for Unicode data, use the NVARCHAR data type
' by     : QRS, Roger Strebel
' Date   : 07.07.2018                  Need for correct datetimeoffset strings
'          08.07.2018                  CnvVarDTO skips empty values, CnvMu2Str
'          10.07.2018                  CnvVarDTO blank, CnvMu2Str 2nd last
'          23.07.2018                  Note regarding NVarchar added
'          15.08.2018                  CnvStrSpc and CnvCrr2C1 added
' --- The public interface
'     CnvChkASC                        Check if ASCII OK
'     CnvCM2Str                        Special correction            07.07.2018
'     CnvCrr2C1                        Special 2-character corr      15.08.2018
'     CnvCrrStr                        Special correction            10.07.2018
'     CnvApoStr                        Special apostroph correcion   10.07.2018
'     CnvMu2Str                        Special mu correction         10.07.2018
'     CnvOHMStr                        Offset HHMM integer to string 07.07.2018
'     CnvStrSpc                        String special correction     15.08.2018
'     CnvVarCor                        Variant special correction    10.07.2018
'     CnvVarCut                        Variant column widht cut      07.07.2018
'     CnvVarDat                        Variant column SQL date       07.07.2018
'     CnvVarDTO                        Variant column datetimeoffset 10.07.2018
'     CnvVarSpc                        Variant special correction    07.07.2018

Public Sub CnvVarCut(v(), lCol As Long, lCut As String)

   Dim lR1 As Long, lRL As Long, lRI As Long
   Dim lC1 As Long, lCL As Long
   Dim s1 As String

   QRS_LibArr.ArrBoundV v(), lR1, lRL, lC1, lCL
   If lCol < lC1 Or lCol > lCL Then Exit Sub

   For lRI = lR1 To lRL
      s1 = v(lRI, lCol)
      If Len(s1) > lCut Then
         v(lRI, lCol) = Left(s1, lCut)
      End If
   Next lRI

End Sub

Public Sub CnvVarCor(v(), lCol As Long, sCorr As String)

' Correct an array column for ill translated encoding
' The position is identified by sBef and sAft strings
' and the ASC code of the character to be replaced
' Example "ABC?EFG", "BC;EF;D;63" identifies "BC?EF"
' and replaces it by "BCDEF"

   Dim sLst() As String
   Dim sBef As String, sAft As String, sIns As String
   Dim lR1 As Long, lRL As Long, lRI As Long
   Dim lC1 As Long, lCL As Long, lAC As Long

   If sCorr = "" Then Exit Sub

   QRS_LibArr.ArrBoundV v(), lR1, lRL, lC1, lCL
   If lCol < lC1 Or lCol > lCL Then Exit Sub

   QRS_LibStr.LstStrLst sCorr, sLst(), ";"
   sBef = sLst(1)
   sAft = sLst(2)
   sIns = sLst(3)
   lAC = CLng(sLst(4))

   For lRI = lR1 To lRL
      v(lRI, lCol) = CnvCrrStr(v(lRI, lCol), sBef, sAft, sIns, lAC)
   Next lRI

End Sub

Public Sub CnvVarDat(v(), lCol As Long, sFmt As String)

   Dim lR1 As Long, lRL As Long, lRI As Long
   Dim lC1 As Long, lCL As Long
   Dim v1

   QRS_LibArr.ArrBoundV v(), lR1, lRL, lC1, lCL
   If lCol < lC1 Or lCol > lCL Then Exit Sub

   For lRI = lR1 To lRL
      v1 = v(lRI, lCol)
      If Not v1 = "" Then
         v(lRI, lCol) = Format(v1, sFmt)
      End If
   Next lRI

End Sub

Public Sub CnvVarDTO(v(), lCol As Long, _
                     Optional lOffHHMM As Long = 0)

' Converts a date time format to the format required for
' ADO database insertion. Ignores empty values
' if the offset is present ("+" or "-" present),
' and ":" is not present, ":" is inserted
' if the offset is not present, the lOffHHMM is appended

   Const CsFmtDTm As String = "YYYY-MM-DD HH:MM:SS"

   Dim lR1 As Long, lRL As Long, lRI As Long
   Dim lC1 As Long, lCL As Long
   Dim lO As Long, lP As Long, lQ As Long, lR As Long
   Dim s As String
   Dim bOff As Boolean
                                       ' --- Table size
   QRS_LibArr.ArrBoundV v(), lR1, lRL, lC1, lCL
   If lCol < lC1 Or lCol > lCL Then Exit Sub

   lR = Len(CsFmtDTm)
   For lRI = lR1 To lRL                ' --- Row by row
      s = v(lRI, lCol)                 ' --- Column value to string
      If Not s = "" Then               ' --- Value not empty
         s = QRS_LibStr.StrRpl(s, "T", " ")
         s = QRS_LibStr.StrRpl(s, "Z", " ")
         lP = InStr(1, s, "+")         ' --- Plus sign for offset?
         If lP = 0 Then                ' --- No: Search plus sign near end
            lP = QRS_LibStr.StrInR(s, "-")
            If lP < Len(s) - 5 Then lP = 0
         End If
         bOff = lP > 0                 ' --- lP > 0: Offset found
         If bOff Then
            lQ = InStr(lP, s, ":")        ' --- colon in offset
            If lQ = 0 Then                ' --- No colon -> use 5 right chars
               lO = CLng(Right(s, 5))
            Else                          ' --- Colon -> 2 and of 3 chars
               lO = CLng(Right(s, 2) + CLng(Mid(s, lP, 3))) * 100
            End If
            v(lRI, lCol) = Format(CDate(Left(s, lR)), CsFmtDTm) & CnvOHMStr(lO)
         Else
            v(lRI, lCol) = Format(s, CsFmtDTm) & CnvOHMStr(lOffHHMM)
         End If
      End If
   Next lRI

End Sub

Public Function CnvStrSpc(sTxt As String) As String

' Corrects non-ascii characters in ill encoded strings
' This version stops if non-ascii characters remain

   Dim sStr As String

   sStr = CnvCrr2C1(sTxt, 195, 182, "ö")
   sStr = CnvCrr2C1(sStr, 195, 188, "ü")
   sStr = CnvCrr2C1(sStr, 195, 164, "ä")
   sStr = CnvCrr2C1(sStr, 195, 168, "è")
   sStr = CnvCrr2C1(sStr, 195, 169, "é")
   sStr = CnvCrr2C1(sStr, 195, 167, "ç")
   sStr = CnvCrr2C1(sStr, 195, 132, "Ä")
   sStr = CnvCrr2C1(sStr, 195, 150, "Ö")
   sStr = CnvCrr2C1(sStr, 195, 156, "Ü")
   sStr = CnvCrr2C1(sStr, 206, 188, "µ")

   If CnvChkASC(sStr) Then
Debug.Print "Invalid string: " & sStr
Stop
   End If

   CnvStrSpc = sStr

End Function

Public Sub CnvVarSpc(v(), lCol As Long)

   Dim lR1 As Long, lRL As Long, lRI As Long
   Dim lC1 As Long, lCL As Long

   Application.StatusBar = "Correcting mu..."
   DoEvents

   QRS_LibArr.ArrBoundV v(), lR1, lRL, lC1, lCL
   If lCol < lC1 Or lCol > lCL Then Exit Sub

   For lRI = lR1 To lRL
      v(lRI, lCol) = CnvMu2Str(v(lRI, lCol))
'     v(lRI, lCol) = CnvCM2Str(v(lRI, lCol))
   Next lRI

   Application.StatusBar = False
   DoEvents

End Sub

Public Function CnvOHMStr(lOffHHMM As Long) As String

' Returns the offset string with colon

   Const Cl100 As Long = 100

   Dim lH As Long, lM As Long

   lM = lOffHHMM Mod Cl100
   lH = (lOffHHMM - lM) / Cl100

   CnvOHMStr = " " & IIf(lOffHHMM < 0, "", "+") & _
               Format(lH, "00") & ":" & Format(lM, "00")

End Function

Public Function CnvChkASC(sTxt As String) As Boolean

' Returns true if any two succeeding the characters in sTxt
' are not ASCII characters

   Const Cl01 As Long = 1

   Dim lL As Long, lI As Long

   lL = Len(sTxt) - Cl01
   For lI = Cl01 To lL
      If Asc(Mid(sTxt, lI, Cl01)) > 127 Then _
         If Asc(Mid(sTxt, lI + Cl01, Cl01)) > 127 Then Exit For
   Next lI

   CnvChkASC = Not lI > lL

End Function

Public Function CnvCrr2C1(sTxt As String, lAsc1 As Long, lAsc2 As Long, _
                          sBy As String) As String

' This replacement function handles the special case of a text
' containing two subsequent non-ASCII-characters to be replaced
' by one character

   Const Cl01 As Long = 1

   Dim sStr As String
   Dim lP1 As Long, lP2 As Long

   sStr = sTxt
   lP1 = InStr(Cl01, sStr, Chr(lAsc1), vbTextCompare)
   While Not lP1 = 0
      If lP1 > 0 Then
         lP2 = InStr(lP1, sStr, Chr(lAsc2), vbTextCompare)
         If lP2 > 0 Then
            sStr = Left(sStr, lP1 - Cl01) & sBy & Mid(sStr, lP2 + Cl01)
         End If
      End If
      lP1 = InStr(lP1 + Cl01, sStr, Chr(lAsc1), vbTextCompare)
   Wend

   CnvCrr2C1 = sStr

End Function

Public Function CnvCrrStr(v, sBef As String, sAft As String, sIns As String, _
                          Optional lCQM As Long = 63) As String

' This replacement function handles the special case of a question mark
' in a string between sBef and sAft. If found, the question mark is
' replaced by sIns

   Const Cl01 As Long = 1

   Dim lA As Long, lB As Long, lI As Long, lP As Long, lQ As Long, lS As Long
   Dim s As String

   If v = "" Then Exit Function

   s = v
   lA = Len(sAft)
   lB = Len(sBef)
   lI = Len(sIns)
   lP = InStr(lP + Cl01, s, sBef)
   lQ = lP + lB
   lS = Len(s)
   While lP > 0 And Not lQ + Cl01 + lA > lS
      If Asc(Mid(s, lQ, Cl01)) = lCQM And _
         Mid(s, lQ + Cl01, lA) = sAft Then
         s = Left(s, lQ - Cl01) & sIns & Mid(s, lQ + Cl01)
         lS = Len(s)
      End If
      lP = InStr(lQ + Cl01 + lA, s, sBef)
      lQ = lP + lB
   Wend

   CnvCrrStr = s

End Function

Public Function CnvMu2Str(v) As String

' This replacement function handles the special case of
' vba input with ill conversion of the mu ("µ") character
' which is converted into "?" and uppercases the next character
' The routine replaces the "?" if it is followed by an uppercase
' character by the mu and lowercases the following character

   Const Cl01 As Long = 1
   Const Cl32 As Long = 32

   Dim lA As Long, lP As Long, lS As Long
   Dim s As String

   s = v
   lS = Len(s)
   For lP = lP + 1 To lS
      If Asc(Mid(s, lP, 1)) = 63 Then Exit For
   Next lP
   If lP > lS Then lP = 0
'   lP = InStr(lP + Cl01, s, "")
   While lP > 0
      If Not lP + Cl01 > lS Then
         lA = Asc(Mid(s, lP + Cl01))
         If (lA And Cl32) = 0 Then
            s = Left(s, lP - Cl01) & "µ" & Chr(lA Or Cl32) & Mid(s, lP + 2)
         End If
      End If
      For lP = lP + 1 To lS
         If Asc(Mid(s, lP, 1)) = 63 Then Exit For
      Next lP
      If lP > lS Then lP = 0
'      lP = InStr(lP + Cl01, s, "?")
   Wend

   CnvMu2Str = s

End Function

Public Function CnvCM2Str(v, _
                          Optional sWhat As String = ",", _
                          Optional sSpec As String = "m", _
                          Optional lOffM As Long = -2) As String

' This replacement routine handles the special case
' in XML processing when a sequence "Mx," is read by
' RefGetArrV and transformed into   "?x,"

   Dim lP As Long, lM As Long, lW As Long
   Dim s As String, t As String

   s = v

   lW = Len(sWhat)
   lM = Len(sSpec)
   lP = InStr(lP + 1, s, sWhat)
   While lP > 0
      If lP > -lOffM Then
         If Asc(Mid(s, lP + lOffM, 1)) = 63 Then
            t = Left(s, lP + lOffM - lM)
            t = t & sSpec & Mid(s, lP + lOffM + lW, lW - (lM + lOffM))
            t = t & Right(s, Len(s) - lP)
            s = t
         End If
      End If
      lP = InStr(lP + lM, s, sWhat)
   Wend

   CnvCM2Str = t

End Function
