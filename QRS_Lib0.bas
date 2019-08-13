Attribute VB_Name = "QRS_Lib0"
Option Explicit

' Module : QRS_Lib0
' Project: any
' Purpose: Some very basic utility VBA routines
'          *** An important hint for VBA font size settings ***
'          At high resolutions the font size settings have no effect
'          In order for font size changes to take effect, set the
'          screen resolution to 900x1600, close the session and reopen
'          The font size may then be selected and becomes effective
'          When the resolution is upscaled again, the font size
'          remains on the set one
' By     : QRS, Roger Strebel
' Date   : 21.01.2018
'          18.02.2018                  DivMod functions added
'          19.03.2018                  FloorF function added, DivMod for fVal<0
'          02.04.2018                  Factorial functions added
'          29.06.2018                  WithinL function added
'          27.07.2018                  RngWbShRg utility aded
'          11.08.2018                  RngWbShRg utility improved
'          01.11.2018                  MaxOfL, MinOfL added
'          10.08.2019                  DivModL function slightly modified
' --- The public interface
'     BitGetMSB                        Return index and value of MSB 26.03.2018
'     CnvStrTyp                        Convert string to typed var   19.03.2018
'     CnvVarTyp                        Convert variant to typed var  19.03.2018
'     DivModDbl                        Division and remainder (dbl)  19.03.2018
'     DivModLon                        Division and remainder (long) 10.08.2019
'     FactlF                           Factorial of real             02.04.2018
'     FactlL                           Factorial of integer          02.04.2018
'     FloorF                           Next lower integer            19.03.2018
'     MaxOfL                           Max value of some integers    01.11.2018
'     MinOfL                           Min value of some integers    01.11.2018
'     PathApp                          Append folder to path         21.01.2018
'     RngWbShRg                        Excel objects quick utility   11.08.2018
'     WithinL                          Is within interval            29.06.2018
' --- The private sphere

Public Sub CnvStrTyp(sStr As String, a, Optional bTrim As Boolean = False)

   Select Case VarType(a)
   Case VbVarType.vbBoolean
      Dim sBoo As String
      sBoo = UCase(Left(sStr, 1))
      a = sBoo = "Y" Or sBoo = "T" Or sBoo = "J" Or sBoo = "1" Or sStr = "-1"
   Case VbVarType.vbDate
      a = CDate(sStr)
   Case VbVarType.vbDouble
      a = CDbl(sStr)
   Case VbVarType.vbLong
      a = CLng(sStr)
   Case VbVarType.vbString
      If bTrim Then a = Trim(sStr) Else a = sStr
   Case VbVarType.vbVariant
      a = sStr
   End Select

End Sub

Public Sub CnvVarTyp(vVar As Variant, a)

   Select Case VarType(a)
   Case VbVarType.vbBoolean            ' --- Boolean value conversion
      a = vVar Or vVar = 1 Or vVar = -1
      If Not a Then                    '     Text value conversion
         Dim sBoo As Variant
         sBoo = UCase(Left(vVar, 1))
         a = sBoo = "Y" Or sBoo = "T" Or sBoo = "J"
      End If
   Case VbVarType.vbDate
      a = CDate(vVar)
   Case VbVarType.vbDouble
      a = CDbl(vVar)
   Case VbVarType.vbLong
      a = CLng(vVar)
   Case VbVarType.vbString
      a = vVar
   Case VbVarType.vbVariant
      a = vVar
   End Select

End Sub

Public Sub DivModDbl(fVal As Double, fDiv As Double, _
                     Optional fQtn As Double, Optional fRmd As Double)

' Returns quotient and remainder of lVal/lDiv
' 19.03.2018: Handles negative fVal correctly

   Dim fTmp As Long                    ' --- Allow return by reference

   If fDiv = 0 Then                    ' --- Exception: fDiv=0
      fQtn = 0                         '     Avoid static remanence
      fRmd = 0
      Exit Sub
   End If

   fTmp = fVal
   fQtn = FloorF(fTmp / fDiv)
   fRmd = fTmp - fQtn * fDiv

End Sub

Public Sub DivModLon(lVal As Long, lDiv As Long, _
                     Optional lQtn As Long, Optional lRmd As Long)

' Returns quotient and remainder of lVal/lDiv
' 10.08.2019: Handles negative lVal correctly

   Const Cl01 As Long = -1

   Dim lTmp As Long                    ' --- Allow return by reference

   If lDiv = 0 Then                    ' --- Exception: lDiv=0
      lQtn = 0                         '     Avoid static remanence
      lRmd = 0
      Exit Sub
   End If

   lTmp = lVal
   lQtn = Fix(lTmp / lDiv)
   lRmd = lTmp - lQtn * lDiv
   If lRmd < 0 Then
      lRmd = lRmd + lDiv
      lQtn = lQtn + Cl01
   End If

End Sub

Public Function FactlF(fVal As Double) As Double

' Returns the factorial fVal! of real value fVal

   Const Cf01 As Double = 1#, Cf02 As Double = 2#

   Dim fTmp As Double, fInc As Double

   fTmp = Cf01
   For fInc = Cf01 To fVal
      fTmp = fTmp * fInc
   Next fInc
   FactlF = fTmp

End Function

Public Function FactlL(lVal As Long) As Long

' Returns the factorial lVal! of long integer value lVal
' The highest argument value is 12,
' above an overflow results and the function returns -1

   Const Cl01 As Long = 1, Cl02 As Double = 2

   Dim lTmp As Long, lInc As Long

   If lVal > 12 Then
      lTmp = -1
   Else
      lTmp = Cl01
      For lInc = Cl01 To lVal
         lTmp = lTmp * lInc
      Next lInc
   End If
   FactlL = lTmp

End Function

Public Function FloorF(fVal As Double) As Double

' Returns the nearest integer not higher than fVal

   Const Cf01 As Double = 1#

   Dim fTmp As Double

   If fVal < 0 Then
      fTmp = Cf01 - Fix(fVal)
      fTmp = Fix(fVal + fTmp) - fTmp
   Else
      fTmp = Fix(fVal)
   End If

   FloorF = fTmp

End Function

Public Sub RngWbShRg(sWkS As String, sRng As String, aRng As Range, _
                     Optional sWkB As String = "", _
                     Optional aWkS As Worksheet, _
                     Optional aWkB As Workbook)

' Useful for quick testing, must be part of Lib0
' Particular case: sWkS = "", sRng contains sheet!Range

   If sWkB = "" Then
      Set aWkB = ThisWorkbook
   Else
      Set aWkB = Workbooks(sWkB)
   End If
   If sWkS = "" Then
      QRS_LibStr.StrSplit2 sRng, "!", sWkS, sRng
   End If
   Set aWkS = aWkB.Worksheets(sWkS)
   Set aRng = aWkS.Range(sRng)

End Sub

Public Function MaxOfL(lN As Long, _
                       Optional l1 As Long, Optional l2 As Long, _
                       Optional l3 As Long, Optional l4 As Long, _
                       Optional l5 As Long, Optional l6 As Long) As Long

' Returns the maximum of the first lN optiomal arguments

   Dim lm

   If Not lN < 1 Then lm = l1
   If Not lN < 2 Then If l2 > lm Then lm = l2
   If Not lN < 3 Then If l3 > lm Then lm = l3
   If Not lN < 4 Then If l4 > lm Then lm = l4
   If Not lN < 5 Then If l5 > lm Then lm = l5
   If Not lN < 6 Then If l6 > lm Then lm = l6

   MaxOfL = lm

End Function

Public Function MinOfL(lN As Long, _
                       Optional l1 As Long, Optional l2 As Long, _
                       Optional l3 As Long, Optional l4 As Long, _
                       Optional l5 As Long, Optional l6 As Long) As Long

' Returns the minimum of the first lN optiomal arguments

   Dim lm

   If Not lN < 1 Then lm = l1
   If Not lN < 2 Then If l2 < lm Then lm = l2
   If Not lN < 3 Then If l3 < lm Then lm = l3
   If Not lN < 4 Then If l4 < lm Then lm = l4
   If Not lN < 5 Then If l5 < lm Then lm = l5
   If Not lN < 6 Then If l6 < lm Then lm = l6

   MinOfL = lm

End Function

Public Function WithinL(lVal As Long, lLow As Long, lUpp As Long, _
                        Optional bInside As Boolean) As Boolean

' Returns true if lVal is within L
' if bInside is set, the condition is     lLow < lVal and lVal < lUpp
' else               the condition is not lLow > lVal  or lVal > lUpp

   If bInside Then
      WithinL = lLow < lVal And lVal < lUpp
   Else
      WithinL = Not (lLow > lVal Or lVal > lUpp)
   End If

End Function

Public Function PathApp(sPath As String, sFolder As String) As String

' Appends a folder (or file name) to a path
' If sFolder is empty, returns just sPath
' if sPath is empty, returns just sFolder
' if sPath is shorter than the separator, always inserts the separator
' If the path already contains the separator and sFolder is not empty,
' sFolder is just appended, else the separator is inserted

   Dim sSep As String, sRet As String
   Dim lLen As Long

   sSep = Application.PathSeparator

   If sFolder = "" Then                ' --- Folder empty
      sRet = sPath                     '     Return just path
   Else                                ' --- Folder not empty
      If sPath = "" Then               ' --- Path empty
         sRet = sFolder                '     Return just folder
      Else                             ' --- Path and folder not empty
         lLen = Len(sSep)
         If Len(sPath) < lLen Then     ' --- Path cannot contain separator
            sRet = sPath & sSep & sFolder
         Else                          ' --- Path may contain separator
            If Right(sPath, lLen) = sSep Then
               sRet = sPath & sFolder  ' --- Path ends by separator
            Else                       ' --- Path does not end by separator
               sRet = sPath & sSep & sFolder
            End If
         End If
      End If
   End If

   PathApp = sRet

End Function

Public Function BitGetMSB(lVal As Long, _
                          Optional lValMSB As Long = 0) As Long

' Returns the bit number (0-based) of the MSB in lVal
' Returns the corresponding value in lValMSB
' The intended B-tree search was too tedious

   Const Cl01 As Long = 1, Cl02 As Long = 2

   Dim lChk As Long, lBit As Long, lMSB As Long

   lBit = -Cl01                        ' --- Check 0: No bit set
   lChk = 0

   If lVal > lChk Then                 ' --- Value >0
      lValMSB = lChk                   '     Last MSB check value
      lBit = lMSB                      '     Test Bit 0
      lChk = Cl01                      '     MSB check value is 1

      While Not lVal < lChk            ' --- While MSB not found
         lValMSB = lChk                '     Last MSB check value
         lBit = lMSB                   '     Last bit tested
         lMSB = lMSB + Cl01            '     Next bit to test
         lChk = lChk * Cl02            '     Next MSB check value
      Wend
   End If

   BitGetMSB = lBit

End Function

