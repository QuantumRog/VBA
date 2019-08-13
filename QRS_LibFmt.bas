Attribute VB_Name = "QRS_LibFmt"
Option Explicit

' Module : QRS_LibFmt
' Project: any
' Purpose: Utility VBA routines for locale independent string formatting
'
' By     : QRS, Roger Strebel
' Date   : 13.02.2019                  StrFmtVBA works
'
' --- The public interface
'     StrFmtVBA                        Format date using VBA Format  13.02.2019
'     StrFmtQRS                        Format date using QRS Format

Public Function StrFmtVBA(sStr As String, dDat As Date, lInt As Long)

' Parses a string containing formatting tags in the form of
' <FFF> by extracting these tags and passing them to the VBA
' format statement  format   element format
'                    d        day     #0
'                    dd       day     00
'                    ddd      day     Mo.
'                    dddd     day     Montag
'                    m        month   #0
'                    mm       month   00
'                    mmm      month   Feb.
'                    mmmm     month   Februar
'                    w        weekday 1: Sun - 7:Sat
'                    ww       week-nr week-number, week 1 contains jan. 1
'                    y    day of year ##0
'                    yy       year    00
'                    yyyy     year    0000
'                    #        number  digit
'                    0        number  digit or leading zero

   Const CsBeg As String = "<", CsEnd As String = ">"

   Dim sTmp As String, sFmt As String
   Dim sTag As String, sDat As String
   Dim lPos As Long

   sTmp = sStr
   QRS_LibStr.StrFld sTmp, CsBeg, CsEnd, lPos, sFmt, , lPos
   While lPos > 1
      If InStr(1, sFmt, "#") > 0 Or InStr(1, sFmt, "0") > 0 Then
         sDat = Format(lInt, sFmt)
      Else
         sDat = Format(dDat, sFmt)
      End If
      sTag = CsBeg & sFmt & CsEnd
      sTmp = QRS_LibStr.StrRpl(sTmp, sTag, sDat)
      lPos = lPos + Len(sDat) - Len(sTag)
      QRS_LibStr.StrFld sTmp, CsBeg, CsEnd, lPos, sFmt, , lPos
   Wend

   StrFmtVBA = sTmp

End Function
