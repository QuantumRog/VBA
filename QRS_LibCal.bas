Attribute VB_Name = "QRS_LibCal"
Option Explicit

' Module :
' Project: any
' Purpose: Basic stateless calendar functions
' By     : QRS, Roger Strebel
' Date   : 14.03.2018
'          19.03.2018                  Weekday, WeekNbr added
' --- The public interface
'     DateAscension                    Date of ascension thursday    15.03.2018
'     DateCorpusChristi                Date of corpus christi thu    15.03.2018
'     DateEaster                       Date of easter sunday         15.03.2018
'     DateWhitsun                      Date if whitsun monday        15.03.2018
'     DateYMD                          Date elements                 15.03.2018
'     DayOffW                          Work day offset to date       19.03.2018
'     DayDST3                          March DST switch day          15.03.2018
'     DayDSTX                          October DST switch day        15.03.2018
'     Quarter                          Year quarter of a day         14.03.2018
'     WeekNbr                          ISO week number               19.03.2018
'     WkDNext                          Date of WkDay not before date 19.03.2018
'     WkDPrev                          Date of WkDay not after date  19.03.2018

Public Const MCf_LibCal_OffAsc As Double = 39
Public Const MCf_LibCal_OffWhi As Double = 50
Public Const MCf_LibCal_OffCpC As Double = 60

Public Const MCl_LibCal_WkDays As Long = 7

Public Function DateYMD(dDate As Date, Optional lY As Long = 0, _
                 Optional lM As Long = 0, Optional lD As Long = 0)

' Decomposes a date into year, month and date for convencence

   lY = Year(dDate)
   lM = Month(dDate)
   lD = Day(Date)

End Function

Public Function DateEaster(lY As Long) As Date

' Returns the date of easter Sunday according to Gauss' date rule
' Extracted from "Astronomical Formulae for Calculators" by Jean Meeus
' Avoids repeated calculation of the same year in a row

   Const Cl01 As Long = 1
   Const Cl02 As Long = 2
   Const Cl03 As Long = 3
   Const Cl04 As Long = 4
   Const Cl07 As Long = 7
   Const Cl08 As Long = 8
   Const Cl11 As Long = 11
   Const Cl15 As Long = 15
   Const Cl19 As Long = 19
   Const Cl25 As Long = 25
   Const Cl30 As Long = 30
   Const Cl31 As Long = 31
   Const Cl32 As Long = 32
   Const ClCC As Long = 100
   Const ClCX As Long = 114
   Const ClCQ As Long = 451

   Static dEaster As Date, lYrL As Long

   Dim lA As Long, lB As Long, lC As Long
   Dim lD As Long, lE As Long, lF As Long
   Dim lG As Long

   If Not lY = lYrL Then
      QRS_Lib0.DivModLon lY, Cl19, , lA
      QRS_Lib0.DivModLon lY, ClCC, lB, lC
      QRS_Lib0.DivModLon lB, Cl04, lD, lE
      QRS_Lib0.DivModLon lB + Cl08, Cl25, lF
      QRS_Lib0.DivModLon lB + Cl01 - lF, Cl03, lG
      QRS_Lib0.DivModLon lA * Cl19 + Cl15 + lB - lD - lG, Cl30, , lF
      QRS_Lib0.DivModLon lC, Cl04, lB, lD
      QRS_Lib0.DivModLon (lB + lE) * Cl02 + Cl32 - lF - lD, Cl07, , lC
      QRS_Lib0.DivModLon (lF - lC * Cl02) * Cl11 + lA, ClCQ, lD
      QRS_Lib0.DivModLon lC + lF + ClCX - lD * Cl07, Cl31, lA, lB
      dEaster = DateSerial(lY, lA, lB + Cl01)
      lYrL = lY
   End If
   DateEaster = dEaster

End Function

Public Function DateAscension(lYr As Long) As Date

' Returns the date of ascension (Thursday) in the given year

   DateAscension = DateEaster(lYr) + MCf_LibCal_OffAsc

End Function

Public Function DateWhitsun(lYr As Long) As Date

' Returns the date of whitsun (Monday) in the given year

   DateWhitsun = DateEaster(lYr) + MCf_LibCal_OffWhi

End Function

Public Function DateCorpChrist(lYr As Long) As Date

' Returns the date of corpus christi (Thursday) in the given year

   DateCorpChrist = DateEaster(lYr) + MCf_LibCal_OffCpC

End Function

Public Function DayOffW(dDate As Date, lNWorkdays As Long, _
                        Optional bBeforeWeekend As Long = False) As Date

' Returns a date lNWorkdays days after dDate
' if dDate is on a week-end, by default counts from next workday unless
' bBeforeWeekend is set, then counts from previous workday

   Const ClWDMon As Long = 1
   Const ClWDFri As Long = 5
   Const ClWDCnt As Long = 7           ' --- Week day count
   Const ClWECnt As Long = 2           ' --- Week end day count

   Dim dOffDay As Date
   Dim lOffDay As Long, lSgnOff As Long, lWkD As Long, lNWk As Long

   lSgnOff = Sgn(lNWorkdays)           ' --- May be zero
   If lSgnOff = 0 Then
      If bBeforeWeekend Then lSgnOff = -ClWDMon Else lSgnOff = ClWDMon
   End If

   dOffDay = dDate
   lWkD = Weekday(dOffDay, vbMonday)   ' --- Starting workday
   If lWkD > ClWDFri Then              ' --- falls on week-end
      If bBeforeWeekend Then           '     use previous workday?
         dOffDay = WkDPrev(dOffDay, ClWDFri)
      Else                             '     use next workday
         dOffDay = WkDNext(dOffDay, ClWDMon)
      End If
   End If
                                       ' --- Entire weeks and remaining days
   QRS_Lib0.DivModLon lNWorkdays, ClWDCnt, lNWk, lOffDay
   dOffDay = dOffDay + lNWk * ClWDCnt  ' --- Add entire weekds
                                       ' --- Remaining days
   lWkD = Weekday(dOffDay, vbMonday) + lOffDay
                                       ' --- Past labour days? -> skip week-end
   If lWkD < ClWDMon Or lWkD > ClWDFri Then
      lOffDay = lOffDay + ClWECnt * lSgnOff
   End If
   DayOffW = dOffDay + lOffDay

End Function

Public Function DayDST3(lYr As Long) As Date

' Returns the date of last sunday in march

   Const ClD As Long = 1
   Const ClM As Long = 4

   Dim d1 As Date                      ' --- auxiliary date

   d1 = DateSerial(lYr, ClM, ClD)      ' --- Day 1 of following month
   DayDST3 = d1 - Weekday(d1, vbMonday)

End Function

Public Function DayDSTX(lYr As Long) As Date

' Returns the date of last sunday in march

   Const ClD As Long = 1
   Const ClM As Long = 11

   Dim d1 As Date                      ' --- auxiliary date

   d1 = DateSerial(lYr, ClM, ClD)      ' --- Day 1 of following month
   DayDSTX = d1 - Weekday(d1, vbMonday)

End Function

Public Function IntNSat(dDayF As Date, dDayT As Date) As Long

' Returns the number of saturdays in the interval from dDayF to dDayT
' The number of days is the number of entire weeks, plus one
' if not( weekday(dDayF) > 6 and weekday(dDayT) < 6)

   Const Cl01 As Long = 1
   Const ClWD As Long = 6

   Dim lNWk As Long, lNDay As Long, lNRem As Long

   lNDay = Fix(dDayF) + Cl01 - Fix(dDayT)
   QRS_Lib0.DivModLon lNDay, MCl_LibCal_WkDays, lNWk, lNRem
'   If Weekday(dDayF, vbMonday) > ClWD And Weekday(dDayT, vbMonday) < ClWD Then
'      lNDay = lNDay + 1
'   End If
   IntNSat = lNDay

End Function

Public Function Quarter(dDay As Date) As Long

' Returns the quarter number of a day
' Avoids conversion to real value
'   1/3 = 0
'   2/3 = 1   5/3 = 2   8/3 = 3  11/3 = 4
'   4/3 = 1   7/3 = 2  10/3 = 3

   Const Cl01 As Long = 1, Cl03 As Long = 3

   Quarter = (Month(dDay) + Cl01) / Cl03

End Function

Public Function WeekNbr(dDay As Date, _
                        Optional lYear As Long = 0) As Long

' Returns the ISO calendar week number of the given date
' In particular cases, the year may be the previous year

   Const Cl01 As Long = 1, Cl07 As Long = 7
   Const ClWDThu As Long = 4
   Const CfOffTM As Long = -3#         ' --- Thursday - Monday

   Dim dDay11Y As Date, dThur1Y As Date, dMonPrv As Date
   Dim lWkOfYr As Long

   lYear = Year(dDay)
   dDay11Y = DateSerial(lYear, 1, 1)   ' --- 1st of january in year
   dThur1Y = WkDNext(dDay11Y, ClWDThu) '     1st thursday in year
   dMonPrv = dThur1Y + CfOffTM
   QRS_Lib0.DivModLon CLng(dDay - dMonPrv), Cl07, lWkOfYr
   lWkOfYr = lWkOfYr + Cl01
   If lWkOfYr < Cl01 Then
      lWkOfYr = 53
      lYear = lYear - Cl01
   End If
   WeekNbr = lWkOfYr

End Function

Public Function WkDNext(dDay As Date, lWkD As Long) As Date

' Returns the date of the specific weekday on or after dDay

   Const Cf07 As Double = 7#

   Dim fWD0 As Double

   fWD0 = lWkD - Weekday(dDay, vbMonday)
   If fWD0 < 0 Then fWD0 = fWD0 + Cf07
   WkDNext = dDay + fWD0

End Function

Public Function WkDPrev(dDay As Date, lWkD As Long) As Date

' Returns the date of the specific weekday before or on dDay

   Const Cf07 As Double = 7#

   Dim fWD0 As Double

   fWD0 = Weekday(dDay, vbMonday) - lWkD
   If fWD0 < 0 Then fWD0 = fWD0 + Cf07
   WkDPrev = dDay - fWD0

End Function
