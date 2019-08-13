VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QRS_TimeIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Module :
' Purpose: Realize an Android-style time input widget
'          in an Excel range
'          The Android-Style time input widget has the
'          shape of an analog clock dial. This version
'          has just 12 hour dials that show either day
'          or night hours
'                Day              Night
'             11 12 13          23  0  1
'          10          14    22           2
'           9          15    21           3
'           8          16    20           4
'              7 18 17          19  6  5
'          The first day hour can be set
'          The dials are toggled by clicking on the center
'          The minute dials allow direct selection of 1/4h
'          and variable steps up (1 to 5 min)
'             +1  0 +5
'             45    15
'             -1 30 -5
'          This dial looks good at column widths of 2.33 twips
' Use    : This module can be used by two functions of the kind
'--- Copy from line below, paste to a VBA module and uncomment
'Private McTI as QRS_TimeIn
'
'Sub Dial_Init
'
'   Dim aRg as Range
'
'   If McTI Is Nothing Then Set McTI = New QRS_TimeIn
'
'   QRS_Lib0.RngWbShRg "<Sheet>", "<Cell_TimeIn>", aRg
'   Set McTI.RangeTime_In = aRg
'   QRS_Lib0.RngWbShRg "<Sheet>", "<CellTopLeft>", aRg
'   Set McTI.RangeTopLeft = aRg         ' --- Initialize
'   McTI.Draw_Dial
'   McTI.Show_Time
'
'End Sub
'
'Sub DialClick(aRange As Range)
'
'' Called from a worksheet Selection-Change event
'
'   If Not McTI Is Nothing Then
'      McTI.Click aRange
'      If McTI.Is_Canceled Or McTI.Is_OK Then Set McTI = Nothing
'   End If
'
'End Sub
'--- Copy till line above
'--- Copy from line below, paste to a worksheet code and uncomment
'Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'
'   Mod_TimeIn.DialClick Target
'
'End Sub
' --- Copy till line above
' By     : Roger Strebel
' Date   : 01.11.2018                  First tryouts
'          20.01.2019                  More ideas, some order
'          21.01.2019                  Implemented and tested        V0.97
'          23.01.2019                  Input-output cell added
'
' --- The public interface
' Properties
'    Day_Hour1       (     let)        Set first day hour
'    RangeTopLeft    (     set)        Set top left dial cell
'    RangeTime_In    (     set)        Set input/output cell
'    TimeFmt         (     let)        Set time display format
'    TimeHM          (get, let)        Time (hour and minute)
' Methods
'    ClearDial                         Clear dial output range
'    Day_Night                         Toggle day-night dial
'    Draw_Dial                         Draw the dial
'    Show_Time                         Display time currently set
' --- The private sphere
'    Class_Initialize                  Set up internal structures
'    Get_Hour                          Get hour from cell
'    Get_Time                          Get date type time
'    Res_Time                          Reset time alterations
'    Set_Hours                         Set Day/Night hour labels
'    Set_Labels                        Set 2D label table

Private MlDayHr1 As Long               ' --- Day Hour 1
Private MlNight As Long                '     0: Day, 1: Night

Private MlHour As Long, MlHr0 As Long
Private MlMinute As Long, MlMn0 As Long

Private MbOK As Boolean
Private MbCancel As Boolean
Private MbUsy As Boolean

Private MsLbl() As String
Private MvLbl() As Variant

Private MaKXY As QRS_LUT_XY

Private MsFmt As String
Private MaRngIO As Range               ' --- Time input/output cell
Private MaRngTL As Range               ' --- Top left cell of dial range
Private MlCol0 As Long                 '     Column number of MaRngTL
Private MlRow0 As Long                 '     Row number of MaRngTL
Private MlColW() As Long               '     Column widths
Private MlDXY As Long                  '     X and Y difference 1-based
Private MlOXY As Long                  '     X and Y difference 0-based

Public Property Let TimeHM(dTime As Date)

   MlHour = Hour(dTime)                ' --- Round minutes
   MlMinute = (dTime - MlHour / 24) * 1440

   MlHr0 = MlHour
   MlMn0 = MlMinute

End Property

Public Property Get TimeHM() As Date

   TimeHM = Get_Time()

End Property

Public Property Let TimeFmt(sFmtTime As String)

   If sFmtTime = MsFmt Then Exit Property

   MsFmt = sFmtTime
   If Not MaRngTL Is Nothing Then MaRngTL.Offset(5, 0).NumberFormat = MsFmt

End Property

Public Property Set RangeTopLeft(aRng As Range)

   Set MaRngTL = aRng
   MlCol0 = MaRngTL.Column
   MlRow0 = MaRngTL.Row

End Property

Public Property Set RangeTime_In(aRng As Range)

   Dim v

   Set MaRngIO = aRng                  ' --- memorize io range
   v = aRng.Value                      ' --- Auto-set time
   If Not v = "" Then Me.TimeHM = v    '     if io range not empty

   MbOK = False                        ' --- clear flags
   MbCancel = MbOK

End Property

Public Property Get Is_Canceled() As Boolean

   Is_Canceled = MbCancel

End Property

Public Property Get Is_OK() As Boolean

   Is_OK = MbOK

End Property

Public Sub Click(aRng As Range)

   Const Cl60 As Long = 60

   Dim lX As Long, lY As Long, lK As Long
   Dim bShowTime As Boolean

   If MbUsy Then Exit Sub

   MbUsy = True

   lX = aRng.Row - (MlRow0 + MlOXY)
   lY = aRng.Column - (MlCol0 + MlOXY)
   lK = MaKXY.GetKey(lX, lY)

   bShowTime = True
   Select Case lK
   Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11:
      MlHour = Get_Hour(lK)
   Case 15, 30, 45:
      MlMinute = lK
   Case 60:
      MlMinute = 0
   Case 14:                            ' --- +5 minutes
      MlMinute = MlMinute + 5
      If MlMinute > Cl60 Then MlMinute = MlMinute - Cl60
      MaRngTL.Activate                 '     Select other cell
   Case 16:                            ' --- -1 minute
      MlMinute = MlMinute - 1
      If MlMinute < 0 Then MlMinute = MlMinute + Cl60
      MaRngTL.Activate                 '     Select other cell
   Case 44:                            ' --- +1 minute
      MlMinute = MlMinute + 1
      If MlMinute > Cl60 Then MlMinute = MlMinute - Cl60
      MaRngTL.Activate                 '     Select other cell
   Case 46:                            ' --- -5 minutes
      MlMinute = MlMinute - 5
      If MlMinute < 0 Then MlMinute = MlMinute + Cl60
      MaRngTL.Activate                 '     Select other cell
   Case 24:
      MlNight = 1 - MlNight            '     Toggle night/day
      Set_Hours
      Draw_Dial
      bShowTime = False
   Case 17:                            '     Cancel
      MbCancel = True                  ' --- Mark canceled flag
      Res_Time                         '     reset time
      ClearDial                        '     clear dial range
   Case 18:                            '     OK
      MbOK = True
      ClearDial
      bShowTime = False
   Case Else
      bShowTime = False
   End Select

   If bShowTime Then Show_Time

   MbUsy = False

End Sub

Public Sub Day_Night()

' --- Toggle day and night

   MlNight = 1 - MlNight

End Sub

Public Sub Draw_Dial()

' --- Draw dial (labels in MvLbl)
'     Color indices used:
'        Day    :  6
'        Night  : 23
'        Quarter: 15
'        Cancel : 40   alternatively 38
'        OK     : 35
'        Rest    : -4142

   Const ClCndxD As Long = 6, ClCndxN = 23, ClCndxQ = 15
   Const ClCndxC As Long = 40, ClCndxO = 35

   Dim lIObj As Long, lX As Long, lY As Long, lK As Long
   Dim lCndxH As Long, lCNdxD As Long
   
   QRS_LibLst.LstAllocL MlColW(), 5    ' --- Column width memory
   For lIObj = 0 To 4                  '     Back up column widths
      MlColW(lIObj + 1) = MaRngTL.Offset(0, lIObj).ColumnWidth
   Next lIObj                          ' --- Small columns
   Range(MaRngTL, MaRngTL.Offset(0, 4)).ColumnWidth = 2.33

   QRS_LibXL.RngPutArrV MaRngTL, MvLbl()

   If MlNight = 0 Then
      lCndxH = ClCndxD
      lCNdxD = ClCndxN
   Else
      lCndxH = ClCndxN
      lCNdxD = ClCndxD
   End If
   MaKXY.GetRow 22, lY, lX, lK         ' --- Cancel
   With MaRngTL.Offset(lY + MlOXY, lX + MlOXY)
      .Interior.ColorIndex = ClCndxC
      .HorizontalAlignment = xlCenter
   End With
   MaKXY.GetRow 23, lY, lX, lK         ' --- OK
   With MaRngTL.Offset(lY + MlOXY, lX + MlOXY)
      .Interior.ColorIndex = ClCndxO
      .HorizontalAlignment = xlCenter
   End With
   MaKXY.GetRow 21, lY, lX, lK         ' --- Day/Night
   With MaRngTL.Offset(lY + MlOXY, lX + MlOXY)
      .Interior.ColorIndex = lCNdxD
      .HorizontalAlignment = xlCenter
   End With
   For lIObj = 1 To 12                 ' --- Hours
      MaKXY.GetRow lIObj, lY, lX, lK
      MaRngTL.Offset(lY + MlOXY, lX + MlOXY).Interior.ColorIndex = lCndxH
   Next lIObj
   For lIObj = 13 To 16                ' --- Quarter hours
      MaKXY.GetRow lIObj, lY, lX, lK
      MaRngTL.Offset(lY + MlOXY, lX + MlOXY).Interior.ColorIndex = ClCndxQ
   Next lIObj

End Sub

Public Sub ClearDial()

' Clears the dial at the specified range

   Dim lI As Long

   With Range(MaRngTL, MaRngTL.Offset(4, 4))
      .ClearFormats
      .ClearContents
   End With

   For lI = 0 To 4
      MaRngTL.Offset(0, lI).ColumnWidth = MlColW(lI + 1)
   Next lI

End Sub

Public Sub Show_Time()

' Shows the time in the IO cell

   If MaRngTL Is Nothing Then Exit Sub

   MaRngIO.Value = Get_Time

End Sub

Public Property Let Day_Hour1(lDay_Hour1 As Long)

' --- Set the first "day" hour
' --- Redraw dial if output range specified

                                       ' --- No new hour
   If lDay_Hour1 = MlDayHr1 Then Exit Property

   MlDayHr1 = lDay_Hour1
                                       ' --- Not yet ready to draw
   If MaRngTL Is Nothing Then Exit Property

   Set_Hours
   Draw_Dial

End Property

Private Function Get_Hour(lK As Long) As Long

' Return the hour depending on the clicked hour cell
' If lK < Day hour 1:
'    If MlNight = 0 (daytime dial): Add+12 h
'    If MlNight = 1 (night time dial) Add 0h
' Else:
'    If MlNight = 0 (daytime dial): Add 0h
'    If MlNight = 1 (night time dial) Add+12h
' 12 hours must be added if:
'     MlNight = 0 And lK < MlDayHour1 Or MlNight = 1 And Not l
' Works OK

   Dim lH As Long

   If MlNight = 0 Xor lK < MlDayHr1 Then lH = 0 Else lH = 12

   Get_Hour = lK + lH

End Function

Private Function Get_Time() As Date

' Return time as date type variable from minutes and hours

   Const CfMPH As Double = 60#
   Const CfHPD As Double = 24#

   Get_Time = (MlMinute / CfMPH + MlHour) / CfHPD

End Function

Private Sub Res_Time()

' --- Resets time to last value got from property Let TimeHM

   MlHour = MlHr0
   MlMinute = MlMn0

End Sub

Private Sub Set_Hours()

' --- Sets the hour labels according to day/night
' --- For the day:   Hours < MlDayHr1 are incremented by 12
' --- For the night: Hours >= MlDayHr1 are decremented by 12

   Const ClIOH As Long = -1
   Const ClNHr As Long = 12

   Dim lI As Long, lK As Long, lX As Long, lY As Long, lO As Long

   If MlNight = 0 Then lO = ClNHr Else lO = 0
   For lI = 1 To ClNHr
      MaKXY.GetRow lI, lY, lX, lK
      MvLbl(lY + MlDXY, lX + MlDXY) = lI + lO + ClIOH
      If lI = MlDayHr1 Then lO = ClNHr - lO
   Next lI

End Sub

Private Sub Set_Labels()

' --- Fill the label table from the 1D label list at KXY offsets

   Dim lI As Long, lK As Long, lX As Long, lY As Long

   For lI = 1 To MaKXY.Count
      MaKXY.GetRow lI, lY, lX, lK
      MvLbl(lY + MlDXY, lX + MlDXY) = MsLbl(lI)
   Next lI

End Sub

Private Sub Class_Initialize()

' Initialize the lists and tables to manage the dial
' Indices:  1...12: Hours
'          15...60: Quarter-hours  15,30,45,00
'          14     : +5 min
'          46     : -5 min
'          44     : +1 min
'          16     : -1 min
'          24     : Day/Night
'          17     : Quit (X)


   Dim sIni As String
   Dim lIni() As Long

                                       ' --- Index, offsets from center
                                       ' --- Hour dial
   sIni = " 0,-2, 0/ 1,-2, 1/ 2,-1, 2/ 3, 0, 2/ 4, 1, 2/ 5, 2, 1/" _
        & " 6, 2, 0/ 7, 2,-1/ 8, 1,-2/ 9, 0,-2/10,-1,-2/11,-2,-1/" _
                                       ' --- Quarter-hour dial
   sIni = sIni & "15, 0, 1/30, 1, 0/45, 0,-1/60,-1, 0/"
                                       ' --- Inc, dec, 5 min, 1 min
   sIni = sIni & "14,-1, 1/46,-1,-1/44, 1,-1/16, 1, 1/"
                                       ' --- Day/night, Cancel(X), OK (/)
   sIni = sIni & "24, 0 ,0/17,-2, 2/18, 2, 2"

   QRS_LibStr.ArrStrLon sIni, lIni()   ' --- Index and offsets
   Set MaKXY = New QRS_LUT_XY
   MaKXY.InitFromArrColL lIni()

   MlDayHr1 = 7                        ' --- Day hour 1

   MsFmt = "HH:MM:SS"                  ' --- Default time format

   MlOXY = 2
   MlDXY = MlOXY + 1

   sIni = "0,1,2,3,4,5,6,7,8,9,10,11,15,30,45,0,'+5,-5,'+1,-1,D,X,/"
   QRS_LibStr.LstStrLst sIni, MsLbl(), ","

   QRS_LibArr.ArrAllocV MvLbl(), 5, 5
   Set_Labels                          ' --- Init default labels
   Set_Hours

End Sub
