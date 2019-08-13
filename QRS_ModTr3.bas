Attribute VB_Name = "QRS_ModTr3"
Option Explicit

Private Type tCfg
   fPUp As Double                      ' --- Probability level for  up  move
   fPDn As Double                      ' --- Probability level for down move
   fSUp As Double
   fSDn As Double
   fV00 As Double
   lLen As Long                        ' --- Sequence length
   lCnt As Long                        ' --- Sequence count in ensemble
End Type

Private MCfg As tCfg

Public Sub CfgSet(v(), lColCfg As Long, _
                       lRowPUp As Long, lRowPDn As Long, _
                       lRowSUp As Long, lRowSDn As Long, _
                       lRowV00 As Long, _
                       lRowLen As Long, lRowCnt As Long)

   Const Cf01 As Double = 1#

   QRS_LibArr.XtrEleVarAbs v(), lRowPUp, lColCfg, MCfg.fPUp, _
                                lRowPDn, lColCfg, MCfg.fPDn, _
                                lRowSUp, lColCfg, MCfg.fSUp, _
                                lRowSDn, lColCfg, MCfg.fSDn
   QRS_LibArr.XtrEleVarAbs v(), lRowV00, lColCfg, MCfg.fV00, _
                                lRowLen, lColCfg, MCfg.lLen, _
                                lRowCnt, lColCfg, MCfg.lCnt
   With MCfg                           ' --- Convert probability to level
      .fPUp = 1 - .fPUp
   End With

End Sub

Public Sub GenTr3Seq(fSeq() As Double)

' Generate one ternary tree sequence
' from the MCfg settings

   Dim lI As Long
   Dim fP As Double, fS As Double, fV As Double

   With MCfg
      QRS_LibLst.LstAllocF fSeq(), .lLen

      fV = .fV00
      For lI = 1 To .lLen
         fSeq(lI) = fV
         fP = Rnd(CSng(fP))
         If fP < .fPDn Then
            fS = .fSDn
         Else
            If fP > .fPUp Then fS = .fSUp Else fS = 0
         End If
         fV = fV + fS
      Next lI
   End With

End Sub

Public Sub GenTr3Ens(fEns() As Double)

   Dim fSeq() As Double

   Dim lI As Long

   With MCfg
      QRS_LibArr.ArrAllocF fEns(), .lLen, .lCnt
   
      For lI = 1 To .lCnt
         GenTr3Seq fSeq()
         QRS_LibA2L.PutColDblDbl fSeq(), fEns(), lI
      Next lI
   End With

End Sub
