Attribute VB_Name = "StatusBar"
' Copyright (c) 2025 Alberto Lopez Yus
' Licensed under the Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)
' See the LICENSE file for details.

Option Explicit

Public strUpdate As String
Public dblStatusPrg As Double

Public Sub UpdateProgressBar(dblPrg As Double)
    dblPrg = Round(dblPrg, 1)
    
    If dblPrg >= dblStatusPrg Then
        frmStatusBar.frameStatus.Caption = format(dblPrg, "0%")
        frmStatusBar.lblStatus.Width = dblPrg * (frmStatusBar.frameStatus.Width - 5)
        
        dblStatusPrg = dblPrg + 0.1
        DoEvents
        
        If dblStatusPrg > 1 Then
            frmStatusBar.lblStatus.BackColor = RGB(0, 204, 0)
            Application.Wait (Now + TimeValue("0:00:01") * 3 / 4)
            Unload frmStatusBar
        End If
    End If
End Sub

Public Sub ShowStatusBar()
    With frmStatusBar
      .StartUpPosition = 0
      .Width = 340
      .Height = 115
      .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
      .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
      .Show
    End With
End Sub

