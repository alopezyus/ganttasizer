Attribute VB_Name = "Shortcuts"
' Copyright (c) 2025 Alberto Lopez Yus
' Licensed under the Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)
' See the LICENSE file for details.

Option Explicit

Public Sub CreateShortcuts()
    Application.OnKey "+%{G}", "ACT_CreateTemplate"
    Application.OnKey "+%{L}", "openColSelForm"
    Application.OnKey "+%{T}", "ACT_CreateCalendar"
    Application.OnKey "+%{C}", "ACT_CreateChart_shortcut"
    If intEdition = 0 Then
        Application.OnKey "+%{W}", "ACT_FormatWBS" 'Disabled in Free Edition and Pro Edition
        Application.OnKey "+%{S}", "ACT_CalculateNetwork_shortcut" 'Disabled in Free Edition and Pro Edition
        Application.OnKey "+%{U}", "ACT_DistributeUnits" 'Disabled in Free Edition and Pro Edition
    End If
    If Not intEdition = 1 Then
        Application.OnKey "+%{R}", "MngRel" 'Disabled in Free Edition
        Application.OnKey "+%{N}", "ACT_CreateConnectors" 'Disabled in Free Edition
    End If
End Sub


Public Sub DeleteShortcuts()
    Application.OnKey "+%{G}"
    Application.OnKey "+%{L}"
    Application.OnKey "+%{T}"
    Application.OnKey "+%{C}"
    If intEdition = 0 Then
        Application.OnKey "+%{W}" 'Disabled in Free Edition and Pro Edition
        Application.OnKey "+%{S}" 'Disabled in Free Edition and Pro Edition
        Application.OnKey "+%{U}" 'Disabled in Free Edition and Pro Edition
    End If
    If Not intEdition = 1 Then
        Application.OnKey "+%{R}" 'Disabled in Free Edition
        Application.OnKey "+%{N}" 'Disabled in Free Edition
    End If
End Sub
