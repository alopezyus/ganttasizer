Attribute VB_Name = "ShpObject"
' Copyright (c) 2025 Alberto Lopez Yus
' Licensed under the Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)
' See the LICENSE file for details.

Dim ShpEvents As stdShapeEvents

Public Sub startbarEvents()
  Set ShpEvents = New stdShapeEvents
  Call ShpEvents.HookSheet(ActiveSheet)
End Sub

Public Sub stopBarEvents()
  Set ShpEvents = Nothing
End Sub

