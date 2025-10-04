VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStatusBar 
   Caption         =   "Status Bar"
   ClientHeight    =   1530
   ClientLeft      =   -140
   ClientTop       =   -650
   ClientWidth     =   5350
   OleObjectBlob   =   "frmStatusBar.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmStatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright (c) 2025 Alberto Lopez Yus
' Licensed under the Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)
' See the LICENSE file for details.


Private Sub UserForm_Activate()
    frmStatusBar.lblStatus.Width = 0
    frmStatusBar.frameStatus.Caption = format(0, "0%")
    dblStatusPrg = 0.1
    'frmStatusBar.Caption = "Ganttasizer"
    
    Select Case strUpdate
    Case "Template"
        frmStatusBar.lblUpdate.Caption = "Creating Headers..."
        StatusBar_CreateTemplate
    Case "CopyWs"
        frmStatusBar.lblUpdate.Caption = "Copying Worksheet..."
        StatusBar_CopyWs
    Case "ClearCalendar"
        frmStatusBar.lblUpdate.Caption = "Deleting Calendar..."
        StatusBar_ClearCalendar
    Case "CreateCalendar"
        frmStatusBar.lblUpdate.Caption = "Creating Calendar..."
        StatusBar_CreateCalendar
    Case "CreateChart"
        frmStatusBar.lblUpdate.Caption = "Creating Chart..."
        StatusBar_CreateChart
    Case "ClearChart"
        frmStatusBar.lblUpdate.Caption = "Deleting Chart..."
        StatusBar_ClearChart
    Case "FilterShapes"
        frmStatusBar.lblUpdate.Caption = "Checking Filtered Shapes..."
        StatusBar_FilterShapes
    Case "FormatWBS"
        frmStatusBar.lblUpdate.Caption = "Formatting WBS..."
        StatusBar_FormatWBS
    Case "FormatWBS"
        frmStatusBar.lblUpdate.Caption = "Formatting WBS..."
        StatusBar_FormatWBS
    Case "ClearConnectors"
        frmStatusBar.lblUpdate.Caption = "Deleting Connectors..."
        StatusBar_ClearConnectors
    Case "CreateConnectors"
        frmStatusBar.lblUpdate.Caption = "Creating Connectors..."
        StatusBar_CreateConnectors
    Case "DistributeUnits"
        frmStatusBar.lblUpdate.Caption = "Distributing Remaining Units..."
        StatusBar_DistributeUnits
    Case "CalculateSchedule"
        frmStatusBar.lblUpdate.Caption = "Scheduling..."
        CalculateSchedule
    End Select
End Sub
