Attribute VB_Name = "Info"
' Copyright (c) 2025 Alberto Lopez Yus
' Licensed under the Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)
' See the LICENSE file for details.

Option Explicit

Public Sub CustomMsgBox(msg As String, Optional btn As Integer = vbOKOnly + vbInformation, Optional title As String = "Ganttasizer", Optional error As Boolean = False)
    MsgBox msg, btn, title
End Sub

Public Sub InfoGanttasizer()
    Dim strInfo As String
    
    'Change according to Edition
    strInfo = "GANTTASIZER" & vbNewLine & vbNewLine & _
                "Designed and Developed by Alberto López Yus" & vbNewLine & _
                "mail: alopezyus@gmail.com"
                'strInfo & vbNewLine & vbNewLine & _

    MsgBox strInfo, vbOKOnly, "About Ganttasizer"
End Sub

Public Sub LicenseFile()
    Dim strInfo As String
    
    Select Case intEdition
    Case 0
        strInfo = "Master Edition"
    Case 1
        strInfo = "Free Edition"
    Case 2
        strInfo = "Pro Edition"
    End Select
    
    'Change according to Edition
    strInfo = "Copyright (c) 2025 Alberto Lopez Yus" & vbNewLine & vbNewLine & _
              "Licensed under the Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)" & vbNewLine & vbNewLine & _
              "See the LICENSE file for details."

    MsgBox strInfo, vbOKOnly, "License Information"
End Sub


Public Sub HelpFile()
    Dim ws As Worksheet
    Dim strInfo1, strInfo2, strInfo3, strInfo4, strInfo5, strWsName As String
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    strWsName = "ganttasizerHelp"
    'Contar número de hojas de configuración existentes
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name = strWsName Then
            ws.Delete
            Exit For
        End If
    Next
    
    'Crear hoja configuracion
    Worksheets.Add.Name = strWsName
    Set ws = ActiveWorkbook.Worksheets(strWsName)


    strInfo1 = "HELP GUIDE" & vbNewLine
    strInfo2 = "1. DISCLAIMER: " & vbNewLine & _
                "The objective of this help guide is to give some useful tips for the operation of Ganttasizer but it is not intended to be neither a comprehensive list of options nor a guide on how to create Gantt charts. " & vbNewLine & vbNewLine & _
                "2. INTRODUCTION: " & vbNewLine & _
                "Ganttasizer is an Excel Add-in for creating project timelines and Gantt charts from any Excel native data. It provides several customization options to make your Gantt charts look professional and unique. Moreover, it works seamlessly with Excel and keeps all its native functions available." & vbNewLine & _
                "Ganttasizer can also manage your WBS structure, add and draw relationships between your activities and use them to calculate your project network, track progress and perform workload time distribution. Discover the rest of its features." & vbNewLine & vbNewLine & _
                "3. ABOUT SHAPES POSITIONING IN THE GANTT CHART" & vbNewLine & _
                "If you are experiencing problems with shapes positioning in the calendar scale, please check if you are working on an screen different from your main one." & vbNewLine & _
                "If that's the case, you will need to adjust your display settings to the same resolution and scaling on both screens: the main one and the working one." & vbNewLine & vbNewLine & _
                "4. GANTTASIZER TAB IN EXCEL RIBBON" & vbNewLine & _
                "All the options included in the Excel Ribbon have their own self explanatory tips. Therefore, those will not be repeated in this guide." & vbNewLine
    strInfo3 = "5. SYSTEM COLUMNS IN WORKSHEETS" & vbNewLine & _
                "All the column headers added when the 'Add Headers' button is pressed contain information used in the Ganttasizer environment. " & _
                "This system columns cannot be deleted but they can be hidden, moved or renamed if required by the user. New columns and rows can also be added." & vbNewLine & _
                "There are eight columns for activity-level setup, eighteen columns with project information and a separation column between the information table and the chart. " & vbNewLine & vbNewLine & _
                "This guide is focused on the activity-level setup columns." & vbNewLine & _
                "   *act/mil style: Four setups supported by this column for activity shapes." & vbNewLine & _
                "       -Seventeen options for bars and milestones shape styles in the drop down cell. Options 1-10 are only supported with bars while options 11-17 are only supported by milestones." & vbNewLine & _
                "       -An additional option in the drop down cell: NO. It is only applicable for activities that are part of a timeline and it allows to hide them in the timeline but never in the detailed representation." & vbNewLine & _
                "       -An additional option in the drop down cell: WINDOW. Turn activities into windows in you schedule." & vbNewLine & _
                "       -Fill color for remaining bars and milestones. Use the regular fill color picker built in Excel." & vbNewLine & _
                "   *shape height: Ten options to adjust bars and milestones size as a percentage of the row height. The height of windows must be introduced as the number of rows the window will cover." & vbNewLine & _
                "   *connect style: Three setups supported by this column for connectors." & vbNewLine & _
                "       -Six options for line styles in the drop down cell." & vbNewLine & _
                "       -An additional option in the drop down cell: NO. It prevents the connector with the predecessor activity to be drawn." & vbNewLine & _
                "       -Line color for connectors. Use the regular fill color picker built in Excel."
    strInfo4 = "    *label pos: Two setups supported by this column for activity labels." & vbNewLine & _
                "       -Nine positioning options for labels in the drop down cell only applicable for timelines. The positioning options are the result of combining two parameters." & vbNewLine & _
                "           Height: Three levels defined as per shape vertical positionning: 0 (same level), 1 (one level above shape), 2 (two levels above shape)." & vbNewLine & _
                "           Allignment: Three positions defined as per shape horizontal positionning: L (allingment to the left of the shape), M (allingment to the middle of the shape), R (allingment to the right of the shape)." & vbNewLine & _
                "       -An additional option in the drop down cell: NO. It prevents the label to be displayed both in the regular representation of the Gantt chart in timelines." & vbNewLine & _
                "   *timeline mode: Three different types of timeline can be created using this setup. It must be selected only in the line where the timeline will be drawn." & vbNewLine & _
                "       SUM: Creates an only summary bar for the whole span of time of the activities in the timeline." & vbNewLine & _
                "       MIL: Creates a milestone for each activity finish date and a bar for the whole span of time of the activities in the timeline. The milestones, labels and the summary bar can be edited with the rest of activity-level setups." & vbNewLine & _
                "       ACT: Creates an activity for each activity in the timeline. The activities and labels can be edited with the rest of activity-level setups." & vbNewLine & _
                "   *timeline code: Free code used to identify the activities that are part of a timeline. The timeline row and the activities that are meant to be part of the timeline must be assigned with the same code." & vbNewLine & _
                "   *schedule mode: ALAP constraint, four constraints for Start date and four constraints for Finish date can be selected. Also, two network calculations exceptions can be selected:" & vbNewLine & _
                "       NO: This activity will not take any part in network calculation." & vbNewLine & _
                "       MANUAL: This activity will not be calculated according to its predecessors but it will affect the calculation of its successors. Its dates are contrained to the data introduce prior to the calculation start." & vbNewLine & _
                "   *units distrib curve: Four predifined distribution curves can be selected: linear, s-curve, front loaded, back loaded." & vbNewLine
    strInfo5 = "There are also some interesting tips about project information columns to point out ." & vbNewLine & _
                "   *ACTIVITY ID / DESCRIPTION: Fill either or both of these columns to define an activity. If, for any row below the headers, neither column is filled, the activities list is deemed to be closed on that row." & vbNewLine & _
                "                               IMPORTANT: Activity ID must be unique and not contain spaces. To define a relationship between two activities, both of them must have Activity ID." & vbNewLine & _
                "   *WBS: WBS levels must be separted by dots (.). You can setup the WBS level color using the regular fill color picker built in Excel on this column." & vbNewLine & _
                "   *TOTAL / REMAINING DURATION: Total Duration is always a calculated column. Remaining Duration is a calculated column when drawing the Gantt chart but a user defined value when scheduling." & vbNewLine & _
                "   *START / FINISH DATE: User defined values are used when drawing the Gantt Chart but dates are calculated using Remaining Duration and Predecessors when scheduling." & vbNewLine & _
                "   *ACTUAL START / FINISH DATE: Actual Dates user defined values are only used when scheduling an they are never used for drawing the Gantt Chart." & vbNewLine & _
                "   *RESUME DATE: Only applicable when the activity is started before the cutoff date but still in progress after the cut off date." & vbNewLine & _
                "       User defined values are used when drawing the Gantt Chart but dates are calculated using Remaining Duration and Predecessors when scheduling." & vbNewLine & _
                "   *CONSTRAINT DATE: Date to be applied in the calculations when a Start or Finish date constraint is selected in the schedule mode column." & vbNewLine & _
                "   *BUDGET UNITS: User defined values are used to weigh Progress % values in summaries." & vbNewLine & _
                "   *REMAINING UNITS: User defined values are used for units distribution."
  
    ws.Range("A1") = strInfo1
    ws.Range("A2") = strInfo2
    ws.Range("A3") = strInfo3
    ws.Range("A4") = strInfo4
    ws.Range("A5") = strInfo5
    
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    With ws.Range("A1:A5")
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
        .ColumnWidth = 200
        .RowHeight = 350
        .EntireRow.AutoFit
    End With
    ActiveWindow.DisplayGridlines = False
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub


