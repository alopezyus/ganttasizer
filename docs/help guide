# HELP GUIDE

---

## 1. DISCLAIMER  
The objective of this help guide is to give some useful tips for the operation of **Ganttasizer** but it is not intended to be either a comprehensive list of options nor a guide on how to create Gantt charts.

---

## 2. INTRODUCTION  
**Ganttasizer** is an Excel Add-in for creating project timelines and Gantt charts from any Excel native data. It provides several customization options to make your Gantt charts look professional and unique. Moreover, it works seamlessly with Excel and keeps all its native functions available.

Ganttasizer can also manage your WBS structure, add and draw relationships between your activities and use them to calculate your project network, track progress, and perform workload time distribution. Discover the rest of its features!

---

## 3. ABOUT SHAPES POSITIONING IN THE GANTT CHART  
If you are experiencing problems with shapes positioning in the calendar scale, please check if you are working on a screen different from your main one.  
If that's the case, you will need to adjust your display settings to the same resolution and scaling on both screens: the main one and the working one.

---

## 4. GANTTASIZER TAB IN EXCEL RIBBON  
All the options included in the Excel Ribbon have their own self-explanatory tips. Therefore, those will not be repeated in this guide.

---

## 5. SYSTEM COLUMNS IN WORKSHEETS  
All the column headers added when the **'Add Headers'** button is pressed contain information used in the Ganttasizer environment. These system columns cannot be deleted but they can be hidden, moved, or renamed if required by the user. New columns and rows can also be added.

There are eight columns for activity-level setup, eighteen columns with project information, and a separation column between the information table and the chart.

This guide is focused on the **activity-level setup columns**:

- **act/mil style**: Four setups supported by this column for activity shapes.
    - Seventeen options for bars and milestones shape styles in the drop-down cell. Options 1-10 are only supported with bars, while options 11-17 are only supported by milestones.
    - Additional option: **NO** — Only applicable for activities that are part of a timeline. It allows hiding them in the timeline but never in the detailed representation.
    - Additional option: **WINDOW** — Turns activities into windows in your schedule.
    - Fill color for remaining bars and milestones: Use the regular fill color picker built into Excel.

- **shape height**: Ten options to adjust bars and milestones size as a percentage of the row height. The height of windows must be introduced as the number of rows the window will cover.

- **connect style**: Three setups supported by this column for connectors.
    - Six options for line styles in the drop-down cell.
    - Additional option: **NO** — Prevents the connector with the predecessor activity from being drawn.
    - Line color for connectors: Use the regular fill color picker built into Excel.

- **label pos**: Two setups supported by this column for activity labels.
    - Nine positioning options for labels in the drop-down cell, only applicable for timelines. The positioning options are the result of combining two parameters:
        - **Height**: Three levels defined as per shape vertical positioning: 0 (same level), 1 (one level above shape), 2 (two levels above shape).
        - **Alignment**: Three positions defined as per shape horizontal positioning: L (alignment to the left), M (middle), R (right) of the shape.
    - Additional option: **NO** — Prevents the label from being displayed in the regular representation of the Gantt chart in timelines.

- **timeline mode**: Three different types of timeline can be created using this setup. It must be selected only in the line where the timeline will be drawn.
    - **SUM**: Creates a single summary bar for the whole span of time of the activities in the timeline.
    - **MIL**: Creates a milestone for each activity finish date and a bar for the whole span of time of the activities in the timeline. The milestones, labels, and the summary bar can be edited with the rest of the activity-level setups.
    - **ACT**: Creates an activity for each activity in the timeline. The activities and labels can be edited with the rest of the activity-level setups.

- **timeline code**: Free code used to identify the activities that are part of a timeline. The timeline row and the activities that are meant to be part of the timeline must be assigned with the same code.

- **schedule mode**: ALAP constraint, four constraints for Start date and four constraints for Finish date can be selected. Also, two network calculations exceptions can be selected:
    - **NO**: This activity will not take any part in network calculation.
    - **MANUAL**: This activity will not be calculated according to its predecessors but it will affect the calculation of its successors. Its dates are constrained to the data introduced prior to the calculation start.

- **units distrib curve**: Four predefined distribution curves can be selected: linear, s-curve, front loaded, back loaded.

---

There are also some interesting tips about **project information columns** worth pointing out:

- **ACTIVITY ID / DESCRIPTION**: Fill either or both of these columns to define an activity. If, for any row below the headers, neither column is filled, the activities list is deemed to be closed on that row.  
  **IMPORTANT:** Activity ID must be unique and not contain spaces. To define a relationship between two activities, both of them must have an Activity ID.

- **WBS**: WBS levels must be separated by dots (.). You can set up the WBS level color using the regular fill color picker built into Excel on this column.

- **TOTAL / REMAINING DURATION**: Total Duration is always a calculated column. Remaining Duration is a calculated column when drawing the Gantt chart but a user-defined value when scheduling.

- **START / FINISH DATE**: User-defined values are used when drawing the Gantt Chart but dates are calculated using Remaining Duration and Predecessors when scheduling.

- **ACTUAL START / FINISH DATE**: Actual Dates user-defined values are only used when scheduling and are never used for drawing the Gantt Chart.

- **RESUME DATE**: Only applicable when the activity is started before the cutoff date but still in progress after the cutoff date.  
  User-defined values are used when drawing the Gantt Chart but dates are calculated using Remaining Duration and Predecessors when scheduling.

- **CONSTRAINT DATE**: Date to be applied in the calculations when a Start or Finish date constraint is selected in the schedule mode column.

- **BUDGET UNITS**: User-defined values are used to weigh Progress % values in summaries.

- **REMAINING UNITS**: User-defined values are used for units distribution.

---
