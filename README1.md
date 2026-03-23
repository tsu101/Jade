# VBA Modules & Macros – README
## VBA Project Password: Bangchan
###This document provides an overview of all UserForms, Modules, and their functions used in the Ticket Management Excel File.

## User Forms

### frmEmployees
- A user form for adding new employees to the employee list.

## Modules Overview

### Module1 – Employee Management
Contains macros for:
- Adding employees
- Removing employees
- These macros are directly associated with frmEmployees.


### Module2 – Import Closed Tickets
Handles importing closed ticket Excel files into the Closed_Ticket worksheet.

### Module3 – Detailed Ticket Count
Performs detailed counting for Helpdesk and TechOps closed tickets.
Uses the data imported to Closed_Ticket (from Module2).
Computed results are written to the Summary sheet.

### Module4 – Group Dashboard (Overall_Report(Group))
Generates dashboard metrics for internal and GSC Helpdesk groups.
Outputs to the Overall_Report(Group) sheet.

### Module5 – Weekly Tracker Compilation
Compiles the detailed count generated in Module3
and appends the results to the WeeklyTracker sheet.

### Module6 – Time-Worked Summary
Calculates total time worked for each Helpdesk and TechOps staff.

### Module7 – Full Process Automation
Runs all essential macros in sequence through a single button.

### Module8 – Time-Worked Summary with Escalations
Performs the same functions as Module6, with additional handling for time-worked escalations.
Also includes detailed counting similar to Module3, and its computed results are written to the Summary sheet.

### Module9 – Computation Sheet Updates
Updates and maintains formulas and computed values in the Computation sheet.

### Module10 – Pivot Filter Handling
Contains macros for filtering PivotTables.

### Module11 – Import Open Tickets
A counterpart to Module2, but for open tickets.
Imports Excel files into the Open_Ticket worksheet.

### Module12 – Pivot Refresh
Automatically refreshes PivotTables after data updates.

### Module13 – Individual Dashboard
Similar to Module4, but generates dashboards for individual Helpdesk members.
