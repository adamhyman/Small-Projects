
#### Input

To process files exported from Advarra's Clinical Conductor

The first few columns (Patient Name, Status, Total Completed, etc) are static.

Then there are attributes, like Prescreen, Screening, Day 1, Day 2, Day 5, Day 10, etc.
The number of attributes and names can change.

There are 3 columns for each attribute:
1.  Date
2.  Status
3.  Monitored

#### Objective

1.  In attribute columns, Date will be cleared, when it's corresponding Status is NOT Completed.
2.  Status & Monitored columns will be deleted.
3.  The table should be sorted by Screen# (Column position can change)
4.  Delete all rows where Status (Column position can change) is Non-Qualified
5.  Delete Prescreen column (Containing Prescreen dates)

Highlighting added to screenshot only, to show cells that will have the date removed.
![Screenshot](Screenshot.png)
