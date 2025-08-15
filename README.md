# Excel_Report_Stage__opportunities_odoo
Overall View

This updated Python model logic changes how lead counts are calculated for CRM stages in Odoo.
Instead of only counting leads that are currently in a stage, it counts all leads that have ever visited that stage — without decreasing the count when a lead moves to another stage.

Key Changes:

Tracks stage visit history using mail.message or stage change logs.

When generating reports, counts all unique leads that passed through each stage.

Ensures old stages retain their counts even if leads have moved.

Still groups results by salesperson, sales team, and date for reporting.

Benefit:

Gives a true historical view of stage traffic.

Useful for understanding how many leads each stage handled over time, not just current occupancy.
