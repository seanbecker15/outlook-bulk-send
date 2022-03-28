## Outlook-Mass-Send-Suurvey
##### Visual Basic script developed to parse through an outlook folder and send surveys.


#### Purpose
At my first internship, one of my tasks was to send a survey for each help-desk ticket we closed during the past week. That took a ton of time (there could be upwards of 100 tickets closed every week) so I decided to make a script to do it for me!

#### How it works
- Traverse through outlook folder
- Parse for name of employee
- Generate employee's email address (<first letter of first name>.<last name>@company.com, e.g. s.becker@company.com)
- Send email with link to survey

#### Email format
Associate Name: Sean Becker

Department Number: 52

Associate Department: IT

Request Number: 160722130706

Status: 
http://notesmail/Prod/NetworkRequest.nsf/0/#####################

