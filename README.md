# Outlook-Auto-Filtering-Based-on-Excel-Attachment-Contents---VBA


This is a simple script that was written to auto-filter the daily status report emails our system generates based upon the contents of their attached XLS spreadsheets. I am posting this software because it contains within it much of the funcationality people may be interested in for automatically filtering emails based on their attachements within outlook.

<b>BACKGOUND</b>

The servers in an AMI System have the ability to generate daily status reports used to indicate the status of our network. Specifically, there is a report which specifies the number of reverse flow alarms our system recieved over a specific period of times. However, we do not want to sort through all the alarms but rather are only interested in the endpoints that have generated a high number of these alarms. Therefore the incoming emails need to be filtered based on the contents of their attached XLS spreadsheet. Once these endpoints are filtered out, a new spreadsheet containing only these endpoints is created and time-stamped. A new email containing this spreadsheet is generated and sent to a person/address for them to take action.

</br>
</br>

<b>DETAILS</b>

![System Diagram](https://cloud.githubusercontent.com/assets/11066939/9765953/aa3ead60-56e4-11e5-909d-294635adc9e6.JPG )

The filtering process used in this script is quite simple and involves making a few generalizations, one being that all emails sent from a specific person/address contain the report we are interested in. A more complex filtering scheme can easily be implemented if so needed. This script begins by looking at all incoming emails on a specific email account and moves any emails recieved from a specific person/address to a seperate folder.

Once in this folder the script looks for a XLS attachement. Another assumption is that any .XLS spreadsheet attached to these emails is in the proper format, again this could be improved. If no attachement is found, or the attachment that was attached is not a .XLS spreadsheet, then the script generates a pop-up message indicating that a "Reverse Flow Report" was recieved with not attachment found. If an attachement is found than the script begins its filtering process.

The filtering process begins by opening the attached spreadsheet and create a new sheet to contain only the filtered endpoints. It than checks each record in the attached spreadsheet looking for endpoints with greater than 'x' alarms (x = user input). Any endpoint matching this description is then copied to this new spreadsheet. 

<img src="https://cloud.githubusercontent.com/assets/11066939/9763511/f089ebf2-56d7-11e5-921e-f9234db2105e.JPG" alt="Sample Reverse Flow Report" width="526" height="467">
EX. Filtered Reverse Flow Spreadsheet

The new spreadsheet is saved locally in a predifined folder and time-stamped. The script then creates a new email containing a scripted response and has this new spread sheet attached. The email is sent to another predefined person/email address for them to process or take action against. An example of an auto-generated message created from this script can be seen below

<img src="https://cloud.githubusercontent.com/assets/11066939/9764766/cb6b889c-56de-11e5-977a-3f56dce97016.JPG" alt="System Diagram" width="567" height="368">


NOTE: This is a locally run script and does not run on the mail server. This means the filtering only occurs while the email application is open on the machine running this script.

<b>HOW TO RUN</b>

1) Write click on the script and open in either a IDE or simply in notepad.

2) Alter the USER INPUT Section near the top of the script to match the specifications you requires:

     -file Location       <--- File location where the time-stamped attachements will be saved
     -Reverse Flow Limit  <--- This value indicates the lower limit to the number of alarms required before it is flagged
     -SubFolder           <--- The name of the outlook subfolder where the incoming emails will be moved prior to filtering
     -SenderFilter        <--- The email address which will be filtered and moved
     -SummaryAddress      <--- The email address which will recieve the summary and filtered spreadsheet
3) Install the Script in your outlook application

    -Open Outlook and navigate to tools->macros->Visual Basic Editor (You will most likely be confronted with a warning message, click run macros)
    -Right click the "ThisOutlookSession" in the window explorer on the right and click import
    -select the modified .cls file containing the changes above
    -This creates a new module, copy the contents of the new module into the "ThisOutlookSession" sheet.
    -Save and restart microsoft outlook
    
4) To Uninstall: navigate to the Visual Basic Editor, delete the contents of the "ThisOutlookSession" sheet, click save and restart Outlook
