Attribute VB_Name = "outlookFilteringScript"
Private WithEvents Items As Outlook.Items


Private Sub Application_Startup()
  Dim olApp As Outlook.Application
  Dim objNS As Outlook.NameSpace
  Set olApp = Outlook.Application
  Set objNS = olApp.GetNamespace("MAPI")
  ' default local Inbox
  Set Items = objNS.GetDefaultFolder(olFolderInbox).Items
End Sub
Private Sub Items_ItemAdd(ByVal item As Object)

  On Error GoTo ErrorHandler
  Dim Msg As Outlook.MailItem
  If TypeName(item) = "MailItem" Then
    Set Msg = item
    
    'call SaveAttachementsToFolder Script
    Call SaveAttachmentsToFolder(item)

  End If
ProgramExit:
  Exit Sub
ErrorHandler:
  MsgBox Err.Number & " - " & Err.Description
  Resume ProgramExit
End Sub



Sub SaveAttachmentsToFolder(newItem As MailItem)

'On Error GoTo SaveAttachmentsToFolder_err
    
' Declare variables

    Dim ns As NameSpace
    Dim Inbox As MAPIFolder
    Dim SubFolder As MAPIFolder
    Dim subFolderName As String
    Dim item As Object
    Dim Atmt As Attachment
    Dim fileLocation As String
    Dim FileName As String
    Dim count As Integer
    Dim varResponse As VbMsgBoxResult
    Dim senderFilter As String
    
'Excel Workbook Variables
    Dim numRows As Long
    
    
'create an excel application object
    Dim appExcel As Excel.Application
    Dim ExWbk As Excel.Workbook
    Dim WS As Excel.Worksheet


'Creating New Sheet Variables
    Dim sheetName As String
    Dim sheetExists As Boolean
    Dim HRFS As Worksheet
        
'HRF sheet Variables
    Dim ReverseFlowLimit As Integer
    Dim NumReverseFlow As Integer
    Dim HRFRow As Long
    
    NumReverseFlow = 0
   
  
'email object
    Dim objMsg As Object
       
    
'****************************
'         USER INPUT
'****************************
    
    'Set FileLocation
    fileLocation = "C:\Users\terradeda\Documents\Drop Box\"
    
    'Set the number of reverse flow alarms required to trigger a response
    ReverseFlowLimit = 12
    
    'Set the name of the Subfolder
    subFolderName = "Drop Box"
    
    'Set the name of the sender to filter out
    senderFilter = "Terrade, David"
    
'****************************
    
' Initialize Variables

    Set ns = GetNamespace("MAPI")
    Set Inbox = ns.GetDefaultFolder(olFolderInbox)
    Set SubFolder = Inbox.Folders(subFolderName)
    count = 0
    
    
    '****************************************************************
    'CHECK IF NEW EMAIL IF FROM A SPECIFIC USER AND MOVE TO SUBFOLDER
    '****************************************************************
    'Debugging
    'MsgBox "item: " & newItem.SenderName & "   " & StrComp(newItem.SenderName, "Terrade, David", 1)
    
    'If Email is from Specific user then move it to the 'Drop Box' Folder
    If StrComp(newItem.SenderName, senderFilter, 1) = 0 Then
        
        newItem.Move SubFolder
    Else
    
        Exit Sub
    End If
    
    Set SubFolder = Inbox.Folders(subFolderName)
    
    '****************************************************************
    
    
' Check for messages in the Subfolder and exit if none/no new

    If SubFolder.Items.count = 0 Then
        MsgBox "There are no emails in the '" & subFolderName & "' subfolder", vbInformation
        Exit Sub
    ElseIf SubFolder.UnReadItemCount = 0 Then
        MsgBox "There are no new emails in the '" & subFolderName & "' subfolder", vbInformation
        Exit Sub
    End If
    
    Set appExcel = CreateObject("Excel.Application")
  
' Check each message for attachments and save the CSV attachments
    For Each item In SubFolder.Items
        
        'Only Process Unread Messages
        If item.UnRead = True Then
        
           For Each Atmt In item.Attachments
            'check for attachments with an "xls" file type
            If Right(Atmt.FileName, 3) = "xls" Then
                
                'debugging
                'MsgBox "Found Attachement: " & Atmt.FileName & vbCrLf & "Attached to Email: " & item

                'Save the attached file at path: filename
                FileName = fileLocation & Format(item.CreationTime, "yymmddThhmmss_") & Atmt.FileName
                Atmt.SaveAsFile FileName
                
                '**************************
                'OPEN AND SETUP EXCEL SHEET
                '**************************
                
                'open the attachement
                Set ExWbk = appExcel.Workbooks.Open(FileName)
                'Hide Excel workbook
                appExcel.Visible = False
                Set WS = ExWbk.ActiveSheet

                'Calculate the number of rows in the spreadsheet
                numRows = WS.UsedRange.Rows.count
                
                '******************************************
                'CREATE HIGH REVERSE FLOW ALARMS WORK SHEET
                '******************************************
                
                'Define the name of the new worksheet
                sheetName = "High # of Reverse Flow alarms"
                
                For Each sh In ExWbk.Sheets
                     If sh.name Like sheetName Then
                        sheetExists = True: Exit For
                     End If
                Next
                 
                    
                If sheetExists = False Then
                    appExcel.DisplayAlerts = True
                    On Error GoTo 0
                    Set HRFS = ExWbk.Worksheets.Add()
                    HRFS.name = sheetName
                End If
                    
    
                'Copy Header
                WS.Range("A1", "O13").Copy
                HRFS.Cells(1, 1).Select
                HRFS.Paste
                WS.Select
    
                '*******************************************************
                'FILTER OUT ENDPOINTS WITH HIGH # OF REVERSE FLOW ALARMS
                '*******************************************************
                
                For i = 16 To numRows Step 4
                
                    If WS.Cells(i, 6).Value > ReverseFlowLimit Then
                        'Add one to counter
                        NumReverseFlow = NumReverseFlow + 1
                        
                        'Copy high Reverse floor account
                        WS.Select
                        WS.Range(WS.Cells(i + 1, 1), WS.Cells(i - 2, 13)).Copy
                        HRFRow = 10 + NumReverseFlow * 4
                        HRFS.Select
                        HRFS.Cells(HRFRow, 1).Select
                        HRFS.Paste
                        WS.Select
                        
                                          
                    End If
    
                Next i
                
                
                'select High # of Reverse Flow WS
                HRFS.Select
                
                
                'Save Excel Workbook
                ExWbk.Save
    
                'Close Excel Workbook
                ExWbk.Close
                '*******************************************************
    
    
    
                '*****************************************
                'CREATE NEW EMAIL AND ATTACHED SPREADSHEET
                '*****************************************
                
                If NumReverseFlow > 0 Then
                
                   'Create A New Email
                   Set objMsg = Application.CreateItem(olMailItem)
                   
                   With objMsg
                     .To = "david.terrade@ottawa.ca"
                     .Subject = "High Number of Reverse Flow Alarms - " & Format(item.CreationTime, "yyymmdd_hhmmss")
                     .Categories = "Test"
                     .HTMLBody = "<b>AMI REVERSE FLOW WARNING: </b> A High Number of Reverse Flow Alarms Have Been Detected <br /> <br />" & _
                                 "<i>(An Endpoint is flagged if it has had more then " & ReverseFlowLimit & " Reverse Flow Alarms in the past 24 hours)</i> <br /><br /><br />" & _
                                 "&nbsp; &nbsp; &nbsp; &nbsp; The number of endpoints with a high number of reverse flow alarms: <b>" & NumReverseFlow & "</b><br /><br /><br />" & _
                                 "Check the attached spreadsheet for a list of the problem endpoints"
                             
                     .Attachments.Add (FileName)
                     .Display
                     .Send
                     
                   End With
                   
    
                ElseIf NumReverseFlow = 0 Then
                
                
                MsgBox "REVERSE FLOW EMAIL RECIEVED - No Endpoints with more then " & ReverseFlowLimit & " Reverse Flow Alarms in the past 24 hours were found"
    
                End If
    
    
                '*****************************************
 
            End If
            
            
            Next Atmt
            
            'Set The Email to Unread
            item.UnRead = False
            
        End If
        
 

    Next item
    
    appExcel.Quit

    If i = 0 Then
    
        MsgBox "REVERSE FLOW EMAIL RECIEVED - No Report Attached"

    End If
    
' Clear memory

ClearMemory_exit:
    Set Atmt = Nothing
    Set item = Nothing
    Set ns = Nothing
    Set ExWbk = Nothing
    Set appExcel = Nothing
    Exit Sub

' Error Handling

SaveAttachmentsToFolder_err:
    MsgBox "An unexpected Error has occured " _
    & vbCrLf & "Please Note the following information" _
    & vbCrLf & "Macro Name : Save Attachments to Folder" _
    & vbCrLf & "Error Number:" & Err.Number _
    & vbCrLf & "Error Description: " & Err.Description _
    , vbCritical, "Error"
    Resume ClearMemory_exit
    
    End Sub





