Option Explicit


Function stage2_automation(row_num As Integer) As Boolean


    'Ensure that the main worksheet for candidate data in the database file is active
    database.Worksheets("Main data").Activate

    ' Define the variables to store the data needed for emas form info and sending of emails
    Dim fullname As String, designation As String, directorate As String, start_date As String, unit As String, inflow_mode As String
    
    fullname = Worksheets("Main data").Cells(row_num, 1)
    designation = Worksheets("Main data").Cells(row_num, 4)
    directorate = Worksheets("Main data").Cells(row_num, 5)
    start_date = Worksheets("Main data").Cells(row_num, 8)
    unit = Worksheets("Main data").Cells(row_num, 6)
    inflow_mode = Worksheets("Main data").Cells(row_num, 8)
    


    ' Define the variables for the process to create and edit the emas access form
    Dim emas_access_default As String, emas_access_candidate As String
    emas_access_default = candidates_folder_path + "\SNDGO Onboarding Form Template.xlsx"
    emas_access_candidate = specific_candidate_folder + "\SNDGO Onboarding Form Template.xlsx"
    Debug.Print "Emas access file path for candidate: " & emas_access_candidate
    
    Call Module2.createNewEmasform(emas_access_default, emas_access_candidate)
    Call Module2.edit_emas_form(emas_access_candidate, fullname, designation, directorate, start_date, unit)
    
    

    ' Define the arrays to store the different types of inflow mode for each category of inflow
    Dim new_hire_array As Variant
    Dim secondment_array As Variant
    
    
    new_hire_array = Array("Open Recruitment on MX", "MOPO", _
    "PSLP GP 1st", "PSLP GP 2nd", "PSLP GP 3rd", "PSLP GP - Eng", "PSLP SP")
    
    secondment_array = Array("Secondment", "Transfer", "IO Posting", "AO Posting", "LS Posting", _
                            "Open Recruitment on GovTech", "Open Recruitment on OGP")
    
    
    
    ' Define the final boolean value for the function to return after checking for the successful creation of each email
    Dim final_result As Boolean
    final_result = True 'set the default value as True
    
    ' Define the variable for the function to return the email subject for checking the success status of the email creation
    Dim check_subject As String
    
    
    
    ' Check for candidate's inflow mode in order to determine the correct onboarding email to be sent
    Dim Count As Integer
    For Count = LBound(new_hire_array) To UBound(secondment_array)
        If new_hire_array(Count) = inflow_mode Then
            
            ' If the candidate joined via open recruitment, send onboarding prep form for open recruitment
            Debug.Print "Creating email to send onboarding prep form for open recuitment to candidate"
            
                                
            check_subject = Module2.create_onboarding_form(row_num)
            
            If Module1.IsDraftEmailExist(check_subject) <> True Then 'check if onboarding email has been successfully created
            
                final_result = False 'edit the final_result boolean value if the creation is not successful
        
            End If
    
        ElseIf secondment_array(Count) = inflow_mode Then
            
            ' If the candidate joined via secondment/posting, send onboarding prep form for secondees
            Debug.Print "Creating email to send onboarding prep form for secondees/posting to candidate"

            check_subject = Module2.create_onboarding_form_secondees(row_num)
            
            If Module1.IsDraftEmailExist(check_subject) <> True Then 'check if onboarding email has been successfully created
            
                final_result = False 'edit the final_result boolean value if the creation is not successful
        
            End If


        End If
    
    Next Count


    ' create email for sending info to Dexun/ITD
    Debug.Print "Creating email for sending info to Dexun/ITD"
    
    check_subject = Module2.create_email_dexun_itd(row_num, emas_access_candidate)

    If Module1.IsDraftEmailExist(check_subject) <> True Then 'check if onboarding email has been successfully created
    
        final_result = False 'edit the final_result boolean value if the creation is not successful

    End If

    ' create email for checking with directorate for inputs on supervisor/buddy
    Debug.Print "Creating email for checking with directorate for inputs on supervisor/buddy"
    
    check_subject = Module2.create_ask_supervisor_buddy_info_email(row_num)
    
    If Module1.IsDraftEmailExist(check_subject) <> True Then 'check if onboarding email has been successfully created
    
        final_result = False 'edit the final_result boolean value if the creation is not successful

    End If
    
    
    
    ' return the boolean value for checking the status for the successful creation of all emails
    stage2_automation = final_result
    
    

End Function





Function search_and_reply_email_dexun_itd(row_num As Integer) As Boolean


    'Ensure that the main worksheet for candidate data in the database file is active
    database.Worksheets("Main data").Activate


    ' Define the variables for the outlook application
    Dim outlookApp As Object
    Dim outlookNamespace As Object
    Dim inboxFolder As Object
    Dim mailItem As Object, mailItems As Object, mailAttachment As Object
    Dim myReply As Variant
    
    
    
    ' define the unique ID variable and value
    Dim unique_ID As String
    unique_ID = Worksheets("Main data").Cells(row_num, 1)
    
    
    ' Define the search term to look for in the email subject
    Dim searchSubject As String, search_result As Boolean
    searchSubject = unique_ID
    
    
    
    ' Define the variables for each of the table contents
    Dim inflow_date As String, name As String, designation As String, work_email As String, hp_number As String, grab_hp As String, work_phone As String, secure_email_access As String
    Dim pki_card As String, supervisor As String, reporting_officer As String
    ' Define the variable for the email body
    Dim email_body As String

    
    ' assign the values for the personal detail variables
    inflow_date = Format(Worksheets("Main data").Cells(row_num, 9), "dd, mmmm yyyy") + ", " + WeekdayName(Weekday(Worksheets("Main data").Cells(row_num, 9)))
    name = Worksheets("Main data").Cells(row_num, 2)
    designation = Worksheets("Main data").Cells(row_num, 5) + " (" + Worksheets("Main data").Cells(row_num, 7) + ")" + ", " + Worksheets("Main data").Cells(row_num, 6)
    work_email = Worksheets("Main data").Cells(row_num, 3)
    hp_number = CStr(Worksheets("Main data").Cells(row_num, 12))
    grab_hp = CStr(Worksheets("Main data").Cells(row_num, 21))
    work_phone = "Tbc"
    secure_email_access = Worksheets("Main data").Cells(row_num, 27)
    pki_card = Worksheets("Main data").Cells(row_num, 28)
    supervisor = Worksheets("Main data").Cells(row_num, 24)
    reporting_officer = Worksheets("Main data").Cells(row_num, 26)

    
    
    
    
    'Create the email body in HTML
    email_body = "<p>Dear All,</p>" + _
    "<p>Please see below inflow for your respective follow up. Thanks!</p>" + _
    "<table style=""border-collapse: collapse; border: 1px solid black;""><tbody><tr>" + _
    "<td style=""border: 1px solid black; padding: 5px; background: rgb(222,234,246)""><p><strong><span>Inflow date</span></strong></p></td>" + _
    "<td style=""border: 1px solid black; padding: 5px; ""><p>" + inflow_date + "</p></td></tr>" + _
    "<tr><td style=""border: 1px solid black; padding: 5px; background: rgb(222,234,246)""><p><strong><span>Name</span></strong></p></td>" + _
    "<td style=""border: 1px solid black; padding: 5px;""><p>" + name + "</p></td></tr>" + _
    "<tr><td style=""border: 1px solid black; padding: 5px; background: rgb(222,234,246)""><p><strong><span>Designation</span></strong></p></td><td style=""border: 1px solid black; padding: 5px;""><p>" + designation + "</p></td></tr>" + _
    "<tr><td style=""border: 1px solid black; padding: 5px; background: rgb(222,234,246)""><p>< strong><span>Email Creation</span></strong></p></td><td style=""border: 1px solid black; padding: 5px;""><p><a><span><span>" + work_email + "</span></span></a>&nbsp;</p></td></tr>" + _
    "<tr><td style=""border: 1px solid black; padding: 5px; background: rgb(222,234,246)""><p><strong><span>Personal Mobile Number</span></strong></p></td><td style=""border: 1px solid black; padding: 5px;""><p>" + hp_number + "</p></td></tr>" + _
    "<tr><td style=""border: 1px solid black; padding: 5px; background: rgb(222,234,246)""><p><strong><span>HP for Grab for Work</span></strong></p></td><td style=""border: 1px solid black; padding: 5px;""><p><span style=""background-color: yellow;"">" + grab_hp + "</span></p></td></tr>" + _
    "<tr><td style=""border: 1px solid black; padding: 5px; background: rgb(222,234,246)""><p><strong><span>iPad/iPhone Choice</span></strong></p></td><td><p>iPad</p></td></tr>" + _
    "<tr><td style=""border: 1px solid black; padding: 5px; background: rgb(222,234,246)""><p><strong><span>Work Phone</span></strong></p></td><td style=""border: 1px solid black; padding: 5px;""><p><span>Tbc</span></p></td></tr>" + _
    "<tr><td style=""border: 1px solid black; padding: 5px;  background: rgb(222,234,246)""><p><strong><span>Secure email access</span></strong></p></td><td style=""border: 1px solid black; padding: 5px;""><p><span style=""background-color: yellow;"">" + secure_email_access + "</span></p></td></tr>" + _
    "<tr><td style=""border: 1px solid black; padding: 5px; background: rgb(222,234,246)""><p><strong><span>To issue PKI card/S notebook</span></strong></p></td><td style=""border: 1px solid black; padding: 5px;""><p><span style=""background-color: yellow;"">" + pki_card + "</span></p></td></tr>" + _
    "<tr><td style=""border: 1px solid black; padding: 5px;  background: rgb(222,234,246)""><p><strong><span>EMAS access request form</span></strong></p></td><td style=""border: 1px solid black; padding: 5px;""><p><span>See attached</span></p></td></tr>" + _
    "<tr><td style=""border: 1px solid black; padding: 5px; background: rgb(255,242,204)""><p><strong><u><span>For Outlook Org Structure:</span></u></strong></p><p><strong><span>Supervisor</span></strong></p><p><strong><span>(who the officer reports to)</span></strong></p></td><td style=""border: 1px solid black; padding: 5px;""><p><span style=""background-color: yellow;"">" + supervisor + "</span></p></td></tr>" + _
    "<tr><td style=""border: 1px solid black; padding: 5px; background: rgb(255,242,204)""><p><strong><span>Reporting Officer of</span></strong></p><p><strong><span>(who reports to the officer)</span></strong></p></td><td style=""border: 1px solid black; padding: 5px;""><p><span style=""background-color: yellow;"">" + reporting_officer + "</span></p></td></tr>" + _
            "</tbody></table><br>" + _
            "<p>Thank you.&nbsp;<span></span></p>"

    

    
    ' Create a new instance of Outlook Application
    Set outlookApp = CreateObject("Outlook.Application")
    ' Get the MAPI namespace
    Set outlookNamespace = outlookApp.GetNamespace("MAPI")
    ' Get the inbox folder
    Set inboxFolder = outlookNamespace.GetDefaultFolder(6).Folders("Onboarding automation emails")

    ' Get all items in the inbox folder and sort them in reverse order
    Set mailItems = inboxFolder.Items
    mailItems.Sort "[ReceivedTime]", True

                       
    Debug.Print "Looping through each mail item subject in the folder"
    
    For Each mailItem In mailItems 'inboxFolder.Items
        Debug.Print "Mail subject found: " & mailItem.subject
        If InStr(1, mailItem.subject, searchSubject, vbTextCompare) > 0 Then
            
'            previous_header = "<blockquote>" & "<p><strong>From:</strong> " & mailItem.SenderName & " &lt;" & mailItem.SenderEmailAddress & "&gt;</p>" & "<p><strong>Sent:</strong> " & mailItem.SentOn & "</p>" & _
'                              "<p><strong>To:</strong> " & mailItem.To & "</p>" & "<p><strong>Cc:</strong> " & mailItem.CC & "</p>" & "<p><strong>Subject:</strong> " & mailItem.Subject & "</p>" & "<br>" & mailItem.HTMLBody & "</blockquote>"

            Set myReply = mailItem.ReplyAll
            myReply.HTMLBody = email_body & vbCrLf & vbCrLf & mailItem.ReceivedTime & mailItem.HTMLBody
        
            
            myReply.Display
            myReply.Save
            
            
            search_result = True
            
            Exit For
            
        End If
        
    Next mailItem
    
    
    If search_result Then
        
        search_and_reply_email_dexun_itd = True
        
    Else
    
        search_and_reply_email_dexun_itd = False
    
    End If
    


End Function



