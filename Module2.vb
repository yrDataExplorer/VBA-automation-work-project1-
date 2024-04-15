Option Explicit



Function create_pre_offer_email(row_num As Integer)

    Dim outlookApp As Object 'Outlook.Application for early binding
    Dim outlookMail As Object 'Outlook.MailItem
    
    ' Define the variables for the data and attachment file paths in the email
    Dim email_body As String, email_subject As String, formsg_link As String, G50_file_path As String
    Dim unique_ID As String, name As String, designation As String, unit As String, personal_email As String, directorate As String
    
    
    ' define the variables to import html formatted text from text file
    Dim TextFilePath As String, HTMLContent As String, FileSystem As Object, TextFile As Object
    
    ' Create an instance of the file system object
    Set FileSystem = CreateObject("Scripting.FileSystemObject")
    ' Define the path to the text file containing HTML-formatted content
    TextFilePath = candidates_folder_path & "\1. Resource folder for VBA code (do not edit)\Pre-offer email\pre-offer email content 1.txt"
    
    
    
    'Ensure that the main worksheet for candidate data in the database file is active
    database.Worksheets("Main data").Activate
    
    ' assign the values for the personal detail variables
    unique_ID = Worksheets("Main data").Cells(row_num, 1)
    name = Worksheets("Main data").Cells(row_num, 2)
    designation = Worksheets("Main data").Cells(row_num, 5)
    unit = Worksheets("Main data").Cells(row_num, 7)
    personal_email = Worksheets("Main data").Cells(row_num, 4)
    directorate = Worksheets("Main data").Cells(row_num, 6)
    
    
    'define the email subject
    email_subject = "[" & name & "] Application for " & designation & " role with " & directorate & " in Smart Nation Group " & "[Onboarding ID: " & unique_ID & "]"
        
    'Create a new instance of Outlook
    Set outlookApp = CreateObject("Outlook.Application")
    
    'Create a new email
    Set outlookMail = outlookApp.CreateItem(0)
    
   
    ' Check if the file exists
    'If FileSystem.FileExists(TextFilePath) Then
        ' Open the text file and read its content
        Set TextFile = FileSystem.OpenTextFile(TextFilePath, 1) ' 1 for reading
    
        'Create the variable for the form link address
        formsg_link = "form.gov.sg/64b8a3d8fefca20012f9ec30?64b8a4f7b963460012cde379=" & unique_ID
        
        'Create the link for the G50 form
        G50_file_path = candidates_folder_path & "\Annex A - G50.pdf"

        'Create the email body in HTML
        ' Read the entire file content
        HTMLContent = TextFile.ReadAll
        
        ' Close the text file
        TextFile.Close


        'Create the email body in HTML
        email_body = "Dear " + name + ",<br><br>We are keen to explore the role for  " & designation & ", " & unit & ", " + directorate + " with you." + _
        "<br>As for the next step, please complete the <u><a href=" + formsg_link + ">formSG link</a></u> and upload the following documents.<br><br>" + _
        "<ol><li>Softcopy of NRIC (Front & Back)</li><li>Last drawn payslip and latest bonus letter or payslip (if any)</li>" + _
        "<li>Education Certificates & Transcript</li><li>Softcopy of colour Passport-sized photo</li><li>Bank Particulars</li></ol>" + _
        "In addition, please furnish us the attached G50 form by replying to this email by " & numberOfworkingDaysFromNow(3) & "." + _
        "<br><br><br>" & HTMLContent
        
        
        'Add details for the email
        With outlookMail
            .To = personal_email
            .CC = "karolyn_cheong@pmo.gov.sg;   oh_si_qi@pmo.gov.sg"
            .subject = email_subject
            .HTMLBody = email_body
            .Attachments.Add G50_file_path
            
    '        signature = GetOutlookSignature()
    '        If signature <> "" Then
    '            .HTMLBody = .HTMLBody & signature
    '
    '        End If
            
            .Display
            .Save
    '        .Send
        
        End With
        
    'End If
    ' Send email
    'outlookMail.Send
    
    
    ' Check if email was sent successfully
    ' If the SentOn property is not equal to the default date of "1/1/4501"
'    If outlookMail.SentOn <> "1/1/4501" Then
'        MsgBox "Email sent successfully."
'    Else
'        MsgBox "Email not sent."
'    End If



    ' Clean up Outlook objects
    Set outlookMail = Nothing
    Set outlookApp = Nothing
    
    'return the email subject for checking of the successful status
    create_pre_offer_email = email_subject


End Function






Function numberOfworkingDaysFromNow(days As Integer) As String
    
    Dim TodaysDate As Date, dayOfWeek As String
    Dim NumDaysFromNow As Date
    
    TodaysDate = Date
    dayOfWeek = Format(Now(), "dddd")
    
    
    NumDaysFromNow = TodaysDate + days
    
    numberOfworkingDaysFromNow = Format(NumDaysFromNow, "dd, mmmm yyyy")
    
'    If dayOfWeek = "Wednesday" Or dayOfWeek = "Thursday" Or dayOfWeek = "Friday" Then
'
'        NumDaysFromNow = TodaysDate + days + 2
'
'    Else:
'
'        NumDaysFromNow = TodaysDate + days
'
'    End If
'
'    numberOfworkingDaysFromNow = Format(NumDaysFromNow, "dd, mmmm yyyy")

    
End Function







Function create_candidate_folder_documents(row_num As Integer) As Boolean


    'Ensure that the main worksheet for candidate data in the database file is active
    database.Worksheets("Main data").Activate

    Dim directorate As String, unique_ID As String
    directorate = Worksheets("Main data").Cells(row_num, 6)
    unique_ID = Worksheets("Main data").Cells(row_num, 1)
    
    
    Debug.Print "Candidate folder path to be created: " + specific_candidate_folder
    ' Check if the folder already exists
    If Dir(specific_candidate_folder, vbDirectory) = "" Then
        
        MkDir specific_candidate_folder
        Debug.Print "Candidate folder path created."
        
    End If
    
    

    ' Copy the personal and bank particulars form to the candidate folder
    ' FileCopy pp_form, pp_form_candidate
    Call createNewPPform(pp_form, pp_form_candidate)


    ' Check if the folder already exists
    If Dir(specific_candidate_folder, vbDirectory) <> "" And Dir(pp_form_candidate) <> "" Then
        
        create_candidate_folder_documents = True
        
    Else
    
        create_candidate_folder_documents = False
        
    End If
    

End Function


Sub createNewPPform(original_doc_path As String, new_doc_path As String)

    Dim wdApp As Object
    Dim wdDoc As Object
    
    ' Create a new instance of Word
    Set wdApp = CreateObject("Word.Application")
    
    ' Open the existing Word document
    Set wdDoc = wdApp.Documents.Open(original_doc_path) ' Replace with the path to your existing Word document
    
    ' Make changes to the document if needed
    ' For example, you can modify the content, formatting, etc.
    
    ' Save the document with a new name and path
    wdDoc.SaveAs2 new_doc_path ' Replace with the desired new path and file name
    
    ' Close the document
    wdDoc.Close
    
    ' Quit Word application
    wdApp.Quit
    
    ' Release the Word objects
    Set wdDoc = Nothing
    Set wdApp = Nothing



End Sub

Function create_onboarding_form(row_num As Integer) As String

    Dim outlookApp As Object 'Outlook.Application for early binding
    Dim outlookMail As Object 'Outlook.MailItem
    
    
    
    ' Define the variables for storing email data
    Dim email_body As String, link As String, email_subject As String
    Dim signature As String
    Dim name As String, designation As String, unit As String, personal_email As String, directorate As String, inflow_date As String, unique_ID As String
    
    
    
    ' define the variables to import html formatted text from text file
    Dim TextFilePath As String, HTMLContent As String, FileSystem As Object, TextFile As Object
    ' Create an instance of the file system object
    Set FileSystem = CreateObject("Scripting.FileSystemObject")
    ' Define the path to the text file containing HTML-formatted content
    TextFilePath = candidates_folder_path & "\1. Resource folder for VBA code (do not edit)\Onboarding new hire\Onboarding new hire email content 1.txt"
    
    
    ' assign the values for the personal detail variables
    name = Worksheets("Main data").Cells(row_num, 2)
    designation = Worksheets("Main data").Cells(row_num, 5)
    unit = Worksheets("Main data").Cells(row_num, 7)
    personal_email = Worksheets("Main data").Cells(row_num, 4)
    directorate = Worksheets("Main data").Cells(row_num, 6)
    inflow_date = Format(Worksheets("Main data").Cells(row_num, 9), "dd mmmm yyyy")
    unique_ID = Worksheets("Main data").Cells(row_num, 1)

    
    'define the email subject
    email_subject = "[For New Hire] Onboarding preparation for " & designation & " role with " & directorate & " in Smart Nation Group " & "[Onboarding ID: " & unique_ID & "]"

    'Create a new instance of Outlook
    Set outlookApp = CreateObject("Outlook.Application")
    
    'Create a new email
    Set outlookMail = outlookApp.CreateItem(0)
    
    
    ' Check if the file exists
    If FileSystem.FileExists(TextFilePath) Then
        ' Open the text file and read its content
        Set TextFile = FileSystem.OpenTextFile(TextFilePath, 1) ' 1 for reading
        
        'Create the variable for the form link address
        link = "https://form.gov.sg/64b8bf99c3c8e30012de1500?64b8c08b9c64f4001174364e=" & unique_ID
        

        'Create the email body in HTML
        ' Read the entire file content
        HTMLContent = TextFile.ReadAll
        
        ' Close the text file
        TextFile.Close

        'Create the email body in HTML
        email_body = "Hi " + name + ",<br><br>Welcome to SNG! <br>To prepare for your onboarding on " + inflow_date + _
        ", and allow time for preparation on our end, please complete the form via this <u><a href=" + link + ">formSG link</a></u>" + _
        " and upload the following documents by <b><u>" + numberOfworkingDaysFromNow(3) + "</u></b>.<br><br>" + _
        HTMLContent & "<br>Feel free to let me know if you have any questions. Thank you!<br><br>"
        

        
        'Add details for the email
        With outlookMail
            .To = personal_email
    '        .CC = "karolyn_cheong@pmo.gov.sg;   oh_si_qi@pmo.gov.sg"
            .subject = email_subject
            .HTMLBody = email_body
            
    '        signature = GetOutlookSignature()
    '        If signature <> "" Then
    '            .HTMLBody = .HTMLBody & signature
    '
    '        End If
            
            .Display
            .Save
    '        .Send
        
        End With
        
    End If
    ' Send email
    'outlookMail.Send

    ' Clean up Outlook objects
    Set outlookMail = Nothing
    Set outlookApp = Nothing
    
    
    'return the email subject for checking of the successful status
    create_onboarding_form = email_subject


End Function


Function create_onboarding_form_secondees(row_num As Integer) As String

    Dim outlookApp As Object 'Outlook.Application for early binding
    Dim outlookMail As Object 'Outlook.MailItem
    
    
    
    ' Define the variables for storing email data and attachment file paths
    Dim email_body As String, link As String, G50_file As String, email_subject As String
    Dim signature As String
    Dim name As String, designation As String, unit As String, personal_email As String, directorate As String, inflow_date As String, unique_ID As String
    
    
    
    ' define the variables to import html formatted text from text file
    Dim TextFilePath As String, HTMLContent As String, FileSystem As Object, TextFile As Object
    ' Create an instance of the file system object
    Set FileSystem = CreateObject("Scripting.FileSystemObject")
    ' Define the path to the text file containing HTML-formatted content
    TextFilePath = candidates_folder_path & "\1. Resource folder for VBA code (do not edit)\Onboarding secondee\Onboarding secondee email content 1.txt"
    


    
    ' assign the values for the personal detail variables
    name = Worksheets("Main data").Cells(row_num, 2)
    designation = Worksheets("Main data").Cells(row_num, 5)
    unit = Worksheets("Main data").Cells(row_num, 7)
    personal_email = Worksheets("Main data").Cells(row_num, 4)
    directorate = Worksheets("Main data").Cells(row_num, 6)
    inflow_date = Format(Worksheets("Main data").Cells(row_num, 9), "dd mmmm yyyy")
    unique_ID = Worksheets("Main data").Cells(row_num, 1)

    'define the email subject
    email_subject = "Onboarding preparation for " & designation & " role with " & directorate & " in Smart Nation Group " & "[Onboarding ID: " & unique_ID & "]"
    

    ' assign the file path to the variable for the G50 file
    G50_file = candidates_folder_path & "\Annex A - G50.pdf"
    
    'Create a new instance of Outlook
    Set outlookApp = CreateObject("Outlook.Application")
    
    'Create a new email
    Set outlookMail = outlookApp.CreateItem(0)
    

   ' Check if the file exists
    If FileSystem.FileExists(TextFilePath) Then
        ' Open the text file and read its content
        Set TextFile = FileSystem.OpenTextFile(TextFilePath, 1) ' 1 for reading
        
        'Create the variable for the form link address
        link = "https://form.gov.sg/64c1a1e356742e001105d835?64b8c08b9c64f4001174364e=" & unique_ID
        

        'Create the email body in HTML
        ' Read the entire file content
        HTMLContent = TextFile.ReadAll
        
        ' Close the text file
        TextFile.Close

        'Create the email body in HTML
        email_body = "Hi " + name + ",<br><br>Welcome to SNG! <br>To prepare for your onboarding on " + inflow_date + _
        ", and allow time for preparation on our end, please complete the form via this <u><a href=" + link + ">formSG link</a></u><br><br>" + _
        "In addition, please furnish us the attached G50 form by replying to this email.<br><br>" & HTMLContent + _
        "We look forward to receiving your response and inputs by <b><u>" + numberOfworkingDaysFromNow(3) + "</u></b>.<br><br>" + _
        "Feel free to let me know if you have any questions. Thank you! "


        
        'Add details for the email
        With outlookMail
            .To = personal_email
    '        .CC = "karolyn_cheong@pmo.gov.sg;   oh_si_qi@pmo.gov.sg"
            .subject = email_subject
            .HTMLBody = email_body
            .Attachments.Add G50_file
            
    '        signature = GetOutlookSignature()
    '        If signature <> "" Then
    '            .HTMLBody = .HTMLBody & signature
    '
    '        End If
            
            .Display
            .Save
    '        .Send
        
        End With
        
    End If
    ' Send email
    'outlookMail.Send
    
    
    ' Clean up Outlook objects
    Set outlookMail = Nothing
    Set outlookApp = Nothing
    
    
    'return the email subject for checking of the successful status
    create_onboarding_form_secondees = email_subject


End Function




Function create_email_dexun_itd(row_num As Integer, emas_form As String) As String

    Dim outlookApp As Object 'Outlook.Application for early binding
    Dim outlookMail As Object 'Outlook.MailItem
    
    
    ' Define the variables for storing email data and attachment file paths
    Dim email_body As String, link As String, fullname As String, email_subject As String
    Dim unique_ID As String
    ' Define the variables for each of the table contents
    Dim inflow_date As String, name As String, designation As String, work_email As String, hp_number As String, grab_hp As String, work_phone As String, secure_email_access As String
    Dim pki_card As String, supervisor As String, reporting_officer As String
    

    ' assign the values for the personal detail variables
    inflow_date = Format(Worksheets("Main data").Cells(row_num, 9), "dd mmmm yyyy") + ", " + WeekdayName(Weekday(Worksheets("Main data").Cells(row_num, 9)))
    name = Worksheets("Main data").Cells(row_num, 2)
    designation = Worksheets("Main data").Cells(row_num, 5) + " (" + Worksheets("Main data").Cells(row_num, 7) + ")" + ", " + Worksheets("Main data").Cells(row_num, 6)
    work_email = Worksheets("Main data").Cells(row_num, 3)
    hp_number = CStr(Worksheets("Main data").Cells(row_num, 12))
    grab_hp = "Tbc"
    work_phone = "Tbc"
    secure_email_access = "Tbc"
    pki_card = "Tbc"
    supervisor = "Tbc"
    reporting_officer = "Tbc"
    unique_ID = Worksheets("Main data").Cells(row_num, 1)
    
    
    email_subject = " Inflow - " & name & " [Onboarding ID: " & unique_ID & "]"
    

    'Create a new instance of Outlook
    Set outlookApp = CreateObject("Outlook.Application")
    
    'Create a new email
    Set outlookMail = outlookApp.CreateItem(0)
    
    
    'Create the email body in HTML
    email_body = "<p>Dear All,</p>" + _
    "<p>Please see below inflow for your respective follow up. Thanks!</p>" + _
    "<table style=""border-collapse: collapse; border: 1px solid black;""><tbody><tr>" + _
    "<td style=""border: 1px solid black; padding: 5px; background: rgb(222,234,246)""><p><strong><span>Inflow date</span></strong></p></td>" + _
    "<td style=""border: 1px solid black; padding: 5px; ""><p>" + inflow_date + "</p></td></tr>" + _
    "<tr><td style=""border: 1px solid black; padding: 5px; background: rgb(222,234,246)""><p><strong><span>Name</span></strong></p></td>" + _
    "<td style=""border: 1px solid black; padding: 5px;""><p>" + name + "</p></td></tr>" + _
    "<tr><td style=""border: 1px solid black; padding: 5px; background: rgb(222,234,246)""><p><strong><span>Designation</span></strong></p></td><td style=""border: 1px solid black; padding: 5px;""><p>" + designation + "</p></td></tr>" + _
    "<tr><td style=""border: 1px solid black; padding: 5px; background: rgb(222,234,246)""><p><strong><span>Email Creation</span></strong></p></td><td style=""border: 1px solid black; padding: 5px;""><p><a><span><span>" + work_email + "</span></span></a>&nbsp;</p></td></tr>" + _
    "<tr><td style=""border: 1px solid black; padding: 5px; background: rgb(222,234,246)""><p><strong><span>Personal Mobile Number</span></strong></p></td><td style=""border: 1px solid black; padding: 5px;""><p>" + hp_number + "</p></td></tr>" + _
    "<tr><td style=""border: 1px solid black; padding: 5px; background: rgb(222,234,246)""><p><strong><span>HP for Grab for Work</span></strong></p></td><td style=""border: 1px solid black; padding: 5px;""><p><span>" + grab_hp + "</span></p></td></tr>" + _
    "<tr><td style=""border: 1px solid black; padding: 5px; background: rgb(222,234,246)""><p><strong><span>iPad/iPhone Choice</span></strong></p></td><td><p>iPad</p></td></tr>" + _
    "<tr><td style=""border: 1px solid black; padding: 5px; background: rgb(222,234,246)""><p><strong><span>Work Phone</span></strong></p></td><td style=""border: 1px solid black; padding: 5px;""><p><span>Tbc</span></p></td></tr>" + _
    "<tr><td style=""border: 1px solid black; padding: 5px;  background: rgb(222,234,246)""><p><strong><span>Secure email access</span></strong></p></td><td style=""border: 1px solid black; padding: 5px;""><p><span>" + secure_email_access + "</span></p></td></tr>" + _
    "<tr><td style=""border: 1px solid black; padding: 5px; background: rgb(222,234,246)""><p><strong><span>To issue PKI card/S notebook</span></strong></p></td><td style=""border: 1px solid black; padding: 5px;""><p><span>" + pki_card + "</span></p></td></tr>" + _
    "<tr><td style=""border: 1px solid black; padding: 5px;  background: rgb(222,234,246)""><p><strong><span>EMAS access request form</span></strong></p></td><td style=""border: 1px solid black; padding: 5px;""><p><span>See attached</span></p></td></tr>" + _
    "<tr><td style=""border: 1px solid black; padding: 5px; background: rgb(255,242,204)""><p><strong><u><span>For Outlook Org Structure:</span></u></strong></p><p><strong><span>Supervisor</span></strong></p><p><strong><span>(who the officer reports to)</span></strong></p></td><td style=""border: 1px solid black; padding: 5px;""><p><span>" + supervisor + "</span></p></td></tr>" + _
    "<tr><td style=""border: 1px solid black; padding: 5px; background: rgb(255,242,204)""><p><strong><span>Reporting Officer of</span></strong></p><p><strong><span>(who reports to the officer)</span></strong></p></td><td style=""border: 1px solid black; padding: 5px;""><p><span>" + reporting_officer + "</span></p></td></tr>" + _
            "</tbody></table><br>" + _
            "<p>Thank you.&nbsp;<span></span></p>"
            
    
    'Add details for the email
    With outlookMail
        .To = ""
        .CC = "tan_yan_rui@pmo.gov.sg"
        .subject = email_subject
        .HTMLBody = email_body
        .Attachments.Add emas_form
        
'        signature = GetOutlookSignature()
'        If signature <> "" Then
'            .HTMLBody = .HTMLBody & signature
'
'        End If
        
        .Display
        .Save
    
    End With

    ' Clean up Outlook objects
    Set outlookMail = Nothing
    Set outlookApp = Nothing
    
    
    'return the email subject for checking of the successful status
    create_email_dexun_itd = email_subject
    
    

End Function


Function create_ask_supervisor_buddy_info_email(row_num As Integer) As String

    Dim outlookApp As Object 'Outlook.Application for early binding
    Dim outlookMail As Object 'Outlook.MailItem
    
    
    
    ' Define the variables for storing email data and attachment file paths
    Dim email_body As String, officer_details As String, link As String, email_subject As String
    ' Define the variables for each of the table contents
    Dim name As String, designation As String, inflow_date As String, inflow_mode As String, hp_number As String, personal_email As String, unique_ID As String
    


    
    ' assign the values for the personal detail variables
    name = Worksheets("Main data").Cells(row_num, 2)
    designation = Worksheets("Main data").Cells(row_num, 5) + " (" + Worksheets("Main data").Cells(row_num, 7) + ")" + ", " + Worksheets("Main data").Cells(row_num, 6)
    inflow_date = Format(Worksheets("Main data").Cells(row_num, 9), "dd mmmm yyyy") + ", " + WeekdayName(Weekday(Worksheets("Main data").Cells(row_num, 9)))
    inflow_mode = Worksheets("Main data").Cells(row_num, 8)
    hp_number = CStr(Worksheets("Main data").Cells(row_num, 12))
    personal_email = Worksheets("Main data").Cells(row_num, 4)
    unique_ID = Worksheets("Main data").Cells(row_num, 1)

    'define the email subject
    email_subject = "[For Inputs] Onboarding preparation for " & name & " [Onboarding ID: " & unique_ID & "]"
    
    
    'Create the variable for the form link address
    link = "https://form.gov.sg/64c1a4a884674a0012c9aacf?64c1a528dba112001142f209=" & name & "&64e30774ffdaa600133393fd=" & designation & "&660ae4f0c8e9ba592fba08d9=" & unique_ID

    'Create a new instance of Outlook
    Set outlookApp = CreateObject("Outlook.Application")
    
    'Create a new email
    Set outlookMail = outlookApp.CreateItem(0)
    
    'table for the details of new hire
    officer_details = "<p><strong><u>Details of new hire:</u></strong></p><table style=""border-collapse: collapse; border: 1px solid black;""><tbody><tr>" + _
            "<td style=""border: 1px solid black; padding: 5px; background: rgb(222,234,246)""><p><strong><span>Name</span></strong></p></td><td style=""border: 1px solid black; padding: 5px; ""><p>" + name + "</p></td></tr><tr>" + _
            "<td style=""border: 1px solid black; padding: 5px; background: rgb(222,234,246)""><p><strong><span>Designation</span></strong></p></td><td style=""border: 1px solid black; padding: 5px; ""><p><span>" + designation + "</p></td></tr>" + _
            "<tr><td style=""border: 1px solid black; padding: 5px; background: rgb(222,234,246)""><p><strong><span>Start date</span></strong></p></td><td style=""border: 1px solid black; padding: 5px; ""><p>" & inflow_date & "</p></td></tr>" + _
            "<tr><td style=""border: 1px solid black; padding: 5px; background: rgb(222,234,246)""><p><strong><span>Mode of joining&nbsp;</span></strong></p></td><td style=""border: 1px solid black; padding: 5px; ""><p>" + inflow_mode + "</p></td></tr>" + _
            "<tr><td style=""border: 1px solid black; padding: 5px; background: rgb(222,234,246)""><p><strong><span>Personal number</span></strong></p></td>" + _
            "<td style=""border: 1px solid black; padding: 5px; ""><p><span>" + hp_number + "</span></p></td></tr><tr>" + _
            "<td style=""border: 1px solid black; padding: 5px; background: rgb(222,234,246)"" ><p><strong><span>Personal email</span></strong></p></td>" + _
            "<td style=""border: 1px solid black; padding: 5px; ""><p><span><a><span><span>" + personal_email + "</span></span></a></span></p></td></tr></tbody></table>"

    
    
    'Create the email body in HTML
    email_body = "Hi &lt;hiring manager&gt;,<br><br>We are happy to share that " + name + _
    " will be joining us on " + inflow_date + "." + _
    "<br><br>In preparation of the his onboarding, we need your inputs via this <u><a href=""" + link + """>formSG link</a></u> by <b><u>" + numberOfworkingDaysFromNow(3) + "</u></b>" + _
    "<br>We will then share with the supervisor and buddy assigned on their specific roles, and work with the respective teams to prepare for the necessary access request.<br><br>" + _
    officer_details + "<br><br>Thank You!<br>"
    
    
    
    'Add details for the email
    With outlookMail
        .To = "tan_yan_rui@pmo.gov.sg"
'        .CC = "karolyn_cheong@pmo.gov.sg;   oh_si_qi@pmo.gov.sg"
        .subject = email_subject
        .HTMLBody = email_body
        
'        signature = GetOutlookSignature()
'        If signature <> "" Then
'            .HTMLBody = .HTMLBody & signature
'
'        End If
        
        .Display
        .Save
    
    End With
    

    ' Clean up Outlook objects
    Set outlookMail = Nothing
    Set outlookApp = Nothing
    
    'return the email subject for checking of the successful status
    create_ask_supervisor_buddy_info_email = email_subject

End Function


Sub createNewEmasform(original_doc_path As String, new_doc_path As String)

    Dim wb As Workbook
    Dim newFilePath As String
    
    
    ' Open the existing Excel workbook
    Set wb = Workbooks.Open(original_doc_path) '

    
    ' Make changes to the document if needed
    ' For example, you can modify the content, formatting, etc.
    
    ' Set the new file path and name
    newFilePath = new_doc_path ' Replace with the desired new path and file name
    
    ' Save the workbook with a new name and path
    wb.SaveAs newFilePath
    
    ' Close the workbook
    wb.Close
    
    ' Release the workbook object
    Set wb = Nothing
    
    

End Sub



Sub edit_emas_form(filepath As String, fullname As String, designation As String, directorate As String, start_date As String, unit As String)

    ' define the variable to reference the candidate's emas request form
    Dim wb As Workbook
    Set wb = Workbooks.Open(filepath)
    
    
    ' enter the candidate info into the excel form
    wb.Activate
    Range("D17") = fullname
    Range("D19") = designation
    Range("D21") = directorate
    Range("D23") = start_date
    
    
    If directorate = "AED" Then
        Range("B48") = "Y"
        If designation = "Deputy Director" Or designation = "Director" Then Range("D48") = "Y"
    
    ElseIf directorate = "GDO" Then
        Range("B51") = "Y"
        If designation = "Deputy Director" Or designation = "Director" Then Range("D51") = "Y"
        
    ElseIf directorate = "NAIO" Then
        Range("B63") = "Y"
        If designation = "Deputy Director" Or designation = "Director" Then Range("D63") = "Y"
    
    ElseIf directorate = "PGD" Then
        Range("B66") = "Y"
        If designation = "Deputy Director" Or designation = "Director" Then Range("D66") = "Y"
    
    ElseIf directorate = "PPD" Then
        Range("B69") = "Y"
        If designation = "Deputy Director" Or designation = "Director" Then Range("D69") = "Y"
        
    ElseIf directorate = "FNRD" Then
        Range("B72") = "Y"
        If designation = "Deputy Director" Or designation = "Director" Then Range("D72") = "Y"
        
    ElseIf directorate = "SCPO" Then
        Range("B75") = "Y"
        If designation = "Deputy Director" Or designation = "Director" Then Range("D75") = "Y"
        
    ElseIf directorate = "HCD" Then
        If designation = "Deputy Director" Or designation = "Director" Then Range("B54") = "Y"
        
        
    ElseIf unit = "Human Capital" Then
        Range("B56") = "Y"
        If designation = "Deputy Director" Or designation = "Director" Or designation = "Senior Manager" Then Range("F56") = "Y"
        If designation = "Deputy Director" Or designation = "Director" Or designation = "Senior Manager" Or designation = "Manager" Then Range("D56") = "Y"
        
        
    ElseIf unit = "Talent and Leadership Management" Then
        Range("B58") = "Y"
        If designation = "Deputy Director" Or designation = "Director" Or designation = "Assistant Director" Then Range("F58") = "Y"
        
    ElseIf unit = "Human Resource Policy" Then
        Range("B60") = "Y"
        If designation = "Deputy Director" Or designation = "Director" Or designation = "Assistant Director" Then Range("F60") = "Y"
        
        
    End If
    
    

    'Save the changes and close the candidate's emas request form
    wb.Save
    wb.Close
    


End Sub




