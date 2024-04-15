Option Explicit


Function send_supervisor_buddy_email(row_num As Integer) As String


    Dim outlookApp As Object 'Outlook.Application for early binding
    Dim outlookMail As Object 'Outlook.MailItem
    
    
    ' Define the variables for storing email content data
    Dim email_body As String, email_subject As String
    Dim full_name As String, personal_email As String, buddy_name As String, buddy_hp As String, supervisor_name As String, supervisor_hp As String
    Dim inflow_date As String, designation As String, unit As String, directorate As String, candidate_hp As String
    
        
    'Ensure that the main worksheet for candidate data in the database file is active
    database.Worksheets("Main data").Activate
    

    ' Assign the values for the email content data
    full_name = Worksheets("Main data").Cells(row_num, 2)
    candidate_hp = Worksheets("Main data").Cells(row_num, 12)
    designation = Worksheets("Main data").Cells(row_num, 5)
    unit = Worksheets("Main data").Cells(row_num, 7)
    directorate = Worksheets("Main data").Cells(row_num, 6)
    
    inflow_date = Format(Worksheets("Main data").Cells(row_num, 9), "dd mmmm yyyy") + ", " + WeekdayName(Weekday(Worksheets("Main data").Cells(row_num, 9)))
    buddy_name = Worksheets("Main data").Cells(row_num, 22)
    buddy_hp = Worksheets("Main data").Cells(row_num, 23)
    supervisor_name = Worksheets("Main data").Cells(row_num, 24)
    supervisor_hp = Worksheets("Main data").Cells(row_num, 25)
    
    
    'define the email subject
    email_subject = "Supervisor and Buddy to " & full_name

    

    'Create a new instance of Outlook
    Set outlookApp = CreateObject("Outlook.Application")
    
    'Create a new email
    Set outlookMail = outlookApp.CreateItem(0)
    
    email_body = "Dear " + supervisor_name + " and " + buddy_name + ",<br><br>Thank you for taking on the roles of supervisor and buddy to " + full_name + ", " + designation + " (" + directorate + "). The new officer will commence work on the <u>" + inflow_date + ".</u><br><br>" + _
                 "<b>(" + supervisor_name + ") as a supervisor</b>, your guidance and support to " + full_name + " during the officer's first few months is crucial in ensuring that the officer has a good onboarding experience.<br><br>" + _
                 "As part of the onboarding process, please be in the office to welcome " + full_name + " on the officer's first day to:<br>" + _
                 "<ul><li>Bring the officer to collect the IT equipment from Jorin/Melissa.</li><li>Introduce the officer to the office space and colleagues.</li></ul><br><br>" + _
                 "As you will be an integral part of the officer's journey in SNDGO, this will help " + full_name + " to ease in to the team and organisation on his first day." + _
                 " If you are unavailable, please inform the officer's buddy or other colleagues to do so. As the officer will contact yourself/ officer's buddy or colleague to enter the office, " + _
                 "do let us know who the point of contact should be and provide us the assigned officer's contact number so that we can inform " + full_name + ".<br><br>" + _
                 "In addition, we have prepared the <u><a href=""https://sgdcs.sgnet.gov.sg/sites/PMO-SNDGO/HCD/Training%20Resources/Forms/AllItems.aspx?slrid=bf3bdca0-1ec3-500d-0524-fff4d73099d1&RootFolder=%2fsites%2fPMO%2dSNDGO%2fHCD%2fTraining%20Resources%2fOnboarding%20Materials%2fSupervisor&FolderCTID=0x01200080DE3365590D0A4E9FC92C4D03F0AD44"">[Access here] Supervisor checklist</a></u> intended to serve as a helpful guide, rather than an exhaustive ""to-do"" list for you to engage the officer, for your reference.<br><br>" + _
                 "<b>(" + buddy_name + ") as a buddy</b>, you are integral to the new officer's on-boarding efforts. You would be required to:<br>" + _
                 "<ul><li>Connect " + full_name + " to our fellow SNDGO colleagues (e.g. adding the officer to our Whatsapp Farmers chat group etc.)</li>" + _
                 "<li>Share your work experiences with him and impart valuable tips and resources that will help in his work, e.g. SNDGO Intranet, HRP, Workpal and other directorate-specific resources etc.</li></ul><br><br>" + _
                 "To help you in your role, we have prepared a <u><a href=""https://sgdcs.sgnet.gov.sg/sites/PMO-SNDGO/HCD/Training%20Resources/Forms/AllItems.aspx?slrid=083cdca0-eec7-500d-3d76-23490809cbbd&RootFolder=%2fsites%2fPMO%2dSNDGO%2fHCD%2fTraining%20Resources%2fOnboarding%20Materials%2fBuddy&FolderCTID=0x01200080DE3365590D0A4E9FC92C4D03F0AD44"">[Access here] Buddy checklist</a></u>, comprising of essential information that " + _
                 "the officer will need to learn and know as part of the officer's onboarding. Do go through the checklist with your buddy within the first 2 weeks to help your buddy get set up!<br><br>" + _
                 "For info, " + full_name + "'s first physical day will be on <u>" + inflow_date + ".</u> The supervisor may seek your help to welcome the officer and ease him/her to the team and organisation on the officer's first day. For linking up, you may contact the officer directly at " + candidate_hp + ".<br><br>" + _
                 "Lastly, thank you for agreeing to take on the respective roles and helping your new colleague's onboarding experience be an enjoyable and meaningful one!"

    
    'Add details for the email
    With outlookMail
        .To = "tan_yan_rui@pmo.gov.sg"
        .CC = "Karolyn_CHEONG@pmo.gov.sg; oh_si_qi@pmo.gov.sg"
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
    send_supervisor_buddy_email = email_subject


End Function






Function send_welcome_email(row_num As Integer) As String

    Dim outlookApp As Object 'Outlook.Application for early binding
    Dim outlookMail As Object 'Outlook.MailItem
    
    
    ' Define the variables for storing email content data
    Dim email_body As String, email_subject As String, table As String, content1 As String, content2 As String, content3 As String, content4 As String, content5 As String, content6 As String
    Dim full_name As String, personal_email As String, buddy_name As String, buddy_hp As String, supervisor_name As String, supervisor_hp As String
    Dim inflow_date As String
    
    ' Define the variables for the email attachments
    Dim img1 As String, img2 As String, img3 As String, img4 As String, img5 As String, img6 As String
    Dim fr_guide As String, cyber_security_guide As String, sundae_edm_guide As String, hcs_letter As String, ps_values As String
    
    
    
    'Ensure that the main worksheet for candidate data in the database file is active
    database.Worksheets("Main data").Activate
    
    
    ' Assign the values for the email content data
    full_name = Worksheets("Main data").Cells(row_num, 10)
    personal_email = Worksheets("Main data").Cells(row_num, 4)
    inflow_date = Format(Worksheets("Main data").Cells(row_num, 9), "dd mmmm yyyy") + ", " + WeekdayName(Weekday(Worksheets("Main data").Cells(row_num, 9)))
    buddy_name = Worksheets("Main data").Cells(row_num, 22)
    buddy_hp = Worksheets("Main data").Cells(row_num, 23)
    supervisor_name = Worksheets("Main data").Cells(row_num, 24)
    supervisor_hp = Worksheets("Main data").Cells(row_num, 25)
    
    
    ' Assign the values for the email attachments
    hcs_letter = "https://sgdcs.sgnet.gov.sg/sites/PMO-SNDGO/HCD/_layouts/15/WopiFrame2.aspx?sourcedoc=%7B66CD612A-EFBC-4CEB-837B-DD72955E919C%7D&file=Welcome%20Letter_HCS.pdf&action=default"
    ps_values = "https://sgdcs.sgnet.gov.sg/sites/PMO-SNDGO/HCD/_layouts/15/WopiFrame2.aspx?sourcedoc=%7B094E4516-0F87-4EC7-A360-C48A630233B1%7D&file=20201124_Officers%20Guide%20to%20Public%20Service%20Values.pdf&action=default"
    fr_guide = candidates_folder_path & "\1. Resource folder for VBA code (do not edit)\welcome email\Guide to set up FR.pdf"
    sundae_edm_guide = candidates_folder_path & "\1. Resource folder for VBA code (do not edit)\welcome email\SUNDAE_EDM_v3.pptx"
    'cyber_security_guide = getAllCandidatesFolderPath() & "\1. Resource folder for VBA code (do not edit)\welcome email\2023 Cybersecurity & Data Protection User Guide_100323.pdf"
    img1 = candidates_folder_path & "\1. Resource folder for VBA code (do not edit)\welcome email\welcome_img1.png"
    img2 = candidates_folder_path & "\1. Resource folder for VBA code (do not edit)\welcome email\welcome_img2.png"
    img3 = candidates_folder_path & "\1. Resource folder for VBA code (do not edit)\welcome email\welcome_img3.png"
    img4 = candidates_folder_path & "\1. Resource folder for VBA code (do not edit)\welcome email\welcome_img4.png"
    img5 = candidates_folder_path & "\1. Resource folder for VBA code (do not edit)\welcome email\welcome_img5.png"
    img6 = candidates_folder_path & "\1. Resource folder for VBA code (do not edit)\welcome email\welcome_img6.png"
    
    
    
    'define the email subject
    email_subject = "Warm Welcome to the SNDGO Family!"

    'Create a new instance of Outlook
    Set outlookApp = CreateObject("Outlook.Application")
    
    'Create a new email
    Set outlookMail = outlookApp.CreateItem(0)
    
    
    
    
    content1 = "<b>On Your First Physical Day reporting to Office</b><br><br>" + _
               "On<b> " + inflow_date + "</b>, please report to SNDGO's office at <b>9.30am</b> at:<br>109 North Bridge Road<br>#06-01 Funan, O2 Office<br>Singapore 179097<br><br>" + _
               "Please use the <b>QR code from Funan (will be sent separately to you) to access the automated gantry at Level 1.</b> You would need to increase the screen brightness level to maximum on your phone before scanning the QR code.<br><br>" + _
               "Please enter via <b><u>O2 Lobby, located opposite McDonalds.</u></b> Upon reaching Level 6, you can call your supervisor, " + supervisor_name + " (" + supervisor_hp + "),<br>" + _
               "or your buddy, " + buddy_name + " (" + buddy_hp + ").<br><br>" + _
               "Please also look for Jorin and Melissa to collect/setup your IT equipment and collect your access pass.<br><br>" + _
               "After you gain access to your email, you will need to set up your facial recognition for lobby access with the attached pdf guide to set up facial recognition (refer to page 3 and 5). "

    content2 = "<b>Head Civil Service's (HCS) Welcome Letter and Guide to Public Service Values </b><br><br>" + _
               "HCS has penned a Welcome Letter for all new officers joining the Public Service. Do read through the letter <u><a href=" + hcs_letter + ">here</a></u> to learn about the purpose of the Public Service.<br><br>" + _
               "You can also refer to this <u><a href=" + ps_values + ">resource guide</a></u> to learn about the values and conduct we are guided by. "
               
    content3 = "<b>Agency Intranet</b><br><br>" + _
               "Our <u><a href=""https://sgdcs.sgnet.gov.sg/sites/PMO-SNDGO/SitePages/SNDGO_Home.aspx"">intranet</a></u> which contains useful information on SNDGO, such as <u><a href=""https://sgdcs.sgnet.gov.sg/sites/PMO-SNDGO/SitePages/About_SNDGO.aspx"">organisation charts</a></u>, staff updates and events. You can also find<br>" + _
               "detailed guidelines and policies pertaining to HR, IT and Finance matters. Here are some links that you may find useful:<br><br>" + _
               "<ul><li>On the <u><a href=""https://intranet.hrp.gov.sg"">HRP system</a></u>, you can update your personal particulars, apply for leave, complete your appraisal form, submit claims, view your payslip and more. Your account will be set up within your first 2 weeks  and is accessible either via a " + _
               "single sign-on on the intranet with your work laptop, or via SingPass on the internet on your personal mobile device. You will also be able to perform selected functions via the Workpal app on your mobile device. </li>" + _
               "<li><u><a href=""https://intranet.mof.gov.sg/portal/IM.aspx"">Government Instruction Manual (IM)</a></u> on appointments, benefits, compensation, etc. Circulars can also be found here.</li>" + _
               "<li>Data is an important asset that enables us to carry out effective reviews of our policies and processes to better serve our " + _
               "citizens. The data that we handle in our daily work may include data classified as ""Confidential"" or include sensitive data of " + _
               "individuals. Any form of data compromise or misuse may cause damage to SNDGO, the individuals or even national " + _
               "interests! As public officers, we have the obligation to safeguard all official data in our possession against unauthorised " + _
               "disclosure and any form of data misuse. Check out the <u><a href=""https://sgdcs.sgnet.gov.sg/sites/PMO-SNDGO/HCD/_layouts/15/WopiFrame2.aspx?sourcedoc=%7B943EBC2B-F0E8-468E-8DAB-2263C972DD51%7D&file=01_%20Handbook%20for%20GPOs_23%20July.pdf&action=default"">Data Governance Page</a></u> to access important information that you need to know for the daily handling of data or documents.</li>" + _
               "<li><u><a href=""https://resource.digitalworkplace.gov.sg"">Room Booking System</a></u> (also on WorkPal) for the booking of meeting rooms for meetings, events, functions in Funan.</li>" + _
               "<li>Transport expenses can be claimed for official work trips, please refer to this <u><a href=""https://sgdcs.sgnet.gov.sg/sites/PMO-SNDGO/HCD/_layouts/15/WopiFrame2.aspx?sourcedoc=%7B16F5EB15-CEDD-4ED0-884B-5DF814A2C612%7D&file=Best%20Practices%20Guidelines%20for%20travel%20and%20transport_Apr%202019.pdf&action=default"">link</a></u> for more details. Officers can also book work-related rides via the <u><a href=""https://sgdcs.sgnet.gov.sg/sites/PMO-SNDGO/PPD/FR_subsite/_layouts/15/WopiFrame2.aspx?sourcedoc=%7BE681DF7D-D837-438C-A571-CC0E4D220AAC%7D&file=Grab-for-Business-Deck-Gov_Mar-19-Final-Onboarding-Corp-Billing-Only.pptx&action=default"">Grab for Work</a></u> corporate account.</li>" + _
               "<li>Officers are provided a telecommuting subsidy of up to $200 to purchase approved IT equipment (e.g. keyboard, computer mouse and mouse pad, earphones/headsets) to support them as they work from home. Please refer to this <u><a href=""https://sgdcs.sgnet.gov.sg/sites/PMO-SNDGO/SiteAssets/SitePages/Employee_Central/FAQ%20for%20SNDGO%20BYO@Home%20Initative%20(July%202022).pdf"">link</a></u> for more details.</li></ul>"
    
    
    
    content4 = "<b>Learning & Development</b><br><br>" + _
               "In SNDGO, we take pride in our people and believe in developing officers to their full potential. Check out the <u><a href=""https://intranet.hrp.gov.sg"">HRP system</a></u> learning and development page for resources to support your learning and transition into your role!<br><br>" + _
               "You should complete the newbies' e-Learning, comprising information on knowledge management, basic procurement etc., which will be useful for your work. Do try to complete this within your first month in SNDGO to help facilitate an effective transition." + _
               "For more training-related information, you can approach our Training Coordinator, Sharanyaa (copied in this email) for assistance.<br><br>" + _
               "You can also manage your career planning better with career coaching. Head over to the <u><a href=""https://gccprod.sharepoint.com/sites/psd-EveryOfficerMatters/SitePages/Programmes-%26-Resources.aspx"">Every Officer Matters</a></u> page, to find out more about career coaching and how you can sign up for a <u><a href=""https://gccprod.sharepoint.com/sites/psd-EveryOfficerMatters/SitePages/CareerPlanning.aspx"">career coaching</a></u> session. "

    
    
    
    content5 = "<b>2023 CYBERSECURITY & DATA PROTECTION QUIZ</b><br><br>" + _
               "New officers would need to complete both the e-learning modules and pass the quiz within 3 months of joining service.<br><br>" + _
               "<ol type=""a""><li>BDLCD1: Cybersecurity (30 mins)</li><li>BDLCD2: Data Protection (30 mins)</li><li>BDLCD3: Incident Management (30 mins)</li>" + _
               "<li>BDLQ1: 2023 Cybersecurity & Data Protection Quiz (1 hour) (Mandatory)</li></ol>"


    
    content6 = "<b>Submission of Declarations in <u><a href=""https://intranet.hrp.gov.sg"">HRP system</a></u></b><br><br>" + _
               "Please submit the following declarations via <u><a href=""https://intranet.hrp.gov.sg"">HRP system</a></u> once your account is set up within your first 2 weeks: <br><br>" + _
               "<ol type=""a""><li>Financial Indebtedness</li><li>Property/Land</li><li>Shares</li>" + _
               "<li>Interest in Business Firms</li></ol>"
    
    
    email_body = "Hi " + full_name + ",<br><br>Welcome to SNDGO (and Government Chief Digital Technology Office), and we hope you are having a good start with us! We would like to share some useful information to get you started on your journey with us. You may wish to forward this email to your PMO email address once it is ready to access the intranet links.<br>"
    
    table = "<p><strong><u>Welcome info:</u></strong></p><table style=""border-collapse: collapse; border: 1px solid black;""><tbody><tr>" + _
            "<td style=""border: 1px solid black; padding: 5px;""><p><img src=""" + img1 + """ alt=""Image description"" style=""max-width: 200px; max-height: 200px;""></p></td><td style=""border: 1px solid black;"" valign=""top"" align=""left""><p>" + content1 + "</p></td></tr><tr>" + _
            "<td style=""border: 1px solid black; padding: 5px;""><p><img src=""" + img2 + """ alt=""Image description""></p></td><td style=""border: 1px solid black;"" valign=""top"" align=""left""><p>" + content2 + "</p></td></tr>" + _
            "<tr><td style=""border: 1px solid black; padding: 5px;""><p><img src=""" + img3 + """ alt=""Image description""></p></td><td style=""border: 1px solid black;"" valign=""top"" align=""left""><p>" + content3 + "</p></td></tr>" + _
            "<tr><td style=""border: 1px solid black; padding: 5px;""><p><img src=""" + img4 + """ alt=""Image description""></p></td><td style=""border: 1px solid black;""><p>" + content4 + "</p></td></tr>" + _
            "<td style=""border: 1px solid black; padding: 5px;"" ><p><img src=""" + img5 + """ alt=""Image description""></p></td><td style=""border: 1px solid black;""><p>" + content5 + "</p></td></tr>" + _
            "<td style=""border: 1px solid black; padding: 5px; ""><p><img src=""" + img6 + """ alt=""Image description""></p></td><td style=""border: 1px solid black;""><p>" + content6 + "</p></td></tr></tbody></table>"
    
    
    email_body = email_body + table + "<br><br>Feel free to approach your me if you have any HR-related queries. We hope that you will have a meaningful career journey in SNDGO!"
    
    
    'Add details for the email
    With outlookMail
        .To = personal_email
        .subject = email_subject
        .HTMLBody = email_body
        .Attachments.Add fr_guide
        .Attachments.Add sundae_edm_guide
'        .Attachments.Add cyber_security_guide
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
    send_welcome_email = email_subject


End Function


Function GetOutlookSignature() As String
    Dim outlookApp As Object
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim signatureFilePath As String
    Dim signatureContent As String
    
    On Error Resume Next ' Error handling for Outlook not being open
    
    ' Get the Outlook Application
    Set outlookApp = GetObject(, "Outlook.Application")
    On Error GoTo 0 ' Reset error handling
    
    ' Check if Outlook is running
    If outlookApp Is Nothing Then
        GetOutlookSignature = ""
        Exit Function
    End If
    
    ' Get the Word editor from Outlook (to access the signature)
    Set wordApp = outlookApp.GetNamespace("MAPI").GetDefaultFolder(olFolderDrafts).GetInspector.WordEditor
    
    ' Set the signature file path (default signature)
    signatureFilePath = Environ("APPDATA") & "\Microsoft\Signatures\"
    signatureFilePath = signatureFilePath & Dir(signatureFilePath & "*.htm*")
    
    ' If a signature file is found, read its content
    If signatureFilePath <> "" Then
        Set wordDoc = wordApp.Documents.Open(signatureFilePath)
        signatureContent = wordDoc.Content.Text
        wordDoc.Close SaveChanges:=False
        GetOutlookSignature = signatureContent
    Else
        GetOutlookSignature = ""
    End If
    
    ' Clean up objects
    Set wordDoc = Nothing
    Set wordApp = Nothing
    Set outlookApp = Nothing
End Function


