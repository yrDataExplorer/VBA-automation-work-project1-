Option Explicit

' This module contains the code for searching and extracting of data from email replies for stages 1-3


Function ExtractStringAfterKeyword(inputString As String, keyword As String) As String
    Dim startPos As Long
    Dim endPos As Long
    
    ' Find the position of the keyword in the input string
    startPos = InStr(1, inputString, keyword, vbTextCompare)
    
    ' If the keyword is found, extract the string to the right of it up to the end of the line
    If startPos > 0 Then
        ' Find the position of the end of the line (vbNewLine or vbCrLf)
        endPos = InStr(startPos, inputString, vbNewLine, vbTextCompare)
        If endPos = 0 Then
            endPos = InStr(startPos, inputString, vbCrLf, vbTextCompare)
        End If
        
        ' Extract the desired substring
        If endPos > 0 Then
            ExtractStringAfterKeyword = Mid(inputString, startPos + Len(keyword), endPos - (startPos + Len(keyword)))
        Else
            ExtractStringAfterKeyword = Mid(inputString, startPos + Len(keyword))
        End If
        
    End If
    
End Function






Function search_pre_offer_email(row_num As Integer) As Boolean

    Dim outlookApp As Object
    Dim outlookNamespace As Object
    Dim inboxFolder As Object
    Dim mailItem As Object, mailAttachment As Object
    Dim searchSubject As String
    Dim search_result As String
    
    
    'Ensure that the main worksheet for candidate data in the database file is active
    database.Worksheets("Main data").Activate
    
    
    ' define the unique ID variable and value
    Dim unique_ID As String
    unique_ID = Worksheets("Main data").Cells(row_num, 1)
    
    ' define the variables to store personal particulars data
    Dim fullname As String, nric As String, contact_number As String, nationality As String, race As String, religion As String, e_name As String, e_contact As String
    Dim e_address As String, e_relationship As String
    
    
    ' Search term to look for in the email subject
    searchSubject = "SNG Personal Particulars Form (For Applicant)" ' Replace with the string you want to search for
    
    ' Create a new instance of Outlook Application
    Set outlookApp = CreateObject("Outlook.Application")
    ' Get the MAPI namespace
    Set outlookNamespace = outlookApp.GetNamespace("MAPI")
    ' Get the inbox folder
    Set inboxFolder = outlookNamespace.GetDefaultFolder(6).Folders("Onboarding automation emails")

    
    For Each mailItem In inboxFolder.Items
        If InStr(1, mailItem.subject, searchSubject, vbTextCompare) > 0 And _
           InStr(1, mailItem.Body, unique_ID, vbTextCompare) > 0 Then
            
            fullname = WorksheetFunction.Trim(ExtractStringAfterKeyword(mailItem.Body, "Full Name (as in NRIC)"))
            nric = WorksheetFunction.Trim(ExtractStringAfterKeyword(mailItem.Body, "NRIC No."))
            contact_number = WorksheetFunction.Trim(ExtractStringAfterKeyword(mailItem.Body, "Mobile No."))
            nationality = WorksheetFunction.Trim(ExtractStringAfterKeyword(mailItem.Body, "Nationality"))
            race = WorksheetFunction.Trim(ExtractStringAfterKeyword(mailItem.Body, "Race"))
            religion = WorksheetFunction.Trim(ExtractStringAfterKeyword(mailItem.Body, "Religion"))
            e_name = WorksheetFunction.Trim(ExtractStringAfterKeyword(mailItem.Body, "Full Name of Emergency Contact Person"))
            e_contact = WorksheetFunction.Trim(ExtractStringAfterKeyword(mailItem.Body, "Contact No."))
            e_address = WorksheetFunction.Trim(ExtractStringAfterKeyword(mailItem.Body, "Residential Address"))
            e_relationship = WorksheetFunction.Trim(ExtractStringAfterKeyword(mailItem.Body, "Relationship to Applicant"))
            
            
            Worksheets("Main data").Cells(row_num, 10).Value = fullname
            Worksheets("Main data").Cells(row_num, 11).Value = nric
            Worksheets("Main data").Cells(row_num, 12).Value = contact_number
            Worksheets("Main data").Cells(row_num, 13).Value = nationality
            Worksheets("Main data").Cells(row_num, 14).Value = race
            Worksheets("Main data").Cells(row_num, 15).Value = religion
            Worksheets("Main data").Cells(row_num, 16).Value = e_name
            Worksheets("Main data").Cells(row_num, 17).Value = e_contact
            Worksheets("Main data").Cells(row_num, 18).Value = e_address
            Worksheets("Main data").Cells(row_num, 19).Value = e_relationship
            
            
            For Each mailAttachment In mailItem.Attachments
                ' Save the attachment to the specified folder
                ' MsgBox candidate_folder & "\" & mailAttachment.Filename
                mailAttachment.SaveAsFile specific_candidate_folder & "\" & mailAttachment.Filename
                
            Next mailAttachment
            
            search_result = "Found"
            
        End If
        
    Next mailItem
    
    If search_result = "Found" Then
        
        Call edit_personal_particulars_form(pp_form_candidate, fullname, nric, contact_number, nationality, race, religion, e_name, e_contact, e_address, e_relationship)
        
        search_pre_offer_email = True
        
    Else
    
        search_pre_offer_email = False
    
    End If
    

End Function


Sub edit_personal_particulars_form(filepath As String, fullname As String, nric As String, contact_number As String, nationality As String, race As String, religion As String, e_name As String, e_contact As String, e_address As String, e_relationship As String)

    Dim objWordApp As Object
    Dim objWordDoc As Object
    Dim objWordTable1 As Object, objWordTable2 As Object

    ' Open Word application and document
    Set objWordApp = CreateObject("Word.Application")
    objWordApp.Visible = True
    Set objWordDoc = objWordApp.Documents.Open(filepath)

    Dim docPath As String
    docPath = objWordDoc.fullname
    'MsgBox for testing and debugging
    'MsgBox docPath
    

    
    ' Reference the first table in the Word document
    Set objWordTable1 = objWordDoc.Tables(1)
    
    ' enter the personal particulars info into the word document form for the first table

    objWordTable1.cell(1, 2).Range.Text = fullname ' Place data in first column, first row
    objWordTable1.cell(2, 2).Range.Text = nric ' Place data in first column, second row
    objWordTable1.cell(3, 2).Range.Text = contact_number ' Place data in first column, third row
    objWordTable1.cell(1, 4).Range.Text = nationality ' Place data in second column, first row
    objWordTable1.cell(2, 4).Range.Text = race ' Place data in second column, second row
    objWordTable1.cell(3, 4).Range.Text = religion ' Place data in second column, third row


    ' Reference the second table in the Word document
    Set objWordTable2 = objWordDoc.Tables(2)

    
    ' enter the personal particulars info into the word document form for the second table

    objWordTable2.cell(1, 2).Range.Text = e_name ' Place data in first column, first row
    objWordTable2.cell(2, 2).Range.Text = e_contact ' Place data in first column, second row
    objWordTable2.cell(3, 2).Range.Text = e_address ' Place data in first column, third row
    objWordTable2.cell(4, 2).Range.Text = e_relationship ' Place data in second column, first row
  
    
 
    objWordDoc.Close SaveChanges:=True
    objWordApp.Quit
    
'    On Error GoTo badFilePath
'    objWordDoc.SaveCopy2 savePath1
'    objWordDoc.Close SaveChanges:=False
'    Exit Sub
'
'badFilePath:
'    objWordDoc.SaveCopyAs savePath2
'    objWordDoc.Close SaveChanges:=False


End Sub



Function search_onboarding_prep_email(row_num As Integer) As Boolean


    'Ensure that the main worksheet for candidate data in the database file is active
    database.Worksheets("Main data").Activate

    Dim inflow_mode As String, search_result As Boolean
    inflow_mode = Worksheets("Main data").Cells(row_num, 8)
    
    
    ' Define the arrays to store the different types of inflow mode for each category of inflow
    Dim new_hire_array As Variant
    Dim secondment_array As Variant
    
    
    new_hire_array = Array("Open Recruitment on MX", "MOPO", _
    "PSLP GP 1st", "PSLP GP 2nd", "PSLP GP 3rd", "PSLP GP - Eng", "PSLP SP")
    
    secondment_array = Array("Secondment", "Transfer", "IO Posting", "AO Posting", "LS Posting", _
                            "Open Recruitment on GovTech", "Open Recruitment on OGP")


    ' Check for candidate's inflow mode in order to determine the correct onboarding email to be sent
    Dim Count As Integer
    For Count = LBound(new_hire_array) To UBound(secondment_array)
        If new_hire_array(Count) = inflow_mode Then
            
            ' If the candidate joined via open recruitment, send onboarding prep form for open recruitment
            Debug.Print "Searching for onboarding prep email reply for open recuitment to candidate"
            
            search_result = search_onboarding_prep_new_hire(row_num)

        ElseIf secondment_array(Count) = inflow_mode Then
            
            ' If the candidate joined via secondment/posting, send onboarding prep form for secondees
            Debug.Print "Searching for onboarding prep email reply for secondees/posting to candidate"

            search_result = search_onboarding_prep_email_secondees(row_num)

        End If
    
    Next Count
    
    
    search_onboarding_prep_email = search_result

End Function




Function search_onboarding_prep_new_hire(row_num As Integer) As Boolean


    'Ensure that the main worksheet for candidate data in the database file is active
    database.Worksheets("Main data").Activate

    Dim outlookApp As Object
    Dim outlookNamespace As Object
    Dim inboxFolder As Object
    Dim mailItem As Object, mailAttachment As Object
    Dim searchSubject As String
    Dim search_result As String
    
    ' define the unique ID variable and value
    Dim unique_ID As String
    unique_ID = Worksheets("Main data").Cells(row_num, 1)
    
    
    ' Search term to look for in the email subject
    searchSubject = "Onboarding Preparation (For New Hire)" ' Replace with the string you want to search for
    
    ' Create a new instance of Outlook Application
    Set outlookApp = CreateObject("Outlook.Application")
    ' Get the MAPI namespace
    Set outlookNamespace = outlookApp.GetNamespace("MAPI")
    ' Get the inbox folder
    Set inboxFolder = outlookNamespace.GetDefaultFolder(6).Folders("Onboarding automation emails")

    
    For Each mailItem In inboxFolder.Items
    
        'Debug.Print mailItem.Subject
        
        If InStr(1, mailItem.subject, searchSubject, vbTextCompare) > 0 And _
           InStr(1, mailItem.Body, unique_ID, vbTextCompare) > 0 Then
           
            Worksheets("Main data").Cells(row_num, 20).Value = WorksheetFunction.Trim(ExtractStringAfterKeyword(mailItem.Body, "Self-intro message:"))
            
            Worksheets("Main data").Cells(row_num, 21).Value = WorksheetFunction.Trim(ExtractStringAfterKeyword(mailItem.Body, "Mobile number for Grab corporate account"))
            

            
            For Each mailAttachment In mailItem.Attachments
                ' Save the attachment to the specified folder
                ' MsgBox candidate_folder & "\" & mailAttachment.Filename
                mailAttachment.SaveAsFile specific_candidate_folder & "\" & mailAttachment.Filename
                
            Next mailAttachment
            
            search_result = "Found"
                
            
        End If
        
    Next mailItem
    
    
    If search_result = "Found" Then
        
        
        search_onboarding_prep_new_hire = True
        
    Else
    
        search_onboarding_prep_new_hire = False
    
    End If

End Function




Function search_onboarding_prep_email_secondees(row_num As Integer) As Boolean

    'Ensure that the main worksheet for candidate data in the database file is active
    database.Worksheets("Main data").Activate
    
    Dim outlookApp As Object
    Dim outlookNamespace As Object
    Dim inboxFolder As Object
    Dim mailItem As Object, mailAttachment As Object
    Dim searchSubject As String
    Dim search_result As String
    
    ' define the unique ID variable and value
    Dim unique_ID As String
    unique_ID = Worksheets("Main data").Cells(row_num, 1)
    
    
    ' Search term to look for in the email subject
    searchSubject = "Onboarding Preparation (For Secondees)" ' Replace with the string you want to search for
    
    ' Create a new instance of Outlook Application
    Set outlookApp = CreateObject("Outlook.Application")
    ' Get the MAPI namespace
    Set outlookNamespace = outlookApp.GetNamespace("MAPI")
    ' Get the inbox folder
    Set inboxFolder = outlookNamespace.GetDefaultFolder(6).Folders("Onboarding automation emails")

    
    For Each mailItem In inboxFolder.Items
        If InStr(1, mailItem.subject, searchSubject, vbTextCompare) > 0 And _
           InStr(1, mailItem.Body, unique_ID, vbTextCompare) > 0 Then
            
            Worksheets("Main data").Cells(row_num, 10).Value = WorksheetFunction.Trim(ExtractStringAfterKeyword(mailItem.Body, "Full Name (as in NRIC)"))
            Worksheets("Main data").Cells(row_num, 11).Value = WorksheetFunction.Trim(ExtractStringAfterKeyword(mailItem.Body, "NRIC"))
            Worksheets("Main data").Cells(row_num, 12).Value = WorksheetFunction.Trim(ExtractStringAfterKeyword(mailItem.Body, "Mobile No."))
            Worksheets("Main data").Cells(row_num, 16).Value = WorksheetFunction.Trim(ExtractStringAfterKeyword(mailItem.Body, "Full Name of Emergency Contact Person"))
            Worksheets("Main data").Cells(row_num, 17).Value = WorksheetFunction.Trim(ExtractStringAfterKeyword(mailItem.Body, "Contact No."))
            Worksheets("Main data").Cells(row_num, 18).Value = WorksheetFunction.Trim(ExtractStringAfterKeyword(mailItem.Body, "Residential Address"))
            Worksheets("Main data").Cells(row_num, 19).Value = WorksheetFunction.Trim(ExtractStringAfterKeyword(mailItem.Body, "Relationship to Officer"))
            Worksheets("Main data").Cells(row_num, 20).Value = WorksheetFunction.Trim(ExtractStringAfterKeyword(mailItem.Body, "Self-intro message:"))
            Worksheets("Main data").Cells(row_num, 21).Value = WorksheetFunction.Trim(ExtractStringAfterKeyword(mailItem.Body, "Mobile number for Grab corporate account"))
            
            
            For Each mailAttachment In mailItem.Attachments
                ' Save the attachment to the specified folder
                ' MsgBox candidate_folder & "\" & mailAttachment.Filename
                mailAttachment.SaveAsFile specific_candidate_folder & "\" & mailAttachment.Filename
                
            Next mailAttachment
            
            search_result = "Found"
            
        End If
        
    Next mailItem
    
    If search_result = "Found" Then
        
        
        search_onboarding_prep_email_secondees = True
        
    Else
    
        search_onboarding_prep_email_secondees = False
    
    End If
    


End Function


Function search_supervisor_buddy_inputs(row_num As Integer) As Boolean

    'Ensure that the main worksheet for candidate data in the database file is active
    database.Worksheets("Main data").Activate

    Dim outlookApp As Object
    Dim outlookNamespace As Object
    Dim inboxFolder As Object
    Dim mailItem As Object, mailAttachment As Object
    Dim searchSubject As String
    Dim search_result As String
    
    
    ' define the unique ID variable and value
    Dim unique_ID As String
    unique_ID = Worksheets("Main data").Cells(row_num, 1)
    

    ' Search term to look for in the email subject
    searchSubject = "Onboarding Preparation (For Directorate's Input)" ' Replace with the string you want to search for
    
    ' Create a new instance of Outlook Application
    Set outlookApp = CreateObject("Outlook.Application")
    ' Get the MAPI namespace
    Set outlookNamespace = outlookApp.GetNamespace("MAPI")
    ' Get the inbox folder
    Set inboxFolder = outlookNamespace.GetDefaultFolder(6).Folders("Onboarding automation emails")

    
    For Each mailItem In inboxFolder.Items
        If InStr(1, mailItem.subject, searchSubject, vbTextCompare) > 0 And _
           InStr(1, mailItem.Body, unique_ID, vbTextCompare) > 0 Then
            
            Worksheets("Main data").Cells(row_num, 22).Value = WorksheetFunction.Trim(ExtractStringAfterKeyword(mailItem.Body, "Assigned Buddy's Name"))
            Worksheets("Main data").Cells(row_num, 23).Value = WorksheetFunction.Trim(ExtractStringAfterKeyword(mailItem.Body, "Assigned Buddy's Contact No."))
            Worksheets("Main data").Cells(row_num, 24).Value = WorksheetFunction.Trim(ExtractStringAfterKeyword(mailItem.Body, "Assigned Supervisor's Name"))
            Worksheets("Main data").Cells(row_num, 25).Value = WorksheetFunction.Trim(ExtractStringAfterKeyword(mailItem.Body, "Assigned Supervisor's Contact No."))
            Worksheets("Main data").Cells(row_num, 26).Value = WorksheetFunction.Trim(ExtractStringAfterKeyword(mailItem.Body, "Reporting officer of?"))
            Worksheets("Main data").Cells(row_num, 27).Value = WorksheetFunction.Trim(ExtractStringAfterKeyword(mailItem.Body, "Require access to secret info? (i.e. Cat 1 clearance)"))
            Worksheets("Main data").Cells(row_num, 28).Value = WorksheetFunction.Trim(ExtractStringAfterKeyword(mailItem.Body, "Require Secured Email access? (i.e. PKI card/ S-Notebook)"))
            
            

            
            For Each mailAttachment In mailItem.Attachments
                ' Save the attachment to the specified folder
                ' MsgBox candidate_folder & "\" & mailAttachment.Filename
                mailAttachment.SaveAsFile specific_candidate_folder & "\" & mailAttachment.Filename
                
            Next mailAttachment
            
            search_result = "Found"
                
            
        End If
        
    Next mailItem
    
    
    

    If search_result = "Found" Then
        
        
        search_supervisor_buddy_inputs = True
        
    Else
    
        search_supervisor_buddy_inputs = False
    
    End If

End Function



