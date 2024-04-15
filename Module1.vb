Option Explicit
Public candidates_folder_path As String, specific_candidate_folder As String, database As Workbook, pp_form As String, pp_form_candidate As String


' This module contains the code for the main process and the process to evaluate the completion status for each stage

Function GetCandidateFolderPath()

    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = Application.DefaultFilePath & " \ "
        .Title = "HR Onboarding automation: Select the candidate folder directory"
        .Show
        
        If .SelectedItems.Count = 0 Then
        
            GetCandidateFolderPath = "Canceled"
            
        Else
        
            'MsgBox .SelectedItems(1)
            GetCandidateFolderPath = .SelectedItems(1)
            
        End If
        
    End With

End Function



Sub main_process()

    'define the candidate folder path which is a public variable
    candidates_folder_path = GetCandidateFolderPath()
    If candidates_folder_path = "Canceled" Then
        Exit Sub
    End If


    'define the excel database_path
    Dim database_path As String
    database_path = candidates_folder_path & "\Onboarding automation database.xlsx"
    
    Set database = Workbooks.Open(database_path)
    
    UserForm1.Show vbModeless



End Sub




Sub automation_main_process()

    Application.ScreenUpdating = False
    
    'Ensure that the main worksheet for candidate data in the database file is active
    database.Worksheets("Main data").Activate

    ' Find the last row number for the updated range of data to work with
    Dim last_row As Integer
    last_row = Worksheets("Main data").Cells(Rows.Count, "B").End(xlUp).Row 'get last filled row of Column B
    'Debug.Print "Main process running. Last row number found for the whole dataset: " & Str(last_row)
    
    
    ' Define the variables for keeping track of each iteration in the for loop to evaluate if all stages have been completed for each row
    Dim row_num As Integer
    Dim config As String ' Define the variable to store the options for the final message box
    
    
    ' clear the activity log
    UserForm1.TextBox_log_activity.Text = ""
    
    
    'Ensure that the worksheet for onboarding progress tracker in the database file is active
    database.Worksheets("Progress tracker").Activate
    
    ' if the last row is more than 3, it implies that there is at least one candidate entered
    If last_row > 3 Then
    
    
'            'Create a copy of the form in a variable
'            Set ProgressIndicator = New UserForm2
'            'Show ProgressIndicator in modeless state
'            ProgressIndicator.Show vbModeless
'
'            ' Define the variables to increment during each step and update the progress indicator bar
'            Dim counter As Integer, total As Integer, PctDone As Double
'            total = last_row * 9
'            counter = 1
    
        For row_num = 4 To last_row
        
        
            'to add a test to see if all stages has been completed for a particular candidate and
            If check_stage_completion(row_num, 1, 9) Then
                
                UserForm1.TextBox_log_activity.Text = UserForm1.TextBox_log_activity.Value & "All stages already completed for " & Cells(row_num, 10) & "," & Cells(row_num, 11) & vbNewLine
                GoTo ContinueLoop 'break out of the current loop to continue with the next loop
                
            End If
            
            
        
        
            'define the specific candidate folder path for the candidate
            specific_candidate_folder = define_specific_candidate_folder(row_num)
            'define the file path for the general pp_form and the pp_form for the specific candidate
            pp_form = candidates_folder_path + "\Personal and Bank Particulars Form.docx"
            pp_form_candidate = specific_candidate_folder + "\" + "Personal and Bank Particulars Form.docx"
            
            database.Worksheets("Progress tracker").Activate 'Ensure that the worksheet for onboarding progress tracker in the database file is active
            
            
            'Check if all steps in stage 1 is completed
            If Not check_stage_completion(row_num, 1, 2) Then 'if not all steps are completed, check and run automation for remaining steps in stage 1
                
                
                If Cells(row_num, 1) = "N" Then 'if pre-offer form has not yet been created, call function to create pre-offer form
                    
                    
                    Dim check_subject As String
                    check_subject = Module2.create_pre_offer_email(row_num)
                    
                    'Ensure that the worksheet for onboarding progress tracker in the database file is active
                    database.Worksheets("Progress tracker").Activate
                    
                    If IsDraftEmailExist(check_subject) = True Then 'check if pre-offer email has been successfully created
                    
                        Cells(row_num, 1) = "Y" 'Update the status for step 1 in stage 1 if successfully created
                        
                        UserForm1.TextBox_log_activity.Text = UserForm1.TextBox_log_activity.Value & "Created pre-offer email for " & Cells(row_num, 10) & "," & Cells(row_num, 11) & vbNewLine
                        
                        Call update_timestamp(row_num, 1)
                        
                        
                        
                    End If
                    
                End If
                
'                'update counter and progress indicator bar
'                counter = counter + 1
'                PctDone = counter / total
'                Call update_progress(PctDone)
                
                
                If Cells(row_num, 2) = "N" Then 'if candidate folder and personal particulars form has not been created, call function to create candidate folder and documents
                
                    Debug.Print "Calling function to create candidate folder and documents"
                         
                    'check if candidate folder and documents has been created successfully
                    If Module2.create_candidate_folder_documents(row_num) Then 'function will return True if successful
                    
                        'Ensure that the worksheet for onboarding progress tracker in the database file is active
                        database.Worksheets("Progress tracker").Activate
                        Cells(row_num, 2) = "Y" 'Update the status for step 2 in stage 1 if successfully created
                        
                        UserForm1.TextBox_log_activity.Text = UserForm1.TextBox_log_activity.Value & "Created candidate folder for for " & Cells(row_num, 10) & "," & Cells(row_num, 11) & vbNewLine
                        
                        Call update_timestamp(row_num, 2)
                        
                        
                        
                    End If
                    
                
                End If
                
                If Cells(row_num, 3) = "NA" Then
                    
                    GoTo ForNA
                
                End If

                
            Else
                
                'Check if all steps in stage 2 is completed
                If Not check_stage_completion(row_num, 3, 4) Then 'if not all steps are completed, check and run automation for remaining steps in stage 2
                    
                    If Cells(row_num, 3) = "N" Then 'if pre-offer form email has not be received and extracted
                    
                        'call function to search for pre-offer form reply and check if pre-offer reply has been found and extracted successfully
                        If Module3.search_pre_offer_email(row_num) Then 'function will return True if successful
                    
                            'Ensure that the worksheet for onboarding progress tracker in the database file is active
                            database.Worksheets("Progress tracker").Activate
                            Cells(row_num, 3) = "Y" 'Update the status for step 1 in stage 2 if successfully created
                            
                            UserForm1.TextBox_log_activity.Text = UserForm1.TextBox_log_activity.Value & "Received pre-offer form reply for " & Cells(row_num, 10) & "," & Cells(row_num, 11) & vbNewLine
                            
                            Call update_timestamp(row_num, 3)
                        
                        End If
                        
                    
                    End If
                    


ForNA:
                    If Cells(row_num, 4) = "N" And check_inflow_date(row_num) Then
                    
                        'call function to send emails and create onboarding prep form
                        If Module4.stage2_automation(row_num) Then 'function will return True if successful
                        
                            'Ensure that the worksheet for onboarding progress tracker in the database file is active
                            database.Worksheets("Progress tracker").Activate
                            Cells(row_num, 4) = "Y" 'Update the status for step 2 in stage 2 if successfully completed
                            
                            UserForm1.TextBox_log_activity.Text = UserForm1.TextBox_log_activity.Value & "Created draft emails to ITD, Dexun and onboarding prep email for " & Cells(row_num, 10) & "," & Cells(row_num, 11) & vbNewLine
                            
                            Call update_timestamp(row_num, 4)
                            
                            
                        End If
                    
                    End If
                

                    
                
                Else
                    
                    
                    'check if all steps in stage 3 is completed
                    If Not check_stage_completion(row_num, 5, 7) Then 'if not all steps are completed, check and run automation for remaining steps in stage 3
                        
                        
                        If Cells(row_num, 5) = "N" Then 'if onboarding prep form reply has not be received and extracted
                    
                            'call function to search for onboarding prep email reply and check if onboarding prep form reply has been found and extracted successfully
                            If Module3.search_onboarding_prep_email(row_num) Then 'function will return True if successful
                    
                                'Ensure that the worksheet for onboarding progress tracker in the database file is active
                                database.Worksheets("Progress tracker").Activate
                                
                                Cells(row_num, 5) = "Y" 'Update the status for step 1 in stage 3 if successfully created
                                
                                UserForm1.TextBox_log_activity.Text = UserForm1.TextBox_log_activity.Value & "Received onboarding prep form reply for " & Cells(row_num, 10) & "," & Cells(row_num, 11) & vbNewLine
                                
                                Call update_timestamp(row_num, 5)
                        
                            End If
                    
                    
                        End If
                        

                        If Cells(row_num, 6) = "N" Then 'if supervisor/buddy inputs form reply has not be received and extracted
                    
                            'call function to search for supervisor/buddy inputs form reply and check if supervisor/buddy inputs form reply has been found and extracted successfully
                            If Module3.search_supervisor_buddy_inputs(row_num) Then 'function will return True if successful
                    
                                'Ensure that the worksheet for onboarding progress tracker in the database file is active
                                database.Worksheets("Progress tracker").Activate
                                
                                Cells(row_num, 6) = "Y" 'Update the status for step 2 in stage 3 if successfully created
                                
                                UserForm1.TextBox_log_activity.Text = UserForm1.TextBox_log_activity.Value & "Received inputs from supervisor/buddy for " & Cells(row_num, 10) & "," & Cells(row_num, 11) & vbNewLine
                                
                                Call update_timestamp(row_num, 6)
                        
                            End If
                    
                    
                        End If
                        
                        

                        
                        'if onboarding prep reply and supervisor/buddy inputs are both received but reply to Dexun/ITD email not completed
                        If check_stage_completion(row_num, 5, 6) And Cells(row_num, 7) = "N" Then


                            'call function to check and reply the most recent email for Dexun and ITD
                            If Module4.search_and_reply_email_dexun_itd(row_num) Then
                            
                                'Ensure that the worksheet for onboarding progress tracker in the database file is active
                                database.Worksheets("Progress tracker").Activate
                                
                                Cells(row_num, 7) = "Y" 'Update the status for step 2 in stage 3 if successfully created
                                
                                UserForm1.TextBox_log_activity.Text = UserForm1.TextBox_log_activity.Value & "Updated Dexun/ITD email with new onboarding info for " & Cells(row_num, 10) & "," & Cells(row_num, 11) & vbNewLine
                                
                                Call update_timestamp(row_num, 7)
                            
                            
                            End If
                        
                        
                        End If
                    

                    
                    Else
                    
                        'check if all steps in stage 4 is completed
                        If Not check_stage_completion(row_num, 8, 9) Then 'if not all steps are completed, check and run automation for remaining steps in stage 4

                    
                
                            If Cells(row_num, 8) = "N" Then 'if welcome email has not yet been created, call function to create welcome email for candidate
                                
                                
                                check_subject = Module5.send_welcome_email(row_num)
                                
                                'Ensure that the worksheet for onboarding progress tracker in the database file is active
                                database.Worksheets("Progress tracker").Activate
                                
                                If IsDraftEmailExist(check_subject) = True Then 'check if welcome email has been successfully created
                                
                                    Cells(row_num, 8) = "Y" 'Update the status for step 1 in stage 4 if successfully created
                                    
                                    UserForm1.TextBox_log_activity.Text = UserForm1.TextBox_log_activity.Value & "Created welcome email draft for " & Cells(row_num, 10) & "," & Cells(row_num, 11) & vbNewLine
                                    
                                    Call update_timestamp(row_num, 8)
                                    
                                    
                                End If
                                
                            End If
                            
                            

                        
                        
                        
                            If Cells(row_num, 9) = "N" Then 'if welcome email for supervisor/buddy has not yet been created, call function to create email for supervisor/buddy
                                
                                
                                check_subject = Module5.send_supervisor_buddy_email(row_num)
                                
                                'Ensure that the worksheet for onboarding progress tracker in the database file is active
                                database.Worksheets("Progress tracker").Activate
                                
                                If IsDraftEmailExist(check_subject) = True Then 'check if welcome email for supervisor/buddy has been successfully created
                                
                                    Cells(row_num, 9) = "Y" 'Update the status for step 2 in stage 4 if successfully created
                                    
                                    UserForm1.TextBox_log_activity.Text = UserForm1.TextBox_log_activity.Value & "Created draft welcome email for supervisor/buddy for " & Cells(row_num, 10) & "," & Cells(row_num, 11) & vbNewLine
                                    
                                    Call update_timestamp(row_num, 9)
                                    
                                    
                                End If
                                
                            End If
                    
                    

                    
    
                        End If
                    
                    End If
                
                End If
                
            End If
            
            
            
         
ContinueLoop:

        Next row_num
        
        
        
'        Unload ProgressIndicator
'        Set ProgressIndicator = Nothing
        
        
        
        'show message box to inform user that the automation process has been completed
        config = vbOKOnly + vbInformation
        MsgBox "Automation run completed.", config, "End of automation process"
        
        
    Else
    
    
        'show message box to inform user that the file does not contain any candidate's onboarding info
        config = vbOKOnly + vbExclamation
        MsgBox "No candidate info added to the database file yet.", config, "End of automation process"
    
    
    
    
    End If



End Sub


Sub update_timestamp(row_num As Integer, col_num As Integer)

    'Ensure that the timestamp data worksheet for candidate data in the database file is active
    database.Worksheets("Timestamp data").Activate
    
    Worksheets("Timestamp data").Cells(row_num, col_num) = Now 'add the timestamp data for the completion of that particular stage/step number
    
    'Ensure that the worksheet for onboarding progress tracker in the database file is active
    database.Worksheets("Progress tracker").Activate

End Sub


Sub update_progress(pct)

    With ProgressIndicator
        .FrameProgress.Caption = Format(pct, "0%")
        .LabelProgress.Width = pct * (.FrameProgress _
        .Width - 10)
    End With
    
    'The DoEvents statement is responsible for the form updating
    DoEvents
    
End Sub




Function check_inflow_date(row_num As Integer) As Boolean

    'Ensure that the main worksheet for candidate data in the database file is active
    database.Worksheets("Main data").Activate

    'Variable to check if inflow date is available
    Dim inflow_date As Variant
    inflow_date = Worksheets("Main data").Cells(row_num, 9).Value
    
    If inflow_date <> "" Then
    
        check_inflow_date = True
        
    Else
    
        check_inflow_date = False
        
    End If


End Function


Function define_specific_candidate_folder(row_num) As String

    'Ensure that the main worksheet for candidate data in the database file is active
    database.Worksheets("Main data").Activate
    
    Dim directorate As String, unique_ID As String, name As String
    directorate = Worksheets("Main data").Cells(row_num, 6)
    unique_ID = Worksheets("Main data").Cells(row_num, 1)
    name = Worksheets("Main data").Cells(row_num, 2)

    ' define the candidateFolderPath based on the unique identifier number and the name of the candidate
    define_specific_candidate_folder = candidates_folder_path + "\" + directorate + "\" + name + "-" + unique_ID
    
    
End Function

Function check_stage_completion(row_num As Integer, start_col As Integer, end_col As Integer)
    'A function to check if a particular stage in the progress tracker is completed (defined as all cells in range for that stage having the value of "Y")
    'The function will return false if the stage is not yet completed
    
    Dim cell As Range
    
    'select the progress tracker as the active sheet for the iteration and checks
    Worksheets("Progress tracker").Activate
    
    For Each cell In Range(Cells(row_num, start_col), Cells(row_num, end_col))
        If cell.Value = "N" Then 'if there are any cells with "N" in that stage, that means that stage is not yet completed
            check_stage_completion = False 'function will return False
            Exit Function
        End If
        
    Next cell
    
    'if no cells contain "N" which implies that all cells contain "Y", that means the stage is completed
    check_stage_completion = True 'function will return True
    

End Function


Function IsDraftEmailExist(subject As String) As Boolean
    Dim olApp As Object
    Dim olNs As Object
    Dim olFolder As Object
    Dim olMail As Object
    Dim found As Boolean
    
    ' Create Outlook application object
    Set olApp = CreateObject("Outlook.Application")
    Set olNs = olApp.GetNamespace("MAPI")
    
    ' Get the Drafts folder
    Set olFolder = olNs.GetDefaultFolder(16) ' 16 represents the Drafts folder
    
    ' Loop through each item in the Drafts folder
    For Each olMail In olFolder.Items
        If olMail.subject = subject Then
            found = True
            Exit For
        End If
    Next olMail
    
    ' Return True if the email was found, False otherwise
    IsDraftEmailExist = found
    
    ' Clean up
    Set olMail = Nothing
    Set olFolder = Nothing
    Set olNs = Nothing
    Set olApp = Nothing
    
    
End Function

