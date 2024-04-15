
Option Explicit



Sub UserForm_Initialize()

    'set the width and height of the dialog box when it is being loaded
    UserForm1.Width = 480
    UserForm1.Height = 415
    
    'set the "Add New Candidate" tab to be the active tab
    UserForm1.MultiPage1.TabIndex = 1

    'Ensure that the data source worksheet in the database file is active
    database.Worksheets("Ref data").Activate
    
    'Set the options available for the directorate combo box by referencing the data source in the database file
    UserForm1.ComboBox_directorate.RowSource = Worksheets("Ref data").Range(Cells(2, 2), Cells(11, 2)).Address
    'Set the options available for the inflow mode combo box by referencing the data source in the database file
    UserForm1.ComboBox_inflow_mode.RowSource = Worksheets("Ref data").Range(Cells(2, 1), Cells(14, 2)).Address
    
    
    Call update_change_inflow_date_tab
    

End Sub



Private Sub CommandButton_run_automation_Click()
    
    Call Module1.automation_main_process
    
End Sub



Private Sub button_add_candidate_Click()

    'Function/sub-procedure to be included here for checking if all required values are entered
    'Message box with warning message will appear if requirements not satisfied and exit sub
    If check_no_blanks <> True Then
    
        MsgBox "Please ensure all required fields at this stage (except for inflow date) are entered.", vbOKOnly + vbExclamation, "Empty fields detected"
        Exit Sub
    End If

    'Ensure that the database file is active
    database.Activate

    Dim last_row As Integer
    'Find the last empty row in the database
    last_row = Worksheets("Main data").Cells(Rows.Count, 2).End(xlUp).Row + 1
    
    
    ' Add the data for the new candidate in the database file
    Worksheets("Main data").Cells(last_row, 2) = UserForm1.TextBox_name.Value
    Worksheets("Main data").Cells(last_row, 3) = UserForm1.TextBox_work_email.Value
    Worksheets("Main data").Cells(last_row, 4) = UserForm1.TextBox_personal_email.Value
    Worksheets("Main data").Cells(last_row, 5) = UserForm1.TextBox_designation.Value
    Worksheets("Main data").Cells(last_row, 6) = UserForm1.ComboBox_directorate.Value
    Worksheets("Main data").Cells(last_row, 7) = UserForm1.TextBox_unit.Value
    Worksheets("Main data").Cells(last_row, 8) = UserForm1.ComboBox_inflow_mode.Value
    Worksheets("Main data").Cells(last_row, 9) = UserForm1.TextBox_inflow_date.Value
    ' Create the initial stage completion status for the new candidate
    Call initialize_condition_values(last_row, UserForm1.ComboBox_inflow_mode.Value)
    
    UserForm1.TextBox_name.Text = ""
    UserForm1.TextBox_work_email.Text = ""
    UserForm1.TextBox_personal_email.Text = ""
    UserForm1.TextBox_designation.Text = ""
    UserForm1.TextBox_unit.Text = ""
    UserForm1.TextBox_inflow_date = ""
    
    
    Call update_change_inflow_date_tab
    
End Sub


Function check_no_blanks() As Boolean

    'define a variable to store the return boolean value
    Dim result As Boolean
    result = True 'default value is true (e.g. no empty fields)
    
    
    If IsBlankIncludingSpaces(UserForm1.TextBox_name.Value) Then
        
        result = False
        
    ElseIf IsBlankIncludingSpaces(UserForm1.TextBox_work_email.Value) Then
        
        result = False
        
    ElseIf IsBlankIncludingSpaces(UserForm1.TextBox_personal_email.Value) Then
    
        result = False
        
    ElseIf IsBlankIncludingSpaces(UserForm1.TextBox_designation.Value) Then
    
        result = False
        
    ElseIf IsBlankIncludingSpaces(UserForm1.ComboBox_directorate.Value) Then
        
        result = False
        
    ElseIf IsBlankIncludingSpaces(UserForm1.TextBox_unit.Value) Then
    
        result = False
        
    ElseIf IsBlankIncludingSpaces(UserForm1.ComboBox_inflow_mode.Value) Then
    
        result = False
    
    End If
    
    check_no_blanks = result

End Function

Function IsBlankIncludingSpaces(inputString As String) As Boolean
    If Len(Trim(inputString)) = 0 Then
        IsBlankIncludingSpaces = True
    Else
        IsBlankIncludingSpaces = False
    End If
End Function


Sub initialize_condition_values(row_num As Integer, inflow_mode As String)
    ' This subroutine checks for the type of inflow mode in order to set the correct initial condition values
    
    ' Navigate to the progress tracker worksheet
    Worksheets("Progress tracker").Select
    
    ' Define the arrays for storing the different types of inflow modes
    Dim new_hire_array As Variant
    new_hire_array = Array("Open Recruitment on MX", "MOPO", _
    "PSLP GP 1st", "PSLP GP 2nd", "PSLP GP 3rd", "PSLP GP - Eng", "PSLP SP")
    
    Dim secondment_array As Variant
    secondment_array = Array("Secondment", "Transfer", "IO Posting", "AO Posting", "LS Posting", _
                            "Open Recruitment on GovTech", "Open Recruitment on OGP")
                            
                        
    
    ' Define the variables for keeping track of each iteration in the for loop to evaluate the inflow mode
    Dim Count As Integer, col_num As Integer
                                   
    For Count = LBound(new_hire_array) To UBound(secondment_array)
    ' check for the type of inflow mode in order to assign the correct initial values
        If new_hire_array(Count) = inflow_mode Then

            Range(Cells(row_num, 1), Cells(row_num, 9)) = "N"
            'optional steps to add a unique ID using a custom function
            'Cells(row_num, 11) = unique_id()
    
        ElseIf secondment_array(Count) = inflow_mode Then
            
            Range(Cells(row_num, 1), Cells(row_num, 1)) = "NA"
            Range(Cells(row_num, 2), Cells(row_num, 2)) = "N"
            Range(Cells(row_num, 3), Cells(row_num, 3)) = "NA"
            Range(Cells(row_num, 4), Cells(row_num, 9)) = "N"
            'optional steps to add a unique ID using a custom function
            'Cells(row_num, 11) = unique_id()

        End If
    
    Next Count
    


End Sub



Private Sub CommandButton_add_inflow_date_Click()
    
    'Ensure that the main worksheet for candidate data in the database file is active
    database.Worksheets("Main data").Activate
    
     Dim last_row As Integer
    'Find the last empty row in the database
    last_row = Worksheets("Main data").Cells(Rows.Count, 2).End(xlUp).Row
    
    Dim item As Range
    'Find the matching candidate ID using a for loop to iterate through each available candidate ID
    For Each item In Worksheets("Main data").Range(Cells(4, 1), Cells(last_row, 1))
        'If a matching ID is found, add the inflow date for the candidate
        If item = UserForm1.ComboBox_no_inflow_date_ID Then
            item.Offset(0, 8) = UserForm1.TextBox_add_inflow_date.Value
        End If
    Next item
    
    Call update_change_inflow_date_tab
    
End Sub

Private Sub CommandButton_change_inflow_date_Click()
    
    'Ensure that the main worksheet for candidate data in the database file is active
    database.Worksheets("Main data").Activate
    
    Dim last_row As Integer
    'Find the last empty row in the database
    last_row = Worksheets("Main data").Cells(Rows.Count, 2).End(xlUp).Row
    
    Dim item As Range
    'Find the matching candidate ID using a for loop to iterate through each available candidate ID
    For Each item In Worksheets("Main data").Range(Cells(4, 1), Cells(last_row, 1))
        'If a matching ID is found, add the inflow date for the candidate
        If item = UserForm1.ComboBox_existing_inflow_date_ID Then
            item.Offset(0, 8) = UserForm1.TextBox_change_inflow_date.Value
        End If
    Next item
    
    Call update_change_inflow_date_tab
    
End Sub


Sub update_change_inflow_date_tab()

    Call clear_comboBox_options

    'Ensure that the main worksheet for candidate data in the database file is active
    database.Worksheets("Main data").Activate

    Dim last_row As Integer
    'Find the last empty row in the database
    last_row = Worksheets("Main data").Cells(Rows.Count, 2).End(xlUp).Row

    Dim item As Range
    'Set the options available in the combo boxes to select the candidate names available for adding new inflow date/updating existing inflow date
    For Each item In Worksheets("Main data").Range(Cells(4, 2), Cells(last_row, 2))
            'Debug.Print "Checking row for name: " & item.Row
        If item.Offset(0, 27) = "Incomplete" And item.Offset(0, 7) = "" Then
            'Debug.Print "Found incomplete and missing inflow date for row: " & item.Row
            UserForm1.ComboBox_no_inflow_date.AddItem item
        ElseIf item.Offset(0, 27) = "Incomplete" And item.Offset(0, 7) <> "" Then
            'Debug.Print "Found incomplete with existing inflow date for row: " & item.Row
            UserForm1.ComboBox_existing_inflow_date.AddItem item
        End If
    Next item
    
    'Set the options available in the combo boxes to select the candidate IDs available for adding new inflow date/updating existing inflow date
    For Each item In Worksheets("Main data").Range(Cells(4, 1), Cells(last_row, 1))
            'Debug.Print "Checking row for ID: " & item.Row
        If item.Offset(0, 28) = "Incomplete" And item.Offset(0, 8) = "" Then
            'Debug.Print item.Row
            UserForm1.ComboBox_no_inflow_date_ID.AddItem item
        ElseIf item.Offset(0, 28) = "Incomplete" And item.Offset(0, 8) <> "" Then
            'Debug.Print item.Row
            UserForm1.ComboBox_existing_inflow_date_ID.AddItem item
        End If
    Next item


End Sub


Sub clear_comboBox_options()

    UserForm1.ComboBox_no_inflow_date.Clear
    UserForm1.ComboBox_no_inflow_date_ID.Clear
    UserForm1.ComboBox_existing_inflow_date.Clear
    UserForm1.ComboBox_existing_inflow_date_ID.Clear

End Sub
