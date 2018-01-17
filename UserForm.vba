'   CurrentCandidate is the candidate row we're looking at
Public CurrentCandidate As Double

'   Variable to store a lookup range, set by `ActivateSheet()`
Public LookupRange As Range

'   Make sure the sheet is active
Public Function ActivateSheet()

    '   Make sure relevant sheet is active
    Sheets("MarkSheet").Activate
    
    '   Set LookupRange
    Set LookupRange = Worksheets("MarkSheet").Range("B2:T1000")

End Function

'   Retrieve the First marker name from the sheet
Public Function GetCurrentFirstMarker() As String
    
    GetCurrentFirstMarker = Application.VLookup(Val(CurrentCandidate), LookupRange, 3, False)
    
End Function

'   Retrieve the First marker comment from the sheet
Public Function GetFirstMarkerComment() As String
    
    GetFirstMarkerComment = Application.VLookup(Val(CurrentCandidate), LookupRange, 5, False)
    
End Function


'   Retrieve the First marker score from the sheet
Public Function GetCurrentFirstMark() As String
    
    GetCurrentFirstMark = Application.VLookup(Val(CurrentCandidate), LookupRange, 4, False)
    
End Function

'   Retrieve the Second marker (moderator) name from the sheet
Public Function GetCurrentSecondMarker() As String
    
    GetCurrentSecondMarker = Application.VLookup(Val(CurrentCandidate), LookupRange, 14, False)
    
End Function


'   Retrieve the Second marker (moderator) score from the sheet
Public Function GetCurrentSecondMark() As String
    
    GetCurrentSecondMark = Application.VLookup(Val(CurrentCandidate), LookupRange, 15, False)
    
End Function


'   Retrieve the Second marker (moderator) comment from the sheet
Public Function GetSecondMarkerComment() As String
    
    GetSecondMarkerComment = Application.VLookup(Val(CurrentCandidate), LookupRange, 16, False)
    
End Function


'   Retrieve the agreed mark from the sheet
Public Function GetCurrentAgreedMark() As String
    
    GetCurrentAgreedMark = Application.VLookup(Val(CurrentCandidate), LookupRange, 12, False)
    
End Function


'   Populate the UserForm with information about any extension, if present; otherwise do nothing
Public Function DisplayExtension()
    
    '   Find value of extension
    Extension = Application.VLookup(Val(CurrentCandidate), LookupRange, 2, False)
    
    '   If empty, do nothing; otherwise population extenion fields in UserForm with information
    If Extension = "" Then
        ExtensionText.Caption = ""
        ExtensionFrame.Visible = False
        ExtensionText.Visible = False
    Else
        ExtensionText.Caption = Extension
        ExtensionFrame.Visible = True
        ExtensionText.Visible = True
    End If
    
End Function

'   Find rating scale values in the sheet and populate the form with them
Public Function DisplayRatings()
    
    ' Get the current candidate as a Range object
    Set ThisRow = Sheets("MarkSheet").Range("B:B").Find(What:=CurrentCandidate)
    
    '   Save rating scales
    RatingArgument.Value = Sheets("MarkSheet").Cells(ThisRow.Row, 8).Value
    RatingEvidence.Value = Sheets("MarkSheet").Cells(ThisRow.Row, 9).Value
    RatingOrganisation.Value = Sheets("MarkSheet").Cells(ThisRow.Row, 10).Value
    RatingWriting.Value = Sheets("MarkSheet").Cells(ThisRow.Row, 11).Value
    RatingUnderstanding.Value = Sheets("MarkSheet").Cells(ThisRow.Row, 12).Value
    
End Function


'   This is the key function that saves changes from the UserForm back to the sheet
Public Function SaveFormValuesToSheet()
    
    ' Get the current candidate as a Range object
    Set ThisRow = Sheets("MarkSheet").Range("B:B").Find(What:=CurrentCandidate)
    
    '   Save Comments
    If Not GetFirstMarkerComment() = FirstMarkerComments.Text Then
        Sheets("MarkSheet").Cells(ThisRow.Row, 6).Value = FirstMarkerComments.Text
    End If
    
    If Not GetSecondMarkerComment() = SecondMarkerComments.Text Then
        Sheets("MarkSheet").Cells(ThisRow.Row, 17).Value = SecondMarkerComments.Text
    End If
    
    '   Save Marks
    If Not GetCurrentFirstMark() = Mark1.Value Then
        Sheets("MarkSheet").Cells(ThisRow.Row, 5).Value = Mark1.Value
    End If
    
    If Not GetCurrentSecondMark() = Mark2.Value Then
        Sheets("MarkSheet").Cells(ThisRow.Row, 16).Value = Mark2.Value
    End If
    
    If Not GetCurrentAgreedMark() = MarkAgreed.Value Then
        Sheets("MarkSheet").Cells(ThisRow.Row, 13).Value = MarkAgreed.Value
    End If
    
    '   Save rating scales
    Sheets("MarkSheet").Cells(ThisRow.Row, 8).Value = RatingArgument.Value
    Sheets("MarkSheet").Cells(ThisRow.Row, 9).Value = RatingEvidence.Value
    Sheets("MarkSheet").Cells(ThisRow.Row, 10).Value = RatingOrganisation.Value
    Sheets("MarkSheet").Cells(ThisRow.Row, 11).Value = RatingWriting.Value
    Sheets("MarkSheet").Cells(ThisRow.Row, 12).Value = RatingUnderstanding.Value
    
End Function

'   Function to clear the form; called when changing candidates
Public Function EmptyForm()

    '   Markers and comments
    FirstMarkerName.Caption = ""
    FirstMarkerComments.Text = ""
    SecondMarkerName.Caption = ""
    SecondMarkerComments.Text = ""
    
    '   Extension
    ExtensionText.Caption = ""
    ExtensionFrame.Visible = False
    ExtensionText.Visible = False
    
    '   Marks
    Mark1.Value = ""
    Mark2.Value = ""
    MarkAgreed.Value = ""
    
    '   Ratings
    RatingArgument.Value = ""
    RatingEvidence.Value = ""
    RatingOrganisation.Value = ""
    RatingWriting.Value = ""
    RatingUnderstanding.Value = ""

End Function

'   Functions to navigate through candidates
'   Disable functions gray out and deactive next/previous buttons when there are no further candidates
Public Function DisableNext()
    NextButton.ForeColor = &H8000000F
    NextButton.Enabled = False
End Function
Public Function DisablePrevious()
    PreviousButton.ForeColor = &H8000000F
    PreviousButton.Enabled = False
End Function
'   Enable functions active next/previous buttons when there are further candidates
Public Function EnableNext()
    NextButton.ForeColor = &H80000012
    NextButton.Enabled = True
End Function
Public Function EnablePrevious()
    PreviousButton.ForeColor = &H80000012
    PreviousButton.Enabled = True
End Function


'   Action in response to click of "Next" button
Private Sub NextButton_Click()

    a = ActivateSheet()
    
    '   Determine the next row
    Dim NextRow As Integer
    If CurrentCandidate = 0 Then
        NextRow = 2
    Else
        s = SaveFormValuesToSheet()
        Set ThisRow = Sheets("MarkSheet").Range("B:B").Find(What:=CurrentCandidate)
        NextRow = ThisRow.Row + 1
    End If
    NextCandidate = Sheets("MarkSheet").Cells(NextRow, 2).Value
        
    '   Reset candidate value to next candidate number
    If Not NextCandidate = "" Then
        CandidateNumber.Value = NextCandidate
        d = EnableNext()
    Else
        d = DisableNext()
    End If

End Sub

'   Action in response to click of "Previous" button
Private Sub PreviousButton_Click()

    a = ActivateSheet()
    s = SaveFormValuesToSheet()

    '   Determine the previous row
    Dim PrevRow As Integer
    Set ThisRow = Sheets("MarkSheet").Range("B:B").Find(What:=CurrentCandidate)
    If ThisRow.Row > 2 Then
        PrevRow = ThisRow.Row - 1
        '   Reset candidate value to previous candidate number
        CandidateNumber.Value = Sheets("MarkSheet").Cells(PrevRow, 2).Value
    Else
        d = DisablePrevious()
    End If

End Sub


'   Action in response to change of candidate number from CandidateNumber listbox
Private Sub CandidateNumber_Change()

    a = ActivateSheet()
    
    ' Clear Form
    e = EmptyForm()
        
    '   Grab CandidateNumber from combo box
    SelectedCandidate = CandidateNumber.Value
    '   if empty, do nothing; otherwise populate fields
    If SelectedCandidate = "Candidate" Or SelectedCandidate = "" Then
        '   Disable PreviousButton
        d = DisablePrevious()
        '   Enable NextButton
        d = EnableNext()
    Else
        ' Set global candidate
        CurrentCandidate = SelectedCandidate
        '   VLOOKUP comment(s) based on CandidateNumber
        FirstMarkerComments.Text = GetFirstMarkerComment()
        SecondMarkerComments.Text = GetSecondMarkerComment()
        '   Display Marker Names and Marks
        FirstMarkerName.Caption = GetCurrentFirstMarker()
        Mark1.Value = GetCurrentFirstMark()
        SecondMarkerName.Caption = GetCurrentSecondMarker()
        Mark2.Value = GetCurrentSecondMark()
        MarkAgreed.Value = GetCurrentAgreedMark()
        '   Display extension details
        d = DisplayExtension()
        '   Display rating scales
        d = DisplayRatings()
        '   Enable buttons
        d = EnablePrevious()
        d = EnableNext()
        '   Set comment box value as comment from sheet
        FirstMarkerComments.SetFocus
    End If

End Sub

'   Set focus to FirstMarkerComment textbox on load
Private Sub FirstMarkerComment_Initialize()

    a = ActivateSheet()
    CandidateNumber.SetFocus

End Sub

'   Set the possible values of the rating scales listboxes
Private Sub UserForm_Initialize()
    RatingValues = Array("", "Unsatisfactory (40-49)", "Satisfactory (50-54)", "Average (55-59)", "Good (60-64)", "Very Good (65-69)", "Excellent (70-75)", "Outstanding (75+)")
    RatingArgument.List = RatingValues
    RatingEvidence.List = RatingValues
    RatingOrganisation.List = RatingValues
    RatingWriting.List = RatingValues
    RatingUnderstanding.List = RatingValues
End Sub
