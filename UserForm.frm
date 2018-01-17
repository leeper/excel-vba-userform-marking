VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FirstMarkerComments 
   ClientHeight    =   10440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14370
   OleObjectBlob   =   "UserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FirstMarkerComments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   Set column numbering variables as needed within LookupRange
Private Const ColCandidateNumber = 1
Private Const ColCourseCode = 2
Private Const ColExtension = 3
Private Const ColFirstMarkerName = 4
Private Const ColFirstMarkerScore = 5
Private Const ColFirstMarkerComment = 6
Private Const ColAdditionalComments = 7
Private Const ColRatingArgument = 8
Private Const ColRatingEvidence = 9
Private Const ColRatingOrganisation = 10
Private Const ColRatingWriting = 11
Private Const ColRatingUnderstanding = 12
Private Const ColAgreedMark = 13
Private Const ColProvisionalClass = 14
Private Const ColSecondMarkerName = 15
Private Const ColSecondMarkerScore = 16
Private Const ColSecondMarkerComment = 17
'   The following are not used anywhere in the UserForm code
Private Const ColExternalName = 18
Private Const ColExternalScore = 19
Private Const ColExternalComment = 20
'   CurrentCandidate is the candidate row we're looking at
Public CurrentCandidate As Double

'   Maximum Row
Private Const LastRow = 500

'   Variables to store a lookup range, set by `ActivateSheet()`
'   Candidate Number Column
Public CandidateNumberRange As Range

'   Full lookup range (candidate number should be first column); used in VLookup calls
Public LookupRange As Range
'   Make sure the sheet is active
Public Function ActivateSheet()

    '   Make sure relevant sheet is active
    Sheets("MarkSheet").Activate
    
    '   Set CandidateNumberRange (Candidate Number in First Column "A")
    Set CandidateNumberRange = Worksheets("MarkSheet").Range("A1:A1000")
    
    '   Set LookupRange (Candidate Number in First Column "A"); max is 1000 but could be anything
    Set LookupRange = Worksheets("MarkSheet").Range("A2:T1000")
    
    '   Populate ListBox of candidate numbers
    CandidateNumber.RowSource = "A2:A" & LastRow
    
    ActivateSheet = 1
    
End Function

'   Retrieve the First marker name from the sheet
Public Function GetCurrentFirstMarker() As String
    
    GetCurrentFirstMarker = Application.VLookup(Val(CurrentCandidate), LookupRange, ColFirstMarkerName, False)
    
End Function

'   Retrieve the First marker comment from the sheet
Public Function GetFirstMarkerComment() As String
    
    GetFirstMarkerComment = Application.VLookup(Val(CurrentCandidate), LookupRange, ColFirstMarkerComment, False)
    
End Function


'   Retrieve the First marker score from the sheet
Public Function GetCurrentFirstMarkerScore() As String
    
    GetCurrentFirstMark = Application.VLookup(Val(CurrentCandidate), LookupRange, ColFirstMarkerScore, False)
    
End Function

'   Retrieve the Second marker (moderator) name from the sheet
Public Function GetCurrentSecondMarker() As String
    
    GetCurrentSecondMarker = Application.VLookup(Val(CurrentCandidate), LookupRange, ColSecondMarkerName, False)
    
End Function


'   Retrieve the Second marker (moderator) score from the sheet
Public Function GetCurrentSecondMarkerScore() As String
    
    GetCurrentSecondMark = Application.VLookup(Val(CurrentCandidate), LookupRange, ColSecondMarkerScore, False)
    
End Function


'   Retrieve the Second marker (moderator) comment from the sheet
Public Function GetSecondMarkerComment() As String
    
    GetSecondMarkerComment = Application.VLookup(Val(CurrentCandidate), LookupRange, ColSecondMarkerComment, False)
    
End Function


'   Retrieve the agreed mark from the sheet
Public Function GetCurrentAgreedMark() As String
    
    GetCurrentAgreedMark = Application.VLookup(Val(CurrentCandidate), LookupRange, ColAgreedMark, False)
    
End Function


'   Populate the UserForm with information about any extension, if present; otherwise do nothing
Public Function DisplayExtension()
    
    '   Find value of extension
    Extension = Application.VLookup(Val(CurrentCandidate), LookupRange, ColExtension, False)
    
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
    Set ThisRow = CandidateNumberRange.Find(What:=CurrentCandidate)
    
    '   Save rating scales
    RatingArgument.Value = Sheets("MarkSheet").Cells(ThisRow.Row, ColRatingArgument).Value
    RatingEvidence.Value = Sheets("MarkSheet").Cells(ThisRow.Row, ColRatingEvidence).Value
    RatingOrganisation.Value = Sheets("MarkSheet").Cells(ThisRow.Row, ColRatingOrganisation).Value
    RatingWriting.Value = Sheets("MarkSheet").Cells(ThisRow.Row, ColRatingWriting).Value
    RatingUnderstanding.Value = Sheets("MarkSheet").Cells(ThisRow.Row, ColRatingUnderstanding).Value
    
End Function

'   This is the key function that saves changes from the UserForm back to the sheet
Public Function SaveFormValuesToSheet()
    
    ' Get the current candidate as a Range object
    Set ThisRow = CandidateNumberRange.Find(What:=CurrentCandidate)
    
    '   Save Comments
    If Not GetFirstMarkerComment() = FirstMarkerComments.Text Then
        Sheets("MarkSheet").Cells(ThisRow.Row, ColFirstMarkerComment).Value = FirstMarkerComments.Text
    End If
    
    If Not GetSecondMarkerComment() = SecondMarkerComments.Text Then
        Sheets("MarkSheet").Cells(ThisRow.Row, ColSecondMarkerComment).Value = SecondMarkerComments.Text
    End If
    
    '   Save Marks
    If Not GetCurrentFirstMarkerScore() = Mark1.Value Then
        Sheets("MarkSheet").Cells(ThisRow.Row, ColFirstMarkerScore).Value = Mark1.Value
    End If
    
    If Not GetCurrentSecondMarkerScore() = Mark2.Value Then
        Sheets("MarkSheet").Cells(ThisRow.Row, ColSecondMarkerScore).Value = Mark2.Value
    End If
    
    If Not GetCurrentAgreedMark() = MarkAgreed.Value Then
        Sheets("MarkSheet").Cells(ThisRow.Row, ColAgreedMark).Value = MarkAgreed.Value
    End If
    
    '   Save rating scales
    Sheets("MarkSheet").Cells(ThisRow.Row, ColRatingArgument).Value = RatingArgument.Value
    Sheets("MarkSheet").Cells(ThisRow.Row, ColRatingEvidence).Value = RatingEvidence.Value
    Sheets("MarkSheet").Cells(ThisRow.Row, ColRatingOrganisation).Value = RatingOrganisation.Value
    Sheets("MarkSheet").Cells(ThisRow.Row, ColRatingWriting).Value = RatingWriting.Value
    Sheets("MarkSheet").Cells(ThisRow.Row, ColRatingUnderstanding).Value = RatingUnderstanding.Value
    
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
'   Utility function to count words in first marker comments text box
Function CountWords(ByVal sentence As String) As Integer
    CountWords = UBound(Split(sentence, " ")) + 1
End Function

'   Calculate median word count of first marker comments
Public Function CalculateMeanWordCount() As Integer
    wordsum = 0
    nonemptycount = 0
    '   Loop over comment rows, retrievign word count of each comment
    For i = 2 To LastRow
        ThisComment = Sheets("MarkSheet").Cells(i, ColFirstMarkerComment).Value
        '   If comment is empty, ignore so it is not factored into mean
        If Not ThisComment = "" Then
            wordsum = wordsum + CountWords(ThisComment)
            nonemptycount = nonemptycount + 1
        End If
    Next i
    '   Return mean word count as integer
    If nonemptycount = 0 Then
        CalculateMeanWordCount = 0
    Else
        CalculateMeanWordCount = Int(wordsum / nonemptycount)
    End If
End Function
'   Function to word count first marker comments
Private Sub FirstMarkerComments_Change()
    '   Get word count and display
    FirstMarkerWordCount = CountWords(FirstMarkerComments.Text)
    
    '   Get mean word count and display
    FirstMarkerMeanWordCount = CalculateMeanWordCount()
    
End Sub

'   Action in response to click of "Next" button
Private Sub NextButton_Click()

    a = ActivateSheet()
    
    '   Determine the next row
    Dim NextRow As Integer
    If CurrentCandidate = 0 Then
        NextRow = 2
    Else
        s = SaveFormValuesToSheet()
        Set ThisRow = CandidateNumberRange.Find(What:=CurrentCandidate)
        NextRow = ThisRow.Row + 1
    End If
    NextCandidate = Sheets("MarkSheet").Cells(NextRow, ColCandidateNumber).Value
        
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
    Set ThisRow = CandidateNumberRange.Find(What:=CurrentCandidate)
    If ThisRow.Row > 2 Then
        PrevRow = ThisRow.Row - 1
        '   Reset candidate value to previous candidate number
        CandidateNumber.Value = Sheets("MarkSheet").Cells(PrevRow, ColCandidateNumber).Value
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
        Mark1.Value = GetCurrentFirstMarkerScore()
        SecondMarkerName.Caption = GetCurrentSecondMarker()
        Mark2.Value = GetCurrentSecondMarkerScore()
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

