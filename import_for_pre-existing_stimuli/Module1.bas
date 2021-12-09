Attribute VB_Name = "Module1"
' ===========================================================================
' README
' Documentation: https://github.com/mariannazhang/PowerPoint-auto-coding
' To get started: search in here for "TODO" to see anything that you need to customize.
' In the Project Window on the left, right-click UserForm, and search in there for "TODO" as well.
' ===========================================================================



' ===========================================================================
' DECLARE GLOBAL VARIABLES
' data is a Dictionary custom class (equivalent to Dictionary class in Microsoft Scripting Runtime, which isn't available on Mac)
' Thanks to: https://github.com/VBA-tools/VBA-Dictionary
' In data, keys = measure names (written to datasheet as column names), items = participant responses (written as rows).
' ===========================================================================
Public data As Dictionary
Public pathSeparator As String
Public datafile_path As String

' ===========================================================================
' HELPER FUNCTIONS
' ===========================================================================

' Force a number (participant of day) to 2 digit string
Public Function ForceTwoDigits(number As String) As String
    If Len(number) = 1 Then
        ForceTwoDigits = "0" & number
    Else
        ForceTwoDigits = number
    End If
End Function
    
' Go up one directory
Public Function GoUpDirectory(path As String)
    ' Starting from the end of the path, find first path separator, keep everything to the left of it
    GoUpDirectory = Left(path, InStrRev(path, pathSeparator) - 1)
End Function


' ===========================================================================
' SETUP
' ===========================================================================
' Setup at beginning
Public Sub Setup()
    ' Create a new instance of the Excel application
    Set oXLApp = CreateObject("Excel.Application")
    
    ' on Mac, the above line weirdly creates a workbook, so make sure to set it to an "application" object type
    If UCase(TypeName(oXLApp)) = "WORKBOOK" Then
    Set oXLApp = oXLApp.Application
    End If
    
    ' Hide Excel in background
    oXLApp.Visible = False
    
    ' Use Excel's pathSeparator function to set path separator to "/" on Mac, "\" on Windows
    pathSeparator = oXLApp.pathSeparator
    
    ' Quit Excel application
    oXLApp.Quit
    Set oXLApp = Nothing



    ' **TODO: adjust filepath based on your folder structure
    ' Locate the datasheet Excel file: go up 1 directory, then into data folder:
    datafile_path = GoUpDirectory(ActivePresentation.path) & pathSeparator & "data" & pathSeparator & "data.xlsx"




    ' Initialize data, make sure empty
    Set data = New Dictionary
    data.RemoveAll
    
    ' Run user form to collect setup info
    ' **TODO: customize this setup userform based on what you want your responses sheet to look like
    ' Right side Project Window > right-click UserForm > View Code > customize
    UserForm.Show
    
    ' Add in_progress tag, so you can save, exit, and resume in middle (inProgress_SaveToExcel)
    ' in_progress tag will be set to "no" when participant is complete (SaveToExcel_end)
    data("in_progress") = "yes"
    
    ' Save to excel, so researcher can setup and quit before session begins, and resume when session begins
    inProgress_saveToExcel
End Sub

' Resume in progress session
Public Sub inProgress_resume()
    ' Initialize data
    Set data = New Dictionary
    data.RemoveAll
End Sub


' ===========================================================================
' COLLECT RESPONSES BY OBJECT NAME
' Set object names in PowerPoint in Selection Pane: Home > Editing > Select > Selection Pane.
' Double click an object in selection pane to edit its name.
' ===========================================================================

' Store object name as response
Public Sub measure_buttonName(oSh As Shape)
    Dim measure As String
    Dim response As String
    
    ' measure = slide title
    measure = ActivePresentation.SlideShowWindow.View.Slide.Shapes.Title.TextFrame.TextRange.Text
    ' response = button name
    response = oSh.Name
    
    ' store measure as key, response as value. overwrites previous value if key already exists
    data(measure) = response
End Sub

' Store object name as response + advance 1 slide
Public Sub measure_buttonName_advance1(oSh As Shape)
    measure_buttonName oSh
    
    ' advance to next slide
    ActivePresentation.SlideShowWindow.View.Next
End Sub

' Store object name as response + advance 2 slides
Public Sub measure_buttonName_advance2(oSh As Shape)
    measure_buttonName oSh
    
    ' advance 2 slides
    ActivePresentation.SlideShowWindow.View.GotoSlide (oSh.Parent.SlideIndex + 2)
End Sub

' Store object name as response + advance 3 slides
Public Sub measure_buttonName_advance3(oSh As Shape)
    measure_buttonName oSh
    
    ' advance 2 slides
    ActivePresentation.SlideShowWindow.View.GotoSlide (oSh.Parent.SlideIndex + 3)
End Sub


' ===========================================================================
' STORE OBJECT NAME AS KEY, AND RESPONSE AS "clicked"
' If you have multiple objects on same slide, and want to track which/how many are clicked
' ===========================================================================

Public Sub measure_buttonNameAsKey(oSh As Shape)
    Dim measure As String
    Dim response As String
    
    ' measure = button name
    measure = oSh.Name
    ' response = clicked
    response = "clicked"
    
    ' store measure as key, response as value. overwrites previous value if key already exists
    data(measure) = response
End Sub

' ===========================================================================
' COLLECT TEXT IN OBJECT AS RESPONSE
' If you'd rather not rename all your objects, and your objects contain exact response text anyway.
' ===========================================================================


' Store button text as response
Public Sub measure_buttonText(oSh As Shape)
    Dim measure As String
    Dim response As String
    
    ' measure = slide title
    measure = ActivePresentation.SlideShowWindow.View.Slide.Shapes.Title.TextFrame.TextRange.Text
    MsgBox (measure)
    ' response = button text
    response = oSh.TextFrame.TextRange.Text
    MsgBox (response)
    
    ' store measure as key, response as value. overwrites previous value if key already exists
    data(measure) = response
    
End Sub

' Store button text as response + advance 1 slide
Public Sub measure_buttonText_advance1(oSh As Shape)
    measure_buttonText oSh
    
    ' advance to next slide
    ActivePresentation.SlideShowWindow.View.Next
End Sub

' Store button text as response + advance 2 slides
Public Sub measure_buttonText_advance2(oSh As Shape)
    measure_buttonText oSh
    
    ' advance 2 slides
    ActivePresentation.SlideShowWindow.View.GotoSlide (oSh.Parent.SlideIndex + 2)
End Sub

' ===========================================================================
' COLLECT TEXT ENTRY
' for free response, open-ended questions
' ===========================================================================

' Store pop out text entry as response
Public Sub measure_textEntry_popOut()
    Dim measure As String
    Dim response As String
    
    ' measure = slide title
    measure = ActivePresentation.SlideShowWindow.View.Slide.Shapes.Title.TextFrame.TextRange.Text
    
    ' response = pop-out text entry
    response = InputBox("Response")
    
    ' store measure as key, response as value. overwrites previous value if key already exists
    data(measure) = response
End Sub

' Store pop out text entry as response + advance 1 slide
Public Sub measure_textEntry_popOut_advance1()
    measure_textEntry_popOut
    
    ' advance 1 slide
    ActivePresentation.SlideShowWindow.View.Next
End Sub

' ===========================================================================
' COLLECT DISTRIBUTION OF OBJECTS ON SLIDE
' allocation measures
' ===========================================================================

' given 2 objects whose name contains "anchor"
' see whether objects whose name contains "target" are closer to one anchor or the other anchor
' returns 2 dictionary pairs:
' anchor object names as keys, number of target objects closer to that anchor object as responses

Public Sub measure_allocation()
    ' declare counters to count how many targets are closer to anchor 1 vs anchor 2, start at 0
    Dim anchor1_count As Long
    Dim anchor2_count As Long
    
    anchor1_count = 0
    anchor2_count = 0
    
    ' make dictionaries to store anchors' vertical and horizontal positions
    Dim anchors_vPosition As Dictionary
    Set anchors_vPosition = New Dictionary
    Dim anchors_hPosition As Dictionary
    Set anchors_hPosition = New Dictionary
    
    ' for each shape on the slide whose name contains "anchor"...
    Dim oSh As Shape
    For Each oSh In ActivePresentation.SlideShowWindow.View.Slide.Shapes
        If oSh.Name Like "*anchor*" Then
            ' store vertical position
            anchors_vPosition(oSh.Name) = oSh.Top + (oSh.Height / 2)
            ' store horizontal position
            anchors_hPosition(oSh.Name) = oSh.Left + (oSh.Width / 2)
        End If
    Next oSh
    
    ' for each shape on the slide whose name contains "target"...
    For Each oSh In ActivePresentation.SlideShowWindow.View.Slide.Shapes
        If oSh.Name Like "*target*" Then
            ' get target's horizontal and vertical position
            Dim target_hPosition As Long
            Dim target_vPosition As Long
            target_hPosition = oSh.Left + (oSh.Width / 2)
            target_vPosition = oSh.Top + (oSh.Height / 2)

            ' calculate target distance to anchor 1 or anchor 2 using Pythagorean theorem
            Dim distance_target1 As Long
            Dim distance_target2 As Long
            
            distance_target1 = _
            Sqr((target_hPosition - anchors_hPosition.Items(0)) ^ 2 + (target_vPosition - anchors_vPosition.Items(0)) ^ 2)
            
            distance_target2 = _
            Sqr((target_hPosition - anchors_hPosition.Items(1)) ^ 2 + (target_vPosition - anchors_vPosition.Items(1)) ^ 2)
            
            ' record if target is closer to anchor 1 or anchor 2
            If distance_target1 < distance_target2 Then
                anchor1_count = anchor1_count + 1
            ElseIf distance_target1 > distance_target2 Then
                anchor2_count = anchor2_count + 1
            End If
        End If
    Next oSh
    
    ' clean up key by removing: "_anchor_", "_anchor", "anchor_", or "anchor"
    ' vbTextCompare = case insensitive
    Dim anchor1_cleanName As String
    anchor1_cleanName = anchors_vPosition.Keys(0)
    anchor1_cleanName = Replace(anchor1_cleanName, "_anchor_", "", 1, 1, vbTextCompare)
    anchor1_cleanName = Replace(anchor1_cleanName, "_anchor", "", 1, 1, vbTextCompare)
    anchor1_cleanName = Replace(anchor1_cleanName, "anchor_", "", 1, 1, vbTextCompare)
    anchor1_cleanName = Replace(anchor1_cleanName, "anchor", "", 1, 1, vbTextCompare)
    
    Dim anchor2_cleanName As String
    anchor2_cleanName = anchors_vPosition.Keys(1)
    anchor2_cleanName = Replace(anchor2_cleanName, "_anchor_", "", 1, 1, vbTextCompare)
    anchor2_cleanName = Replace(anchor2_cleanName, "_anchor", "", 1, 1, vbTextCompare)
    anchor2_cleanName = Replace(anchor2_cleanName, "anchor_", "", 1, 1, vbTextCompare)
    anchor2_cleanName = Replace(anchor2_cleanName, "anchor", "", 1, 1, vbTextCompare)
    
    ' store 2 pairs: clean anchor name as key, target count ass items. overwrites previous value if key already exists
    data(anchor1_cleanName) = anchor1_count
    data(anchor2_cleanName) = anchor2_count
    
End Sub


' Allocation + advance 1
Public Sub measure_allocation_advance1()
    measure_allocation
    
    ' advance to next slide
    ActivePresentation.SlideShowWindow.View.Next
End Sub

' ===========================================================================
' RESET ALLOCATED SHAPES
' for use with measure_allocation
' ===========================================================================


' Reset target objects (name contains "target") at bottom of screen
' **TODO: If you want to use this macro, you must *manually* specify which slide needs to be reset.
' **TODO: Changing the name of this macro will require you reconnect any buttons previously linked to this macro.

Sub reset_horizontallyAtBottom_slide25()
    ' set target slide
    Dim targetSlide As Slide
    Set targetSlide = ActivePresentation.Slides(25) ' manually set this to the reset slide!!
    
    ' desired vertical reset position = at bottom of slide (85% down)
    Dim targets_vPosition As Long
    targets_vPosition = ActivePresentation.PageSetup.SlideHeight * 0.85
    
    ' make dictionary to store target shapes + their desired (horizontal) reset positions
    Dim targets_hPositions As Dictionary
    Set targets_hPositions = New Dictionary
    
    ' add each shape whose name contains "target"
    For Each oSh In targetSlide.Shapes
        If oSh.Name Like "*target*" Then
            targets_hPositions(oSh.Name) = "blah"
        End If
    Next oSh
    
    ' targets should be (horizontally) spread evenly across slide
    Dim interTargetDistance As Long
    interTargetDistance = (ActivePresentation.PageSetup.SlideWidth) _
                            / (targets_hPositions.Count + 1)
    
    ' go through all targets, assign each a reset position that is the calculated distance apart
    Dim targetPosition As Long
    targetPosition = interTargetDistance
    
    For Each oSh In targetSlide.Shapes
        If oSh.Name Like "*target*" Then
            targets_hPositions(oSh.Name) = targetPosition
            targetPosition = targetPosition + interTargetDistance
        End If
    Next oSh
    
    ' actually move the targets now:
    For Each oSh In targetSlide.Shapes
        If oSh.Name Like "*target*" Then
            ' horizontal = assigned reset position - object half-width so object is centered
            oSh.Left = targets_hPositions(oSh.Name) - (oSh.Width / 2)
            ' vertical = vertical reset position - object half-height so object is centered
            oSh.Top = targets_vPosition - (oSh.Height / 2)
        End If
    Next oSh
    
End Sub


' Reset target objects (name contains "target") horizontally, between 2 anchor objects (name contains "anchor")
' **TODO: If you want to use this macro, you must *manually* specify which slide needs to be reset.
' **TODO: Changing the name of this macro will require you reconnect any buttons previously linked to this macro.

Public Sub reset_horizontallyBtwnAnchors_slide26()
    ' set target slide
    Dim targetSlide As Slide
    Set targetSlide = ActivePresentation.Slides(26)
    
    ' store anchors' horizontal positions
    Dim anchors_vPosition As Dictionary
    Set anchors_vPosition = New Dictionary
    
    ' for each shape on the slide whose name contains "anchor"...
    Dim oSh As Shape
    For Each oSh In targetSlide.Shapes
        If oSh.Name Like "*anchor*" Then
            ' store vertical position
            anchors_vPosition(oSh.Name) = oSh.Top + (oSh.Height / 2)
        End If
    Next oSh
    
    ' desired vertical reset position = average point between anchors
    Dim targets_vPosition As Long
    targets_vPosition = (anchors_vPosition.Items(0) + anchors_vPosition.Items(1)) / 2
    
    ' make dictionary to store target shapes + their desired (horizontal) reset positions
    Dim targets_hPositions As Dictionary
    Set targets_hPositions = New Dictionary
    
    ' add each shape whose name contains "target"
    For Each oSh In targetSlide.Shapes
        If oSh.Name Like "*target*" Then
            targets_hPositions(oSh.Name) = "blah"
        End If
    Next oSh
    
    ' targets should be (horizontally) spread evenly across slide
    Dim interTargetDistance As Long
    interTargetDistance = (ActivePresentation.PageSetup.SlideWidth) _
                            / (targets_hPositions.Count + 1)
    
    ' go through all targets, assign each a reset position that is the calculated distance apart
    Dim targetPosition As Long
    targetPosition = interTargetDistance
    
    For Each oSh In targetSlide.Shapes
        If oSh.Name Like "*target*" Then
            targets_hPositions(oSh.Name) = targetPosition
            targetPosition = targetPosition + interTargetDistance
        End If
    Next oSh
    
    ' actually move the targets now:
    For Each oSh In targetSlide.Shapes
        If oSh.Name Like "*target*" Then
            ' horizontal = assigned reset position - object half-width so object is centered
            oSh.Left = targets_hPositions(oSh.Name) - (oSh.Width / 2)
            ' vertical = vertical reset position - object half-height so object is centered
            oSh.Top = targets_vPosition - (oSh.Height / 2)
        End If
    Next oSh
    
End Sub


' Reset target shapes (name contains "target") vertically between 2 anchors (name contains "anchor")
' **TODO: If you want to use this macro, you must *manually* specify which slide needs to be reset.
' **TODO: Changing the name of this macro will require you reconnect any buttons previously linked to this macro.

Public Sub reset_verticallyBtwnAnchors_slide27()
    ' set target slide
    Dim targetSlide As Slide
    Set targetSlide = ActivePresentation.Slides(27)
    
    ' store anchors' horizontal positions
    Dim anchors_hPosition As Dictionary
    Set anchors_hPosition = New Dictionary
    
    ' for each shape on the slide whose name contains "anchor"...
    Dim oSh As Shape
    For Each oSh In targetSlide.Shapes
        If oSh.Name Like "*anchor*" Then
            ' store horizontal position
            anchors_hPosition(oSh.Name) = oSh.Left + (oSh.Width / 2)
        End If
    Next oSh
    
    ' desired horizontal reset position = average point between anchors
    Dim targets_hPosition As Long
    targets_hPosition = (anchors_hPosition.Items(0) + anchors_hPosition.Items(1)) / 2
    
    ' make dictionary to store target shapes + their desired (vertical) reset positions
    Dim targets_vPositions As Dictionary
    Set targets_vPositions = New Dictionary
    
    ' add each shape whose name contains "target"
    For Each oSh In targetSlide.Shapes
        If oSh.Name Like "*target*" Then
            targets_vPositions(oSh.Name) = "blah"
        End If
    Next oSh
    
    ' targets should be (vertically) spread evenly across slide, minus headroom for Zoom video thumbnail
    Dim interTargetDistance As Long
    Dim videoThumbHeight As Long
    
    videoThumbHeight = 150
    interTargetDistance = (ActivePresentation.PageSetup.SlideHeight - videoThumbHeight) _
                            / (targets_vPositions.Count + 1)
    
    ' go through all targets, assign each a reset position that is the calculated distance apart
    Dim targetPosition As Long
    targetPosition = interTargetDistance + videoThumbHeight
    
    For Each oSh In targetSlide.Shapes
        If oSh.Name Like "*target*" Then
            targets_vPositions(oSh.Name) = targetPosition
            targetPosition = targetPosition + interTargetDistance
        End If
    Next oSh
    
    ' actually move the targets now:
    For Each oSh In targetSlide.Shapes
        If oSh.Name Like "*target*" Then
            ' vertical = assigned reset position - object half-height so object is centered
            oSh.Top = targets_vPositions(oSh.Name) - (oSh.Height / 2)
            ' horizontal = horizontal reset position - object half-width so object is centered
            oSh.Left = targets_hPosition - (oSh.Width / 2)
        End If
    Next oSh
End Sub


' ===========================================================================
' MACROS TO SAVE DATA TO EXCEL SPREADSHEET
' ===========================================================================

' Save to Excel function
Function saveToExcel(in_progress As Boolean)
    Dim oXLApp As Object
    Dim oWb As Object
    Dim oSheet As Object
    
    Dim inProgressColumn As Long
    Dim row As Long
    Dim thisKey As Variant
    
    ' Open the Excel application
    Set oXLApp = CreateObject("Excel.Application")
    
    ' on Mac, the above line weirdly creates a workbook, so make sure to set it to an "application" object type
    If UCase(TypeName(oXLApp)) = "WORKBOOK" Then
    Set oXLApp = oXLApp.Application
    End If
    
    ' Hide in background
    oXLApp.Visible = False
    
    
    
    ' Open the datasheet Excel file using the datafile_path
    Set oWb = oXLApp.Workbooks.Open(datafile_path)
    
    ' Get the first sheet
    Set oSheet = oWb.Worksheets(1)
    
    ' ==========================
    ' SET UP HEADER ROW if needed
    ' NOTE: if you already have a header column, columns that don't match keys will be left blank (e.g. for manual fill))
    
    ' If 1st row empty, fill 1st row with measures (keys)
    If oSheet.Range("A1") = "" Then
        oSheet.Range("A1").Resize(1, data.Count).Value = data.Keys
    End If
    
    ' ==========================
    ' FIND TARGET ROW
    ' If there is a row in progress, pick that as target row; otherwise, pick first empty row as target row
    
    ' Try to find the column called "in_progress"
    inProgressColumn = oXLApp.Match("in_progress", oSheet.Rows(1), 0)
    
    ' If no "in_progress" column...
    If IsError(inProgressColumn) Then
        ' throw warning
        MsgBox ("Warning: 'in_progress' column not found in datasheet, writing responses to new row")
        
        ' find empty row: move down column A (starting w row 2), pick the first row with an empty first cell
        row = 2
        While oSheet.Range("A" & row) <> ""
            row = row + 1
        Wend
        
    ' If there is an "in_progress" column...
    Else
        ' Check that column to see if there is a row in progress
        If IsError(oXLApp.Match("yes", oSheet.Columns(inProgressColumn), 0)) Then
            ' if no row in progress, find empty row: move down column A (starting w row 2), pick the first row with an empty first cell
            row = 2
            While oSheet.Range("A" & row) <> ""
                row = row + 1
            Wend
        Else
            ' if there's a row in progress, pick row in progress as target row
            row = oXLApp.Match("yes", oSheet.Columns(inProgressColumn), 0)
        End If
    End If
    
    
    ' ==========================
    ' FILL OUT REMAINING ENTRIES IN DATA
    
    ' Assign participant number
    data("participant") = row - 1
    
    
    ' Set in_progress based on whether session is still in progress
    If in_progress = True Then
        ' inProgress_saveToExcel does this
        data("in_progress") = "yes"
    Else
        ' saveToExcel_end does this
        data("in_progress") = "no"
    End If


    ' Assign file ==========
    ' Try to find columns "file" and "test_date"
    Dim fileColumn As Variant
    Dim testDateColumn As Variant
    Dim participantOfDay As String
    
    fileColumn = oXLApp.Match("file", oSheet.Rows(1), 0)
    testDateColumn = oXLApp.Match("test_date", oSheet.Rows(1), 0)
    
    ' If no "file" or "test_date" column...
    If IsError(fileColumn) Or IsError(testDateColumn) Then
        ' throw warning
        MsgBox ("Warning: 'file' or 'test_date' column not found in datasheet, writing 'file' as 1st participant of day")
        
        ' record as 1st participant of day
        participantOfDay = 1
        
    ' If there is a "test_date" column...
    Else
        ' Check that column to see if test_date of previous participant is today
        Dim prev_test_date As Variant
        prev_test_date = oSheet.Cells(row - 1, testDateColumn)
        
        ' If no previous participant, or previous participant is from diff day...
        If IsError(prev_test_date) Or prev_test_date <> Date Then
            ' record as 1st participant of day
            participantOfDay = 1
            
        ' If previous participant is from today...
        Else
            ' get previous participant file
            Dim prev_file As String
            prev_file = oSheet.Cells(row - 1, fileColumn)
            
            ' increment particpant of day numbering from previous participant
            participantOfDay = CInt(Right(prev_file, 2)) + 1
        End If
    End If
    
    ' save file string
    ' TODO: customize this file string (currently e.g.: storybook_20210109_01)
    data("file") = "studyname_" & Format(Date, "YYYYMMDD") & "_" & ForceTwoDigits(participantOfDay)
    
    
    ' ==========================
    ' WRITE DATA TO TARGET ROW
    
    ' For each key in dictionary, write to corresponding item to the target row:
    For Each thisKey In data.Keys()
        ' a variant can store an integer OR an error
        Dim thisColumn As Variant
        ' look for KEY in header row (0=exact match), return position in 1st row (=column number)
        thisColumn = oXLApp.Match(thisKey, oSheet.Rows(1), 0)
        
        ' if KEY is in header row, write key's corresponding ITEM to the target cell (=selected row, same column as KEY)
        If Not IsError(thisColumn) Then
            oSheet.Cells(row, thisColumn).Value = data(thisKey)
        ' if KEY is not in header row, ITEM will not be stored! so make sure KEYs and column labels match
        Else
            MsgBox ("Warning: '" & thisKey & "' column not found in datasheet")
        End If
    Next
    
    
    ' ==========================
    ' WRAP UP
    
    ' Save and close Excel file
    oWb.Save
    oWb.Close
    Set oWb = Nothing
    
    ' Quit Excel application
    oXLApp.Quit
    Set oXLApp = Nothing
    
    ' If you want a notifcation that save is complete:
    ' MsgBox("Save complete.")

End Function

' Save in middle of session, session still in progress
Public Sub inProgress_saveToExcel()
    saveToExcel in_progress:=True
End Sub

' Save at end of session, session no longer in progress
Public Sub SaveToExcel_end()
    saveToExcel in_progress:=False
End Sub



