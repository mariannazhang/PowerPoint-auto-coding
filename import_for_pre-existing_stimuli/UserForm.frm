VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm 
   Caption         =   "Setup Form"
   ClientHeight    =   3330
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4640
   OleObjectBlob   =   "UserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    '=======================
    '**TODO: customize this part based on your study specifics
    ' make sure to rename fields in the UserForm object if you change the names of the fields below
    '=======================
    
    ' Researcher options (comboBox = can also enter non-specified option)
    researcher.Clear
    With researcher
        .AddItem "AA"
        .AddItem "BB"
        .AddItem "CC"
        .AddItem "DD"
    End With
    
    ' 1st researcher in list will be the default selection
    researcher.ListIndex = 0
    
    ' Condition options (listBox = only accepts specified options)
    condition.Clear
    With condition
        .AddItem "condition1"
        .AddItem "condition2"
        .AddItem "condition3"
    End With

    ' Counterbalance options (listBox = only accepts specified options)
    counterbalance.Clear
        With counterbalance
        .AddItem "left"
        .AddItem "right"
    End With

End Sub

' ===========================
'**TODO: customize this part based on your study and datasheet specifics
' ===========================

' Save setup responses to data
Public Sub submit_Click()
    ' Collect setup info
    data("file") = ""           'handled later by SaveToExcel, customize what it fills in under SaveToExcel
    data("participant") = ""    'handled later by SaveToExcel, which will automatically increment
    
    data("test_date") = Date    'get system date from computer
    data("researcher") = researcher.Value   'the researcher selected in the setup form
    data("location") = "Zoom"   'set to "Zoom"
    
    data("condition") = condition.Value     'the condition selected in the setup form
    data("counterbalance") = counterbalance.Value   'the counterbalance version selected in the setup form
    
    
    ' Change slides based on form data
    
    ' If your different conditions are different between-subjects slides...
    ' unhide corresponding condition slide
    ' If condition = "condition1" Then
    '    ActivePresentation.Slides(2).SlideShowTransition.Hidden = msoFalse
    ' ElseIf condition = "condition2" Then
    ' ElseIf condition = "condition3" Then
    ' End If
    
    Unload Me
End Sub
