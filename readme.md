# About this code
The purpose of this code is to automatically code responses from running children live online using PowerPoint stimuli. This code was written for PowerPoint on Windows by Marianna Zhang in Dec 2020, but should work for PowerPoint on Mac with minor changes (filepath directory). I'm not sure if there's anything like this for Keynote.

## How this code works
This code works by creating 2 arrays:
1. 1-dimensional array of COLUMNS, which is written to the 1st row of an Excel spreadsheet
2. 1-dimensional array of this participant's RESPONSES, which is written to the first empty row of the same Excel spreadsheet

This code contains a variety of macros (sub/subroutines) to collect responses:
- measure_buttonName: the main workhorse, requires you to rename objects to the desired responses.
- measure_buttonText: if you're too lazy to rename all your objects and your objects are textboxes containing response text anyway.
- measure_textEntry_popOut: pops out a box to type in a response (eg transcribe open-ended response).
- Feel free to make more macros/subs if you need to collect more kinds of responses!

Each of the above measures also has _advance1 and _advance2 versions to auto-advance the slides.
- Using the _advance1 version is HIGHLY recommended to avoid accidentally double-entering a response.
- Reasons to not auto-advance: if you want to stay on the slide after the response (eg asking subsequent/follow-up questions)
- Reasons to use _advance2: jumping slides during 2-step contingent measures

And there are 2 setup/save macros:
- Setup: customize this macro and its associated UserForm based on your study specifics and how you want your datasheet formatted.
- SaveToExcel: saves both arrays to the Excel datasheet. filepath will have minor differences on Windows vs Mac


# Getting started
1. In the same directory as your stimuli, create an empty data.xlsx file.
2. Open PowerPoint. Add the Developer menu to your ribbon. Home > Options > Customize Ribbon > Scroll down and check "Developer". Under Developer menu, click "Macro Security" to make sure macros are enabled.
3. If you are using the measures_buttonName series of macros, set object names in PowerPoint via Selection Pane: Home > Editing > Select > Selection Pane. Double click an object in the Selection Pane to edit its name.
4. Link your objects to whatever macros you want to use. Insert > Link or Action > Run a macro > select your macro. There's no easy to way to see at a glance if your macro is linked, besides trying to reinsert macro link, so be careful about this!
5. Now open VBA (Developer menu > Visual Basic). In the left-hand Project Manager sidebar, click "Module 1" to bring up the main code, and customize the setup macro as desired.
6. In the left-hand Project Manager sidebar, click "UserForm" to customize the userform associated with setup.
7. If you are on a Mac, edit the filepath in the SaveToExcel macro to Mac filepath syntax.

# Running participants
- Make sure Excel data sheet is closed so PowerPoint can edit it.
- For macros that don't auto-advance, DO NOT click button twice. Afer clicking, advance slide manually by clicking elsewhere! Otherwise you'll get a double-recorded response that may bump subsequent responses out of range of the array.
- If you use 2 monitors, pop-ups will appear on whichever window you click the button. So if you click on the non-shared screen, pop-up will appear on non-shared screen.
- You may run multiple participants in the same PowerPoint session, since the code will autoreset RESPONSES when setting up.

# Editing the VBA code
- Open VBA: PowerPoint > Developer > Visual Basic

- Remember to always declare your variables ("Dim", "Public", or "Private") before you initialize them.
- VBA does not automatically wrap your code. Add " _ " (space _ space) at the end of a line to continue code on the next line. Or write your code in a code editor like Atom (File > Settings > Packages > install language-vba for VBA syntax highlighting)
- Comment your code! ' begins a comment.

- Note: if you rename a macro, you'll need to relink the buttons that used to reference that macro to that macro.
- In PowerPoint, you can find/edit names of objects in Home > Select > Selection Pane, and names of slides in View > Outline View.

- Debug > Compile VBAProject is your friend. When you make changes, run this and make sure it doesn't freak as first pass for errors.
- (Should be a better way lol but) For jank debugging, add MsgBox (X) at a key point of the code to check that X looks okay.
- If the Excel file did not update, or if a button that should have done something didn't do something, check your code for errors.
- If the Excel file is locked for editing when you open it, force quit Excel from Task Manager and check your code for errors.
- If you're stuck, google "VBA" and whatever you're stuck on. [Microsoft Office VBA documentation](https://docs.microsoft.com/en-us/office/vba/api/overview/powerpoint), StackExchange, and ExcelVBA help forums are your friends. Most info online is about Excel, but it will transfer with occasional minor tweaks (worksheet -> slide, etc)

# Future improvements
Feel free to make edits and improvements to the code yourself!

Improvements I'm thinking to do:
- Move device to the end so responses can be easily pasted into appt tracking.
- Closer correspondence between column names and responses, to avoid accidental double responses/non-responses causing misalignment between the two arrays. Can pull the name of each measure from the name of the slide the macro is on. Not sure whether the best way to address this is with a dynamic 2D array or writing to Excel as you go. Writing to Excel as you go would add significant runtime, but might be necessary to implement coding for a resource allocation measure.
- Implement a resource allocation measure. Resource allocation involves exiting slide show to allow participants to move target objects around. I have code that successfully counts allocation of target objects in left vs right, or top vs bottom of screen, but unfortunately, prematurely exiting slide show before responses are written to Excel apparently causes all stored values to be lost. So resource allocation may require that we write to Excel as we go, or that we save to Excel right before resource allocation. Target objects will also need to be reset after the resource allocation response is recorded.
