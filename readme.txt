****ABOUT THIS CODE****
The purpose of this code is to automatically code responses from running children live online using PowerPoint stimuli. This code was written for PowerPoint on Windows by Marianna Zhang in Dec 2020, but should work for PowerPoint on Mac with minor changes (filepath directory). I'm not sure if there's anything like this for Keynote.

This code works by creates 2 arrays:
1. array of COLUMNS, which is written to the 1st row of an Excel spreadsheet
2. array of this participant's RESPONSES, which is written to the first empty row of the same Excel spreadsheet

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


****GETTING STARTED****
In the same directory as your stimuli:
1. Create an empty data.xlsx file.

In PowerPoint:
2. Add the Developer menu to your ribbon. Home > Options > Customize Ribbon > Scroll down and check "Developer". Under Developer menu, click "Macro Security" to make sure macros are enabled.
3. If you are using the measures_buttonName series of macros, set object names in PowerPoint via Selection Pane: Home > Editing > Select > Selection Pane. Double click an object in the Selection Pane to edit its name.
4. Link your objects to whatever macros you want to use. Insert > Link or Action > Run a macro > select your macro. There's no easy to way to see at a glance if your macro is linked, besides trying to reinsert macro link, so be careful about this!
* Note: if you rename a macro, you'll need to relink the buttons that used to reference that macro to that macro.

In VBA (PowerPoint > Developer > Visual Basic)
5. In the left-hand Project Manager sidebar, click Module 1 to bring up the main code, and customized the setup macro as desired.
6. In the left-hand Project Manager sidebar, click UserForm to customize the userform associated with setup.
7. If you are on a Mac, edit the filepath in SaveToExcel to Mac filepath syntax.

****RUNNING****
- Make sure Excel data sheet is closed so PowerPoint can edit it.
- For macros that don't auto-advance, DO NOT click button twice. Afer clicking, advance slide manually by clicking elsewhere! Otherwise you'll get a double-recorded response that may bump subsequent responses out of range of the array.
- If you use 2 monitors, pop-ups will appear on whichever window you click the button. So if you click on the non-shared screen, pop-up will appear on non-shared screen.
- You may run multiple participants in the same PowerPoint session, since the code will autoreset RESPONSES when setting up.

****EDITING THE VBA CODE****
- Remember to always declare your variables ("Dim", "Public", or "Private") before you initialize them.
- VBA does not automatically wrap your code. Add " _ " (space _ space) at the end of a line to continue code on the next line. Or write your code in a code editor like Atom (File > Settings > Packages > install language-vba for VBA syntax highlighting)
- Comment your code! ' begins a comment.

- In PowerPoint, you can find/edit names of objects in Home > Select > Selection Pane, and names of slides in View > Outline View.

- Debug > Compile VBAProject is your friend. When you make changes, run this and make sure it doesn't freak as first pass for errors.
- (Should be a better way lol but) For jank debugging, add MsgBox (X) at a key point of the code to check that X looks okay.
- If the Excel file did not update, or if a button that should have done something didn't do something, check your code for errors.
- If the Excel file is locked for editing when you open it, force quit Excel from Task Manager and check your code for errors.
- If you're stuck, google "VBA" and whatever you're stuck on. StackExchange and ExcelVBA are your friends. Most info online is about Excel, but it will transfer with occasional minor tweaks (worksheet -> slide, etc)

****IMPROVEMENTS****
Feel free to make edits and improvements to the code yourself!

Improvements I'm thinking to do:
- Move device to the end so responses can be easily pasted into appt tracking.
- Writing column name (measure name) and response to Excel after every response, rather than saving at the end. This will help avoid accidental double responses/non-responses causing misalignment between the two arrays. Currently, array values are also lost when you prematurely exit slide show (eg during resource allocation measure), so writing as you go would address that.
- Adding macros to track resource distributions. I have code that successfully counts distribution of target objects in left vs right, or top vs bottom of screen, but the rest of the code will need to be restructured to write to Excel after every response, since the arrays are lost when you exist slide show. Will also need to reset target objects after distribution response is recorded.
