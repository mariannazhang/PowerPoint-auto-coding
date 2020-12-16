# About this code
The purpose of this code is to automatically code responses from running children live online using PowerPoint stimuli. This code was written for PowerPoint on Windows in Dec 2020, but should work for PowerPoint on Mac with minor changes (filepath directory). I'm not sure if there's anything like this for Keynote.

## How this code works
This code works by creating a **scripting dictionary** called **`data`**, which contains pairs of keys and values in the order they are stored:
1. **`keys`**: column/measure names
2. **`values`**: this participant's response

`Keys` (column/measure names) are taken from the slide names (i.e., the text in a textbox formatted as a Title textbox, which can be off-screen).

`Values` (responses) are taken when buttons are clicked, using your choice of a variety of macros (i.e., sub/subroutines):
- **`measure_buttonName`**: *the main workhorse*. Takes the name of the button clicked as the response. Requires you to name objects to the desired response outputs in Selection Pane.
- `measure_buttonText`: Takes the text within the button clicked as the response. More limited use cases than measure_buttonName, but works if you're too lazy to rename all your objects and your objects are textboxes containing response text anyway.
- `measure_textEntry_popOut`: Takes the text provided in a pop-out text entry box as the response. Works for transcribing open-ended responses.
- Feel free to make more macros/subs if you need to collect more kinds of responses!

Each of the above macros has no auto-advance, _advance1, and _advance2 versions to auto-advance the slides:
- No auto-advance: If you want to stay on the slide after the response (eg asking subsequent/follow-up questions). Note that this opens the possibility of multiple macro clicks on a single slide, in which case the last click will be kept as the response.
- **`_advance1`**: *recommended*, to keep things moving and minimize confusion/the number of actions a typical researcher needs to do.
- `_advance2`: jumping slides during 2-step contingent measures

And there are 3 other important macros:
- `Setup`: customize this macro and its associated UserForm based on your study specifics and how you want your datasheet formatted.
- `putDeviceHere_advance1`: reorders the dictionary by moving the responses to the "device" measure to this point in the dictionary.
- `SaveToExcel`: saves both arrays to the Excel datasheet. filepath will have minor differences on Windows vs Mac


# Getting started
0. Download the files here to the same folder.
1. Incorporate your stimuli into stimuli.pptm (.pptm means it's macro-enabled). Clear the contents of `data.xlsx`.
2. Open PowerPoint. Add the Developer menu to your ribbon: Home > Options > Customize Ribbon > Scroll down and check "Developer". Under Developer menu, click "Macro Security" to make sure macros are enabled.
3. If you are using the `measures_buttonName` series of macros, set object names in PowerPoint via Selection Pane: Home > Editing > Select > Selection Pane. Double click an object in the Selection Pane to edit its name.
4. Link your objects to whatever macros you want to use. Insert > Action > Run a macro > select your macro. There's no easy to way to see at a glance if your macro is linked, besides trying to reinsert macro link, so be careful about this!
5. Add/check titles on all slides where you're collecting responses. View > Outline View to check the titles of all your slides. If your slide lacks a title, double click it in Outline View to add a title (ok drag the resulting textbox somewhere off-screen, but the textbox must be present).
6. Now open VBA (Developer menu > Visual Basic).  Click Tools at the top > References > make sure Microsoft Scripting Runtime is checked. (Dictionaries are not native to VBA, so this makes sure VBA can reference its home environment, Microsoft Scripting Runtime.)
7. Go over to the left-hand Project Explorer sidebar. (If you don't see it, Ctrl + R, or View > Project Explorer). Click "Module 1" to bring up the main code. Optional: customize any code as desired.
8. In the left-hand Project Manager sidebar, click `UserForm` to customize the userform associated with setup. Note that you can edit the appearance of the form if you right click and View Object, and edit the code behind the form if you right-click and View Code.
9. If you are on a Mac, edit the filepath in the `SaveToExcel` macro to Mac filepath syntax.

# Running participants
- Make sure `data`, the Excel data sheet, is closed so PowerPoint can edit it.
- When using macros that don't auto-advance slides, note that clicking multiple macros on a single slide will record the last click as the response.
- If you use 2 monitors, pop-ups will appear on whichever window you click the button. So if you click on the non-shared screen, pop-up will appear on non-shared screen.
- You may run multiple participants in the same PowerPoint session, since the code will autoreset RESPONSES when setting up.

# Editing the VBA code
Open VBA in PowerPoint by going to Developer > Visual Basic.

Tips for editing:
- Remember to always declare your variables ("Dim", "Public", or "Private") before you initialize them.
- VBA does not automatically wrap your code. Add " _ " (space _ space) at the end of a line to continue code on the next line. Or write your code in a code editor like Atom (File > Settings > Packages > install language-vba for VBA syntax highlighting)
- Comment your code! ' begins a comment.
- Note: if you rename a macro, you'll need to relink the buttons that used to reference that macro to that macro.
- In PowerPoint, you can find/edit names of objects in Home > Select > Selection Pane, and names of slides in View > Outline View.

Debugging:
- Debug > Compile VBAProject is your friend. When you make changes, run this and make sure it doesn't freak as first pass for errors.
- Watch > Add Watch to keep an eye on select variables while you're running code blocks, like the R environment.
- If you're running into some bug, try printing a key variable with `print` (it'll print to the Immediate window) or `MsgBox` (pop out window) at a key point in your code to see what's wrong.
- If the Excel file did not update, or if a button that should have done something didn't do something, check your code for errors.
- If the Excel file is locked for editing when you open it, force quit Excel from Task Manager and check your code for errors.

If you're stuck, here are some resources. Most info online is about Excel, but it will usually transfer with minor tweaks (worksheet -> slide, etc)
- Google "VBA" and whatever you're stuck on.
- [Microsoft Office VBA documentation](https://docs.microsoft.com/en-us/office/vba/api/overview/powerpoint)
- On scripting dictionary specifically: [Microsoft documentation on the Dictionary object](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dictionary-object), [Excel VBA Dictionary](https://excelmacromastery.com/vba-dictionary), [dictionary vs collection vs array](https://stackoverflow.com/questions/32479842/comparison-of-dictionary-collections-and-arrays)
- StackExchange
- various ExcelVBA help forums

# Future improvements
Here are some improvements I'm thinking to do. Definitely feel free to make edits and improvements to the code yourself too!
- Implement a resource allocation measure. Resource allocation involves exiting slide show to allow participants to move target objects around. I have code that successfully counts allocation of target objects in left vs right, or top vs bottom of screen, but unfortunately, prematurely exiting slide show before responses are written to Excel apparently causes all stored values to be lost. So resource allocation may require that we write to Excel as we go, or that we save to Excel right before resource allocation. Target objects will also need to be reset after the resource allocation response is recorded.
- Automatically generate `participantOfDay` by referencing previous files in `data` to fully automatically generate the value of `file`.
