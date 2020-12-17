# About this code
The purpose of this code is to automatically code responses from running children live using PowerPoint stimuli. This code was written for PowerPoint on Windows in Dec 2020, but should work for PowerPoint on Mac with minor changes (filepath directory). I'm not sure if there's anything like this for Keynote.

## How this code works
This code works by creating a **scripting dictionary** called **`data`**, which contains *pairs* of keys and values:
1. **`keys`**: column/measure names (e.g. "condition", "measure1", "measure2")
2. **`values`**: this participant's response (e.g. "structural", "yes", "blue")

`Keys` (column/measure names) are taken from the slide names (i.e. the text in a textbox formatted as a Title textbox, which can be off-screen).

`Values` (responses) are taken when buttons are clicked, using your choice of a variety of macros (i.e., sub/subroutines):
- **`measure_buttonName`**: **the main workhorse**. Takes the name of the button clicked as the response. Requires you to name objects to the desired response outputs in Selection Pane.
- `measure_buttonText`: Takes the text within the button clicked as the response. More limited use cases than measure_buttonName, but works if you're too lazy to rename all your objects and your objects are textboxes containing response text anyway.
- `measure_textEntry_popOut`: Takes the text provided in a pop-out text entry box as the response. Works for transcribing open-ended responses.
- Feel free to make more macros/subs if you need to collect more kinds of responses!

Each of the above macros has no auto-advance, `_advance1`, and `_advance2` versions to auto-advance the slides:
- No auto-advance: Stay on the slide after the response (e.g., to ask subsequent/follow-up questions). Note that this opens the possibility of multiple macro clicks on a single slide, in which case the last click will be kept as the response. If you're not auto-advancing, highly recommend adding sounds when button is clicked to minimize confusion about whether the button was clicked or not.
- **`_advance1`**: **recommended**. Advances 1 slide when clicked to keep things moving and minimize confusion/the number of actions a typical researcher needs to do.
- `_advance2`: Advances 2 slides. Useful for jumping slides during 2-step contingent measures.

And there are 2 other important macros:
- `Setup`: customize this macro and its associated UserForm based on your study specifics and what setup info you want researchers to input.
- `SaveToExcel`: saves the scripting dictionary `data` to the Excel datasheet `data.xlsx` by: assigning the keys in `data` as the header row (if header row is empty), picking out the first empty row, looking for each `key` in the header row, and assigning the key's corresponding value in that empty row in the same column as the `key`. The filepath to reference `data.xlsx` will have minor differences on Windows vs Mac.


# Getting started
0. Clone/fork this repo, or download the files here as a .zip and unzip them into a folder.
1. Clear the contents of `data.xlsx`. Add your desired column names to the header row; make sure to use the *exact* same text as your slide titles, so `SaveToExcel` can appropriately assign participants' responses. If you choose not to specify a header row, `SaveToExcel` will automatically fill in a header row using the `keys` in the order they were collected.
2. Open `stimuli.pptm` (`.pptm` means it's macro-enabled) in PowerPoint. Adapt the template slides to be your stimuli slides.
3. Add the Developer menu to your PowerPoint ribbon: Home > Options > Customize Ribbon > scroll down the right-hand column and check "Developer". Go to the new Developer menu > click "Macro Security" to make sure macros are enabled (they are usually disabled by default for security reasons).
4. Add *titles* to each slide where you're collecting responses, which will be/should match the name of its column in `data.xlsx`. View > Outline View to check the titles of all your slides. If your slide lacks a title, double-click it in Outline View to add a title (ok to drag the resulting Title textbox off-screen, but a Title textbox *must* be present in Selection Pane). If you have a pre-existing header row in `data.xlsx`, be sure that your slide titles match the *exact* text in the header row, so `SaveToExcel` can appropriately assign participants' responses. Note: two slides can have the same title, but note that clicking macros on either slide will write to the same column, so only do this if you're okay overwriting responses or if participants will only responding on one of the slides (eg the 2nd step of a 2 step measure).
5. If you are using the `measure_buttonName` series of macros (as recommended), set object names in PowerPoint via Selection Pane: Home > Editing > Select > Selection Pane. Double click an object in the Selection Pane to edit its name.

![PowerPoint slide with pictures of blue berries and pink berries, and textbox reading "measure3" just off screen. Outline view is open on the left, showing slide named measure 2. Selection pane is open on the right, showing shapes named blueberries, pinkberries, Title 3, and Title 3.](/readme_images/slide.png)
_A typical slide._ It has a Title, here just off-screen (`measure2`), that gives the slide its title in Outline View. Each response button (here, two pictures) is named in Selection Pane as whatever the response text should be (here, "blueberries", "pinkberries").

6. Link your objects to whatever macros you want to run when they are clicked. Insert > Action > Mouse click > Run macro > select your macro (e.g. `measure_buttonName_advance1`). There's no easy to way to see at a glance if an object is linked to a macro, besides trying to reinsert the macro link, so make sure everything that you want linked is linked! Particularly if you are not auto-advancing slides, I recommend you check "Play sound" as well, so the button makes a sound (e.g. "Click") when clicked (to minimize confusion about whether it was clicked already or not).

![PowerPoint menu showing options for action button, including Run macro and Play sound](/readme_images/linkToMacro.png)

_Linking a macro._ Here the object is linked to the `measure_buttonName_advance1` macro. Note the "Play sound" option if you'd like to play a sound, in addition to running the macro, when the object is clicked.

7. Now open VBA: Developer menu > Visual Basic.  Click Tools at the top > References > make sure "Microsoft Scripting Runtime" is checked. (Dictionaries are not native to VBA, so this makes sure VBA can reference its home environment, Microsoft Scripting Runtime.)
8. Go over to the left-hand Project Explorer sidebar. (If you don't see it, Ctrl + R, or View > Project Explorer). Click `Module 1` to bring up the main code. Optional: customize any code as desired. If you are on a Mac, edit the filepath in the `SaveToExcel` macro to Mac filepath syntax.
9. In the left-hand Project Manager sidebar, click `UserForm` to customize the userform associated with setup. Right-click `UserForm` > View Object to [edit the fields and aesthetics of the form](https://docs.microsoft.com/en-us/office/vba/powerpoint/how-to/create-custom-dialog-boxes) (View > Toolbox to insert new fields: `TextBox` accepts any value, `ListBox` requires selecting from pre-specified values, `ComboBox` suggests pre-specified values but accepts other values too). Right-click `UserForm` > View Code to edit the code behind the form, including how the values from the form are being saved to the `data` dictionary.

![View in VBA with SaveToExcel macro selected](/readme_images/VBA.png)
_A typical view in VBA._ Here we are looking at the `SaveToExcel` macro. Note that the Project Manager sidebar is at the top left.

# Running participants
- **Make sure `data.xlsx`, the Excel data sheet, is closed** so PowerPoint can edit it.
- If you'd like to fill out the setup form without the participant seeing it (e.g. so participant/guardian are blind to condition), you have a few options:
  - Share screen to share your slideshow. Click "pause share" to freeze the screen share, click "Setup" to fill out the form, and then click "resume share" once you're done with the form. The participant will see a frozen screen while you're filling it out.
  - Launch the slideshow and fill out the "Setup" form on your own *before* you share your screen. Note that exiting slideshow prematurely will erase saved data, so instead of exiting slideshow, you should tab out of slideshow (`Alt + Tab` on Windows or `Cmd + Tab` on Mac to switch between open windows, or try swiping left or right on your mousepad). Tab over to Zoom, share screen to share your slideshow, and you should return to slideshow.
  - If you use 2 monitors and see your slideshow on both (e.g. sharing slides, not sharing presenter view), pop-ups will appear on whichever window you click the button. So if you click pop-up buttons like "Setup" on the non-shared screen, pop-up will appear on non-shared screen. 
- You can run your slides and click macros in any order or number of times. Note that clicking macros(s) on the same slide (or different slides with the same Title) multiple times will overwrite the same value in `data`, since they share the same `key` (slide title). The *last* pressed macro will store the final response.
- **`data` will be lost if you exit slideshow before running `SaveToExcel`**. You can jump between slides without exiting slideshow by right-clicking anywhere on the slide in slideshow > See all slides > select a slide. Be prepared to reference your recording as a backup in case of premature exit.
- You may run multiple participants in the same PowerPoint session, since the code will autoreset the `data` dictionary when setting up, and `data` is also cleared when you exit slideshow.


# Editing the VBA code
Open VBA in PowerPoint: Developer > Visual Basic.

Tips for editing in VBA:
- Remember to always declare your variables (e.g. `Dim`, `Public` if you want to access them throughout your code, or `Private`) before you initialize them.
- Remember to end your loops with the corresponding `End` (e.g. an `If` requires an `End If`).
- VBA does not automatically wrap your code. Add ` _ ` (space _ space) at the end of a line to continue code on the next line. Or write your code in a code editor like [Atom](https://atom.io/) (Atom: File > Settings > Packages > install `language-vba` for VBA syntax highlighting).
- Comment your code! `'` begins a comment.
- Test run test run test run. Build and test new code in small steps so you can more easily isolate problems when your code breaks.

Tips for PowerPoint when editing in VBA:
- In PowerPoint, you can find/edit names of objects in Home > Select > Selection Pane, and names of slides in View > Outline View.
- If you rename a macro, you'll need to relink the buttons that used to reference that macro to that macro. There is no easy way to tell at a glance if a button is linked to a macro or not, so be careful, double-check your links by trying to reinsert a link, and always test run.

Debugging VBA:
- Debug > Compile VBAProject is your friend. When you make changes, run this and make sure it doesn't freak as first pass for errors.
- Watch > Add Watch to keep an eye on select variables while you're running code blocks, like the R environment.
- If you're running into some bug, try printing a key variable with `print` (print to the "Immediate" window at the bottom of your VBA screen), or have it pop out in a `MsgBox` at a key point in your code to see what's wrong.
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
