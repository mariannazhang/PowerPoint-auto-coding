# About this code
The purpose of this code is to automatically code responses from running children live using PowerPoint stimuli. This code was written in Visual Basic ("VBA") for PowerPoint on Windows in Dec 2020, but should work for PowerPoint on Mac with minor changes (filepath directory). I'm not sure if there's anything like this for Keynote.

The template slides (`stimuli.pptm`) are adapted from the [online testing slides from Stanford Social Learning Lab](http://github.com/sociallearninglab/online_testing_materials), and are designed for use by researchers who run live testing sessions by sharing screen to share a PowerPoint slideshow on Zoom.

## How this code works
This code works by creating a [**scripting dictionary**](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dictionary-object) called **`data`**, which contains *pairs* of keys and items:
1. **`keys`**: column/measure names (e.g. "condition", "measure1", "measure2")
2. **`items`**: this participant's response (e.g. "structural", "yes", "blue")

`Keys` (column/measure names) are generally taken from the *slide titles* (as seen in View > Outline View).

`Items` (responses) are recorded when buttons are clicked, using your choice of a variety of macros (i.e., sub/subroutines) that are activated on click:
- **`measure_buttonName`**: **the main workhorse**. Takes the name of the button clicked (as seen in Home > Select > Selection Pane) as the response.
- `measure_buttonText`: Takes the text within the button clicked as the response. More limited use cases than measure_buttonName, but works if you're too lazy to rename all your objects and your objects are textboxes containing response text anyway.
- `measure_buttonNameAsKey`: Unlike the other macros, this macro takes the name of the button clicked as the `key`. It records "pressed" as the response. Used in situations where you have a variety of objects on a slide and you want to record which/how many are clicked.
- `measure_textEntry_popOut`: Takes the text provided in a pop-out text entry box as the response. Works for transcribing open-ended responses.
- `measure_allocation`: Calculates the number of target objects (objects whose name includes `target`) that are closer to each of 2 anchor objects (objects or groups of objects whose name includes `anchor`). Return *2 key-item pairs*: key is name of anchor object, item is number of target objects closer to that anchor vs the other anchor. Designed for generalized resource allocation measures where you exit slideshow, move objects around, and then resume slideshow.
- Feel free to make more macros/subs if you need to collect more kinds of responses!

Here is a summary table of all of the currently available macros to record responses:

| Macro                         | `Key` (column name)        | `Item` (response)   |
| ----------------------------- |----------------------------| --------------------|
| **`measure_buttonName`**      | **slide title**            | **button name**     |
| `measure_buttonText`          | slide title                | button text         |
| `measure_buttonNameAsKey`     | button name                | "pressed"           |
| `measure_textEntry_popOut`    | slide title                | pop-out text entry  |
| `measure_allocation`          | name of 1st anchor object  | number of targets closer to this anchor than the other anchor |
|       (^continued)            | name of 2nd anchor object  | number of targets closer to this anchor than the other anchor |

*Slide titles* can be seen/edited in View > Outline View, or by directly editing the title textbox.
*Object/button names* can be edited in Home > Select > Selection Pane.

Most of the above macros have different versions depending on how you want to advance slides:
- No auto-advance: Stay on the slide after the response (e.g., to ask subsequent/follow-up questions). Note that this opens the possibility of multiple macro clicks to the same key, in which case the last click will be kept as the response. If you're not auto-advancing, highly recommend adding sounds when button is clicked to minimize confusion about whether the button was clicked or not.
- **`_advance1`**: **recommended**. Advances 1 slide when clicked to keep things moving and minimize confusion/the number of actions a typical researcher needs to do.
- `_advance2`: Advances 2 slides. Useful for jumping slides during 2-step contingent measures.

2 important macros everyone will need to use:
- `Setup`: Initializes `data`, sets the session as "in progress", collects some setup info via a UserForm. Customize this macro's associated UserForm based on your study specifics and what setup info you want researchers to input. `Setup` also writes setup info to Excel by calling `inProgress_SaveToExcel`, so you can setup and quit slideshow if you'd like before your participant arrives.
- `SaveToExcel_end`: Saves the scripting dictionary `data` to the Excel datasheet `data.xlsx` by:
  - If header row is empty, assigns the `key`s in the order they were collected.
  - Pick out a target row. Look for a row "in progress" to continue writing to; otherwise, scan down the first column and pick out the first empty row.
  - Sets participant number based on target row, sets "in progress" to "no".
  - Look for each `key` in the header row. Assign the key's corresponding `item` to the cell in the target row in the same column as the `key`.
  - Note: if you are on a Mac, change the filepath used to reference `data.xlsx`.

A few helper macros if you anticipate exiting and resuming slideshow in the middle of your session (designed for resource allocation using `measure_allocation`):
- `inProgress_SaveToExcel`: Exactly the same as `SaveToExcel_end`, except it keeps "in progress" as "yes", so subsequent macros that write to Excel know to continue writing on the same row. Should be used immediately before exiting an in-progress slideshow. Can be used multiple times throughout a study.
- `inProgress_resume`: Exactly the same as `Setup`, except it doesn't run the setup form. Basically it just re-initializes the `data` dictionary. Should be used immediately after resuming an in-progress slideshow. Can be used multiple times throughout a study.

A few helper macros to reset target objects on slides (designed for resource allocation using `measure_allocation`):
- `reset_verticallyBtwnAnchors`: Resets target objects (objects whose name contains `target`) on a *manually specified* slide to be evenly distributed vertically (minus some headroom for your video thumbnail) between anchor objects (objects or groups of objects whose name contains `anchor`)
  - Note: change the midline calculations for these reset if your Zoom video thumbnail will not be top center.
- `reset_horizontallyBtwnAnchors`: Resets target objects (objects whose name contains `target`) on a *manually specified* slide to be evenly distributed horizontally between anchor objects.


# Getting started
0. Clone/fork this repo, or download the files here as a .zip and unzip them into a folder.
1. Clear the contents of `data.xlsx`. Add your desired column names to the header row.
  - Make sure to include a column named "in_progress".
  - Use the *exact* same text as your keys (generally slide titles), so `SaveToExcel` can appropriately assign participants' responses.
  - You can add more columns than macros assign in your slides (e.g. "parental_interference", "exp_error", "comments"). Such columns will be left blank (e.g. for manual entry after running).
  - If you choose not to specify a header row, `SaveToExcel` will automatically fill in a header row using the `keys` *in the order they were collected*.
2. Open `stimuli.pptm` (`.pptm` means it's macro-enabled) in PowerPoint. Adapt the template slides to be your stimuli slides.
  - If you are using `measure_allocation` and its `reset` helper functions:
    - Objects that will be moved/counted should have names containing "target" (e.g. "target1"). The 2 objects (or groups of objects) that will be reference points for grouping should have names containing "anchor" (e.g. "measure9_lowPerf_anchor"). Note that if you name a group "target" or "anchor", do not also name each sub-object "target" or "anchor".
    - Your resource allocation slide should contain at least 3 buttons: `inProgress_SaveToExcel` (save before exiting slideshow), `inProgress_resume` (resume after restarting slideshow), and `measure_allocation` (count the allocation). You should also include the corresponding `reset` somewhere at the end so you can reset the target objects for the next participant, and can also include `reset` on the same allocation slide if participant wants to re-do the allocation. I recommend hiding all these buttons under wherever your video thumbnail will be to reduce visual clutter for the participant.
3. Add the Developer menu to your PowerPoint ribbon: Home > Options > Customize Ribbon > scroll down the right-hand column and check "Developer". Go to the new Developer menu > click "Macro Security" to make sure macros are enabled (they are usually disabled by default for security reasons).
4. Add *titles* to each slide where you're collecting responses, which will generally serve as the name of its column in `data.xlsx`. View > Outline View to check the titles of all your slides. If your slide lacks a title, double-click it in Outline View to add a title (ok to drag the resulting Title textbox off-screen, but a Title textbox *must* be present in Selection Pane). If you have a pre-existing header row in `data.xlsx`, be sure that your slide titles match the *exact* text in the header row, so `SaveToExcel` can appropriately assign participants' responses. Note: two slides can have the same title, but note that clicking macros on either slide will write to the same column, so only do this if you're okay overwriting responses or if participants will only responding on one of the slides (eg the 2nd step of a 2 step measure).
5. If you are using the `measure_buttonName` series of macros, set object names in PowerPoint via Selection Pane: Home > Editing > Select > Selection Pane. Double click an object in the Selection Pane to edit its name.

![PowerPoint slide with pictures of blue berries and pink berries, and textbox reading "measure3" just off screen. Outline view is open on the left, showing slide titled measure 2. Selection pane is open on the right, showing shapes named blueberries, pinkberries, Title 3, and Title 3.](/readme_images/slide.png)
_A typical slide._ It has a Title, here just off-screen (`measure2`), that gives the slide its title in Outline View. Each response button (here, two pictures) is named in Selection Pane as whatever the response text should be (here, "blueberries", "pinkberries").

6. Link your objects to whatever macros you want to run when they are clicked. Insert > Action > Mouse click > Run macro > select your macro (e.g. `measure_buttonName_advance1`). There's no easy to way to see at a glance if an object is linked to a macro, besides trying to reinsert the macro link, so make sure everything that you want linked is linked! Particularly if you are not auto-advancing slides, I recommend you check "Play sound" as well, so the button makes a sound (e.g. "Click") when clicked (to minimize confusion about whether it was clicked already or not).

![PowerPoint menu showing options for action button, including Run macro and Play sound](/readme_images/linkToMacro.png)

_Linking a macro._ Here the object is linked to the `measure_buttonName_advance1` macro. Note the "Play sound" option if you'd like to play a sound, in addition to running the macro, when the object is clicked.

7. Now open VBA: Developer menu > Visual Basic.  Click Tools at the top > References > make sure "Microsoft Scripting Runtime" is checked. (Dictionaries are not native to VBA, so this makes sure VBA can reference its home environment, Microsoft Scripting Runtime.)
8. Go over to the left-hand Project Explorer sidebar. (If you don't see it, Ctrl + R, or View > Project Explorer). Click `Module 1` to bring up the main code. Optional: customize any code as desired.
  - If you are on a Mac, edit the filepath in the `SaveToExcel` macro to Mac filepath syntax.
  - If you are using `reset` macros, make sure to manually specify which slide to reset target objects, since that is currently hard coded.

![View in VBA looking at Module 1, with SaveToExcel macro selected](/readme_images/VBA.png)
_A typical view in VBA._ Here we are in the code for `Module 1`, specifically the `SaveToExcel` macro. Note that the Project Manager sidebar is at the top left.

9. In the left-hand Project Manager sidebar, click `UserForm` to customize the userform associated with setup. Right-click `UserForm` > View Object to [edit the fields and aesthetics of the form](https://docs.microsoft.com/en-us/office/vba/powerpoint/how-to/create-custom-dialog-boxes) (View > Toolbox to insert new fields: `TextBox` accepts any value, `ListBox` requires selecting from pre-specified values, `ComboBox` suggests pre-specified values but accepts other values too). Right-click `UserForm` > View Code to edit the code behind the form, including how the values from the form are being saved to the `data` dictionary.

![View in VBA looking at the UserForm object](/readme_images/VBA_UserForm_object.png)
_Viewing the UserForm object._ Here we are looking at the `UserForm` object, and can move around the fields, change how the fields look, and change how the form generally looks. At right we have the Toolbox, with which we can add new fields (View > Toolbox if you can't see the Toolbox menu). At bottom left we have the Properties menu, where we can see that `condition` is a `ListBox` object (`ListBox` only accepts specified values).

![View in VBA looking at the UserForm code](/readme_images/VBA_UserForm_code.png)
_Viewing the UserForm code._ Here we are looking at the `UserForm` code, specifically how the UserForm is initialized. Note that the form field `condition` is initialized with specified values for `condition` ("condition1", "condition2", and "condition3") that users can choose from.

# Running participants live
- **Make sure `data.xlsx`, the Excel data sheet, is closed** so PowerPoint can edit it.
- Click "Setup" to run `Setup` and fill out the setup form.
  - It's recommended you fill out the setup form before the participant arrives, so the participant/guardian are blind to condition. You may exit slideshow after completing the setup form, because `Setup` will save everything to Excel (using `inProgress_SaveToExcel`). When the participant arrives and you restart the slideshow, be sure to click "resume" (`inProgress_resume`) to pick up from where you left off.
  - Alternatively, you can fill out the form with the participant present, but be aware that the participant/guardian may not be blind to condition if they see you filling out the form. Click "pause share" in Zoom to freeze the screen share, click "Setup" to fill out the form, and then click "resume share" once you're done with the form. The participant will see a frozen screen while you're filling it out.
- If you use 2 monitors and see your slideshow on both (e.g. sharing slides, not sharing presenter view), pop-ups will appear on whichever window you click the button. So if you click pop-up buttons like "Setup" on the non-shared screen, pop-up will appear on non-shared screen.
- You can run your slides and click macros in any order or number of times. Note that clicking macros(s) on the same slide (or different slides with the same Title) multiple times will overwrite the same item in `data`, since they share the same `key` (slide title). The *last* pressed macro will store the final response.
- **`data` will be lost if you exit slideshow before running a `SaveToExcel` macro**.
  - To jump between *slides* without exiting slideshow, right-clicking anywhere on the slide in slideshow > See all slides > select a slide.
  - To switch between *windows* without exiting slideshow, tab out of slideshow using `Alt + Tab` on Windows or `Cmd + Tab` on Mac, or try swiping left or right on your mousepad with 3 fingers.
  - If the session is ended early, and won't be picked up, jump to the last slide and run `SaveToExcel_end`. Otherwise, check that "in_progress" is set to "no".
  - Record your sessions, and be prepared to reference your recording as a backup in case of data loss/premature exit.
- You may run multiple participants in the same PowerPoint session, since the code will autoreset the `data` dictionary during initial setup, and `data` is also cleared when you exit slideshow.


# Editing the VBA code
Open VBA in PowerPoint: Developer > Visual Basic.

Tips for editing in VBA:
- Useful functions for dealing with the `data` dictionary:
  - `data(key) = item` stores the item under the key. If the key already exists, the previous item is overwritten. If the key did not exist, it will create a new pair.
  - `data(key)` calls the item stored under the key.
  - `data.Keys` calls all the keys, and `data.Items` calls all the items. Remember VBA is 0-indexed, so `data.Items(0)` calls the 1st item, `data.Items(1)` the 2nd, and so forth.
- Remember to always declare your variables (e.g. `Dim`, `Public` if you want to access them throughout your code, or `Private`) before you initialize them.
- Remember to end your loops with the corresponding `End` (e.g. an `If` requires an `End If`).
- Check that your types match/are appropriate. `CStr()` coerces non-`String` types into a `String`, and [here are some other coersion functions](https://docs.microsoft.com/en-us/office/vba/language/concepts/getting-started/type-conversion-functions).
- VBA does not automatically wrap your code. Add ` _ ` (space _ space) at the end of a line to continue code on the next line. Or write your code in a code editor like [Atom](https://atom.io/) (Atom: File > Settings > Packages > install `language-vba` for VBA syntax highlighting).
- Comment your code! `'` begins a comment.
- Test run test run test run. Build and test new code in small steps so you can more easily isolate problems when your code breaks.

Tips for PowerPoint when editing in VBA:
- In PowerPoint, you can find/edit names of objects in Home > Select > Selection Pane, and titles of slides in View > Outline View. Note that slide titles (indicated by title text and in Outline View) are apparently different from slide names (an internal name indicated when you open a slide in VBA, somewhat difficult to change).
- If you rename a macro, you'll need to relink the buttons that used to reference that macro to that macro. There is no easy way to tell at a glance if a button is linked to a macro or not, so be careful, double-check your links by trying to reinsert a link, and always test run.

Debugging VBA:
- Debug > Compile VBAProject is your friend. When you make changes, run this and make sure it doesn't freak as first pass for errors.
- Watch > Add Watch to keep an eye on select variables while you're running code blocks, like the R environment.
- If you're running into some bug, try printing a key variable with `print` (print to the "Immediate" window at the bottom of your VBA screen), or have it pop out in a `MsgBox` at a key point in your code to see what's wrong.
- If the Excel file did not update, or if a button that should have done something didn't do something, check your code for errors.
- If the Excel file is locked for editing when you open it, force quit Excel from Task Manager and check your code for errors.

If you're stuck, here are some VBA resources. Most info online is about Excel, but it will usually transfer with minor tweaks.
- Google "VBA" and whatever you're stuck on.
- [Microsoft Office VBA documentation](https://docs.microsoft.com/en-us/office/vba/api/overview/powerpoint)
- On scripting dictionary specifically: [Microsoft documentation on the Dictionary object](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dictionary-object), [Excel VBA Dictionary](https://excelmacromastery.com/vba-dictionary), [dictionary vs collection vs array](https://stackoverflow.com/questions/32479842/comparison-of-dictionary-collections-and-arrays)
- StackExchange
- various ExcelVBA help forums

# Future improvements
Here are some improvements I'm thinking to do. Definitely feel free to make edits and improvements to the code yourself too!
- Automatically detect which slides are resource allocation slides in `reset_allocation`.
- Automatically generate `participantOfDay` by referencing previous files in `data` to fully automatically generate `file`.
