# Why are the notes only on the bottom?

In recent times I've mostly used LibreOffice Impress, or Google Slides, but sometimes for reasons, I need to use Powerpoint. For whatever reason, when editing a slideshow in the main editing view in Powerpoint, the notes are on the bottom of the screen and the panel they are in seems to be  not undockable or moveable.

Solutions may include:
- use LibreOffice - LibreOffice has some features that I think make it better than Powerpoint, except in this case it is in fact worse, you can only use the Notes view (which is in Powerpoint too). In both cases there is a Notes master but this affects how the notes pages print... This applies for LibreOffice Impress v7.2.1 which was installed on my machine
- use Google Slides? I haven't looked recently
- use Keynote? I haven't used this much

Instead, it is possible to use VBA to show two Windows, one in normal edit view and one in notes edit view, and have them side by side, and then sync pages changes in the normal view into the notes view.

# How To Do It

1. Temporarily save your presentation as a `.pptm` (Powerpoint macro-enabled file)
2. Open the macro editor (Tools > Macros > Visual Basic Editor) aka VBE
3. Create a class module
    - right click on `VBAProject (your filename)` in the VBE
    - Insert > Class Module
    - Paste the content of `pptx_add_notes_window_sync/class.vba` into the new module, it will be called `Class1` by default
4. Create a normal module
    - right click on `VBAProject (your filename)` in the VBE
    - Insert > Module
    - Paste the content of `pptx_add_notes_window_sync/module1.vba` into the new module, it will be called `Module1` by default
5. Save your pptm file
6. To enable, both now, and also needed **each time you reopen the file**,
    - Tools > Macro > Macros...
    - Run the macro called `InitializeApp`
7. **When you are finished with your presentation, remember to resave your file as a pptx if needed**

# Extra goodies

- There is a macro called `PositionWindows`, run this after activating to create a second notes view window and fit them side by side
- There is a macro called `CountNotHiddenSlides` which will count the number of slides that are _not_ hidden

# Other ways

- I could create a Powerpoint addin, or much around with macro signing, but this is nice and simple given it is not something I need to do a lot.

# Note

Only tested on Powerpoint for Mac version 16.74
