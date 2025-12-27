# save-contacts-outlook

1. Create `%USERPROFILE%\save-contacts-outlook`
2. Save all files in that directory
3. Make sure python is installed
4. Install deps with `pip install pywin32 eassygui`
3. Create `%LOCALAPPDATA%\save-contacts-outlook`
4. Move `run.bat` to `%LOCALAPPDATA%\save-contacts-outlook`
5. Go into MS Outlook and press `ALT + F11`
6. Import the `.cls` file with right click on the left side. And save and close.
7. Outlook -> File -> Options -> Customize Ribbon
8. Add a new group at the right side and name it.
9. Choose Macros from dropdown and add `SendSelectedMailToPython`

