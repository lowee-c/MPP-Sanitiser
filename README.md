# MPP-Sanitiser

A VBA macro for Microsoft Project desktop that strips identifying detail from a project file so it can be shared, used as test data or sent outside the engagement.
 
**Run it on a copy of the file. There is no reliable undo.**
 
## What it does
 
| Area | Action |
|------|--------|
| Project title | Sets the Title document property to `SANITISED` (shown on the project summary task) |
| Resource names | Any resource that has an email address is renamed `Resource 001`, `Resource 002` and so on. Existing names are checked first so there are no duplicates |
| Resource contact fields | Clears email address and Windows user account on every resource |
| Initials | Sets Initials to `R` on every resource that was renamed |
| Hyperlinks | Clears the hyperlink fields on every resource |
| Notes | Clears notes on every resource, every task, every assignment and the project summary task |
 
When it finishes it shows a summary message box (resources processed, renamed, tasks with notes cleared, failures). The detailed log, including the old name, email and new name for each renamed resource, is written to the Immediate window (Ctrl+G in the VBA editor).
 
## What it does not do
 
- It does not rename the file. `ActiveProject.Name` is read-only. Use File > Save As to give the sanitised copy a generic name.
- Resources with no email address keep their real names.
- Task names, custom fields (Text1 to Text30, outline codes and lookup tables), resource group and code fields, task hyperlinks, rates and costs, calendars, linked or inserted projects, headers and footers, and the Author, Manager, Company and Comments document properties are not touched.
- It only works in Project desktop. Project for the web does not support VBA.
## Settings
 
All settings are constants at the top of the module.
 
| Constant | Default in the supplied file | Meaning |
|----------|------------------------------|---------|
| `DRY_RUN` | `True` | `True` logs what would change and writes nothing. Set to `False` to apply changes |
| `NAME_PREFIX` | `"Resource "` | Prefix for new resource names |
| `PAD_DIGITS` | `3` | Number of digits in the sequence (`001`) |
| `CLEAR_INITIALS` | `True` | Reset initials on renamed resources |
| `CLEAR_HYPERLINK` | `True` | Clear resource hyperlinks |
| `CLEAR_NOTES` | `True` | Clear all notes |
| `PROJECT_TITLE` | `"SANITISED"` | Generic project title |
 
## Importing the macro into Microsoft Project
 
The macro is supplied as `modAnonymise.bas`.
 
### One-off setup
 
1. In Project, go to **File > Options > Customize Ribbon**, tick **Developer** in the right-hand list, and select **OK**.
2. Go to **File > Options > Trust Center > Trust Center Settings > Macro Settings**. Choose **Disable all macros with notification** (or **Enable all macros** if your policy allows it) and select **OK**. If your organisation locks this setting, ask IT to allow macros for this task.
### Import the module
 
3. Open the **copy** of the project file you want to sanitise.
4. Press **Alt+F11** to open the Visual Basic Editor (or use **Developer > Visual Basic**).
5. In the Project Explorer pane, find **ProjectGlobal (Global.MPT)** and select it. If the pane is hidden, press **Ctrl+R**.
6. Choose **File > Import File...**, browse to `modAnonymise.bas` and select **Open**. A module named `modAnonymise` appears under **Modules**.
Import into **ProjectGlobal** rather than into the project file itself. A module stored inside the project file travels with it when you share it, which defeats the purpose of sanitising it. Because the macro lives in the global template, it is available to any open project.
 
### Run it
 
7. Open `modAnonymise` in the editor and check that `DRY_RUN` is `True` (the supplied default).
8. Open the Immediate window with **Ctrl+G**.
9. Click inside the `AnonymiseResourceContacts` procedure and press **F5**. Alternatively, in Project choose **Developer > Macros**, select `AnonymiseResourceContacts` and select **Run**.
10. Review the log in the Immediate window and the summary message box.
11. Change `DRY_RUN` to `False`, run it again, and confirm the prompt.
12. Check the result (see below), then use **File > Save As** to save the sanitised copy under a generic file name.
### Tidy up
 
13. Back in the editor, right-click `modAnonymise` and choose **Remove modAnonymise...**. Select **No** when asked to export it first.
14. Close the editor. Project may ask to save changes to Global.MPT, which is expected.
If you prefer, you can copy the code and paste it into a new module (**Insert > Module**) instead of importing the file.
 
## Checking the result
 
1. Open the Resource Sheet and confirm that names, email, Windows account and initials are generic.
2. Search the file for client names, staff surnames and your email domain.
3. For a stronger check, use **File > Save As** and save an XML copy. XML is plain text, so open it in an editor and search it for those same terms. Anything the macro missed will show up there. Delete the XML copy afterwards.
4. Review **File > Info** for the Author, Manager and Company properties and clear them if needed.
## Troubleshooting
 
- **Macro does not appear or will not run:** check the Trust Center macro setting, and that Project was restarted after changing it.
- **"Failures" is above zero:** the Immediate window lists the UID and the error for each one. Fields on locked or read-only items (for example, resources owned by a published enterprise file) cannot be changed.
- **Run it by mistake on the wrong file:** close the file without saving. If you saved it, restore it from the original, because there is no reliable undo.