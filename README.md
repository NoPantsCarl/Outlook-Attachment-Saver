# Outlook Attachment Saver

This script facilitates the quick saving of Outlook attachments in a batch, rather than individually, through a macro.

### Instructions:

1. Modify Line 90 to specify the folder location where you want to save all attachments: `strFolderPath = "C:\Users\nopantscarl\Desktop\Needs work"`
2. Navigate to "File" in Outlook and select "Options."
3. Choose "Outlook Options," then "Customize Ribbon."
4. On the right side, locate and enable the "Developer" tab.
5. Return to the main Outlook window to access the developer ribbon.
6. Click on the "Macro Security" button.
7. In the subsequent window, select "Notifications for all macros" and confirm.
8. Press Alt + F11 to open the Visual Basic for Applications (VBA) editor.
9. Click on "Insert" > "Module" and paste the provided code into the module, or import it at this stage.
10. Press F5 to display the modules and click "Run" to execute the macro.

### Quick Access Integration:

1. Follow the aforementioned steps to access Outlook options.
2. Choose the "Quick Access Toolbar" tab.
3. Select "Macros" from the "Choose commands from" dropdown menu.
4. From the macro list, select the desired macro (either the name you assigned or the imported name).
5. Click the "Add >>" button at the center.
6. Click "OK" to confirm.
7. To execute the macro from the Quick Access Toolbar, simply click its corresponding button.

