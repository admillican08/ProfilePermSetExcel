Hello all!

Use the ProfilePermSetImportExport.xlsm Macro-Enabled Microsoft Excel workbook to export and import data to and from Salesforce Profiles and Permission Sets using Microsoft Excel. This allows you to view, filter, and compare the data in a much more human-friendly way than the Salesforce UI provides. While there are many handy Web browser extensions that allow you to compare profiles and permission sets, there are, to my knowledge, no free tools that allow you to reports with actual data that can be manipulated and compared...until now!



*) Required Versions of Applications:

Salesforce Metadata API Version 23 or later, Microsoft Excel version 2007 or later (any version using .xlsx)



*) What This Workbook Lets You Do

This template helps to make Salesforce profile and permission set data easier to work with. It has mappings and macros to pull in XML data from Salesforce profile and permission set metadata component files. You can also use macros in the workbook to create and export profiles and permission sets from Excel, albeit please do this with caution. Please note that you may have to do further validation and alteration of any profile created and exported using this workbook before you can successfully deploy it to Salesforce.

The Excel macro-enabled Workbook acts on are Salesforce metadata component files have been previously retrieved from Salesforce using tools such as the Ant Migration Tool .jar, the Eclipse Force.com IDE plugin, or the Visual Studio Code Salesforce extension pack. The file extensions of these XML files should be .profile, .permissionset, .profile--meta.xml, or .permissionset-meta.xml.

Please note this Excel workbook does not connect directly with Salesforce in any way. Either you or a friendly Salesforce Administrator or Developer for your org must have already retrieved the profile or permission set files from Salesforce prior to using this macro-enabled workbook.


*) Files Included:

* The macro-enabled Excel workbook ProfilePermSetImportExport.xlsm
* Profiles.xsd (the Salesforce Profile Schema file) -- required for Profiles
* PermissionSet.xsd (the Salesforce Permission Set Schema file) -- required for Permission Sets
* ImportSfdcMetadata.bas
* ChooseSfdcMacroForm.frm
* ChooseSfdcMacroForm.frx
* some sample .profile and .permissionset files to try out.
* a sample package.xml manifest to retrieve all profiles and permission sets from an org



*) OK, So How Do I Use This Thing?

1) Make sure you have your .xsd files and some .profiles or .permissionsets handy on your PC.
2) Open the ProfilePermSetImportWkbk.xlsm in Microsoft Excel version 2007 or later.
3) Press Ctrl + Shift + U. A Macro Selection Form should display.
3) The rest is hopefully self-explanatory.

If, for some reason, they keystroke Ctrl + Shift + U fails to work, display the Developer tab in Excel if it is not already displayed, select Macros, and run the DisplaySfdcUserForm macro. You can also use the Macros dialog to assign a different keystroke to run the DisplaySfdcUserForm macro if you wish.



*) Macro Security Issues:

Depending on your Excel and PC security settings, you may have to click a button to allow the macros in this template to run. If your security settings are strict enough, you may not be able to run macros from an outside party at all. You may be able to change them by displaying the Excel Developer tab if it is not already visible, then selecting Macro Security, and changing the macro settings to "Disable all macros with notification", which still typically gives the option to enable macros by clicking a button.



*) Showing the Microsoft Excel Developer Tab:

From the main menu, select File > Options to open the Excel Options dialog box. Select Customize Ribbon. Select Main Tabs from the right drop-down menu, and select the Developer check box.



*) Re-creating the Template If Security Won't Let Me Use Macros From Outside:

If you are unable to run macros from this template because it originates from an outside source, then you should be able to re-create your own version of this template on your own PC in Excel using the two exported VBA text files included with this project:

1. Open a new, empty workbook in Excel.
2. From the Excel Developer tab, select the Visual Basic icon to open the VBA Editor.
3. Select the new Excel workbook's name--typically it will be VBAProject(Book#) in the Project editor window, right click and select Import File.
4. Select the .bas file. Repeat and this time import the .frm file. 
5. Select Tools > References, and find and check Microsoft Scripting Runtime (leave other checked checkboxes as-is).
6. Select File > Close and Return to Microsoft Excel
7. Save the open Excel workbook as either a macro-enabled workbook (*.xlsm) or a macro-enabled template (.xlst).



Happy using, and please let me know what you think, what bugs you found, and so on!

Adrienne D. Millican

3X Certified Salesforce Professional

13 March 2019

admillican08@gmail.com
