Attribute VB_Name = "ImportSfdcMetadata"
'First version released 29 Dec 2018 by Adrienne D. Millican, admillican08@gmail.com

Option Explicit
' Double quote constant
Public Const DBQ As String = """"

' Salesforce namespace
Public Const SXMLNS As String = "xmlns=" & DBQ & "http://soap.sforce.com/2006/04/metadata" & DBQ

'The metadata names
Public Const PR As String = "Profile"
Public Const PE As String = "PermissionSet"

'
' DisplaySfdcUserForm Macro
'
' Displays form that allows you to select which SFDC metadata you want to import into Excel
' Keyboard Shortcut: Ctrl+Shift+U; the user can always remap it

Public Sub DisplaySfdcUserForm()
Attribute DisplaySfdcUserForm.VB_ProcData.VB_Invoke_Func = "U\n14"

 ChooseSfdcMacroForm.Show
    DoEvents
End Sub


' createExcelWorkbooksImportData Macro
'
' Allows users to import either Salesforce profiles or permission sets into Excel for viewing, filtering, and overall
' a more human-friendly and readable format

Public Sub createExcelWorkbooksImportData(ByVal mdName As String)

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim mdFilePathArr() As String
    Dim mdDirPath, exDirPath, xsdFilePath, lcMdType, fileExt, secondExt, templDirPath, templFilePath, eXmlNs As String
    Dim strLen, count As Integer
    Dim objFso As FileSystemObject
      
    ' lowercase the metadata type name for use in extensions
    
    lcMdType = LCase(mdName)
    
    ' for use in attaching to the XML files which will then be imported into Excel. My thanks go out to a Udemy instructor whose course on Xpath and Excel I took a few lessons from
    
    eXmlNs = "xmlns:xsi=" & DBQ & "http://www.w3.org/2001/XMLSchema-instance" & DBQ & " xsi:noNamespaceSchemaLocation=" & DBQ & "..\ExcelFiles\xsdFile\" & lcMdType & ".xsd" & DBQ
    
    ' Get a reference to the folder and path of the macro-enabled workbook (technically not a template but effectively acts as one)
    templDirPath = Application.ActiveWorkbook.Path
    templFilePath = Application.ActiveWorkbook.FullName
    
    ' Create a file system object -- extremely handy
    Set objFso = CreateObject("Scripting.FileSystemObject")
    
    ' One of the typical file extensions for Salesforce metadata
    fileExt = "." & lcMdType

    ' Allow the user to select one to several profile or permission set files; user has to choose which type and cannot mix
    ' The chooser returns an array representing the selected filepaths
    
    mdFilePathArr = selectMultipleFiles(mdName, mdFilePathArr, fileExt)
    
    ' Much thanks to https://stackoverflow.com/questions/206324/how-to-check-for-empty-array-in-vba-macro for this!
    ' If the array is empty, exit with an error, otherwise proceed with at least one file
    If (Not mdFilePathArr) = -1 Then
        MsgBox "Error loading " & mdName & " metadata files.", vbCritical
        GoTo Abnormal
    End If
        
    ' The user has to select the location of the xsd file to validate the Salesforce metadata XML against
    
    xsdFilePath = selectedSingleFile(mdName & " XSD ", ".xsd", mdName & ".xsd")
    
    ' If the user selected no XSD or the wrong XSD, exit with an error
    
    If xsdFilePath = "" Or objFso.GetFileName(xsdFilePath) <> mdName & ".xsd" Then
        MsgBox "Problem with the XSD file. Please check setup and try again.", vbCritical
        GoTo Abnormal
    End If

    On Error GoTo Abnormal
        
    ' The output directory is created as a subdirectory of the directory in which the macro workbook resides.
    ' If it exists already, users are warned their existing files are going to be replaced by newly created ones.
    
    exDirPath = templDirPath & "\Excel_" & mdName & "s\"
    Call createSubdir(objFso, exDirPath, True)
    
    ' Put this here in case users want to go back to the operating system to move files around.
    ' It will cause an abnormal exit, but they can just reopen the form and try again.
    
    DoEvents
             
    ' OK, now the fun begins. First, create XML files based on the Salesforce metadata that Excel will accept.
    ' Replace the namespace and also make sure the file ends in .xml. These files are only temporary
    
    Call createXmlFiles(objFso, exDirPath, mdFilePathArr, SXMLNS, eXmlNs, lcMdType)
    
    DoEvents
    
    ' Now that XML files are created, create the Excel workbooks based off of the template of this macro-enabled workbook,
    ' import the XML files created in the prior subroutine, so their data maps to the appropriate XML map attached to the
    ' template workbook, and delete the intermediate XML file. Save each one as a non-macro workbook.
    
    Call createExcelFiles(objFso, exDirPath, templDirPath, mdName, fileExt)
    
    ' Now just finish up
    
    GoTo Normal
  
      
Abnormal:
     
     ' Abnormal exit. Something went wrong. Inform the user, release the file system object, and hopefully they
     ' can solve the issue and try again
     
     Set objFso = Nothing
     MsgBox "Error encountered when importing " & mdName & " file(s) into Excel. Check your setup and try again.", vbCritical
     Exit Sub
    

Normal:
    
    ' Close out each new workbook and return control to the macro-enabled workbook. Set it back to the first sheet.
    ' Release the file system object.
    
    Call closeAllWorkbooksExceptMain(templFilePath)
   
    Set wb = Application.ActiveWorkbook
    Set ws = wb.Sheets("Instructions")
    ws.Activate
    Set objFso = Nothing
   
      
End Sub

    '
' createSubdir Macro
'
' Creates a subdirectory based on supplied filepath. Boolean parameter allows it to display message or not

Private Sub createSubdir(objFso As Object, ByVal dirPath As String, ByVal dispMsg As Boolean)
  
    Dim objFolder As Scripting.Folder
    Dim folderName, parentFolder As String
    
   ' If the directory/folder already exists, get a file system reference to it. If a message is to be displayed, show
   ' a message indicating that existing files will be replaced.
   
    If objFso.FolderExists(dirPath) = True Then
        If dispMsg = True Then
            Set objFolder = objFso.GetFolder(dirPath)
            folderName = objFolder.Name
            MsgBox folderName & "\ folder already exists. Any existing file(s) in folder based on your selected file(s) will be replaced.", vbExclamation
        End If
       
    ElseIf objFso.FolderExists(dirPath) = False Then
        
        ' If the directory/folder doesn't yet exist, create it and get a file system reference to it and its
        ' parent folder. If a message is to be displayed, show a message indicating that the new directory has been created
        ' and its location
        
        objFso.CreateFolder (dirPath)
        Set objFolder = objFso.GetFolder(dirPath)
           
        If dispMsg = True Then
            folderName = objFolder.Name
            parentFolder = objFolder.parentFolder.Name
            MsgBox "Creating subdirectory for Excel worksheets in " & parentFolder, vbInformation
        End If
    End If
    
End Sub

      
' createXmlFiles Macro
'
' Creates XML files for use in the importing of Salesforce metadata into Excel worksheets. These XML files are an intermediate step.
' After all, the Salesforce metadata files are themselves XML. These XML files have a different namespace that conforms to what the Excel
' template is expecting when trying to map data to the Profile_Map and/or PermissionSet_Map. They also end in ".xml". These XML files
' are deleted once they are imported into Excel workbooks and the workbooks are saved.

Private Sub createXmlFiles(objFso As Object, ByVal exDirPath As String, mdFilePathArr() As String, ByVal oldNS As String, ByVal newNS As String, ByVal mdExt As String)

    Dim i, j As Integer
    Dim objFile As Scripting.File
    Dim objR, objW As Scripting.TextStream
    
    j = 0
    
    Dim fileText, oldPath, oldName, oldExt, xmlPath As String

    ' Try to iterate through files, just in case if one is bad, others could still be processed
    
    On Error Resume Next
    
    ' Turn off alerts and updating so you don't see blinking screens
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Iterating through each of the metadata files the user selected earlier.
    
    For i = LBound(mdFilePathArr) To UBound(mdFilePathArr)
        
        oldPath = mdFilePathArr(i)
        
        ' Get an actual file object reference to each file
        Set objFile = objFso.GetFile(oldPath)
        oldName = objFile.Name
        oldExt = objFso.GetExtensionName(oldPath)
        
    ' Technically the user could select a .profile.xlsx or .permissionset.xlsx file by mistake,
    ' so just in case this screens those out
    
        j = InStrRev(oldPath, ".xlsx")
        If j > 0 Then
            MsgBox oldName & " has an incorrect file type. Skipping.", vbInformation
            GoTo LoopEnd
        End If
            
        
        ' If they are indeed metadata files, they can be opened as text files
        Set objR = objFso.OpenTextFile(oldPath)
        
        'Read all of the text into a text stream
        fileText = objR.ReadAll
        
        ' Close the reading stream
        objR.Close
        
        ' Replace the namespaces in the String
        fileText = Replace(fileText, oldNS, newNS)
        
        ' If the file already ended in xml, i.e., was named "-meta.xml", then retain that extension
        If oldExt = ".xml" Then
            xmlPath = exDirPath & oldName
        
        'Otherwise, give it the extension ".xml"
            ElseIf oldExt <> ".xml" Then
                xmlPath = exDirPath & oldName & ".xml"
        End If
        
        ' Use a fileWriter object to create an XML file in the folder where the workbook will eventually go
        Set objW = objFso.CreateTextFile(xmlPath)
        
        ' Write the contents of the read-and-replaced file to the new XML file
        objW.Write (fileText)
        objW.Close
        DoEvents
         
LoopEnd:
    
    ' Go to the next file
    
    Next
        
    
    ' Turn Application screen updating and alerts back on
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    
End Sub

' createExcelFiles Macro
'
' Takes the XML files generated by createXmlFiles and imports them into Excel workbooks based on this macro-enabled workbook,
' then saves each workbook as a regular, non-macro-enabled workbook with the name of the original metadata file

Public Sub createExcelFiles(objFso As Object, ByVal exDirPath As String, ByVal currFilePath As String, ByVal mdName As String, ByVal fileExt As String)
       
   Dim wb As Workbook
    Dim wkSheet As Worksheet
    Dim oExFldr As Scripting.Folder
    Dim oExFile As Scripting.File
    Dim oldName, oldPath, newName, exFilePath As String
    Dim i, j, strLen As Integer
    Dim mapName As String
    Set oExFldr = objFso.GetFolder(exDirPath)
    
    ' The name of the metadata map to use -- Profile or Permission set. This macro-enabled workbook has mappings for both
    mapName = mdName & "_Map"
    
    i = 0
    j = 0
    
    ' turn off screen updating so users don't see new workbooks flashing up on screen as each is created
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' try to continue to create new files even if the subroutine encounters failure
    
    On Error Resume Next
    
    For Each oExFile In oExFldr.Files
      
        ' loop through each file in the directory
        
        oldName = oExFile.Name
        oldPath = exDirPath & oldName
        
        i = InStrRev(oldName, "-meta.xml")
        j = InStrRev(oldName, ".xml")
        
        ' This is done so as to ignore any existing .xlsx files in the directory
        ' and only to operate on the .xml files
        
        If i > 0 And j > 0 Then
            newName = Left(oldName, i - 1)
            GoTo makeFiles
            
            ElseIf i = 0 And j > 0 Then
                newName = Left(oldName, j - 1)
                GoTo makeFiles
                
            ' if the file encountered doesn't have an .xml ending, skip it
            ' and go to the next file
            
            Else: GoTo LoopEnd
        End If
                  
makeFiles:
        
        ' The new name for the Excel workbook
        
        exFilePath = exDirPath & newName & ".xlsx"
        
        ' Creating the new Excel workbook
        Set wb = Workbooks.Add(currFilePath)
        
        ' Make it the active workbook
        Set wb = Application.ActiveWorkbook
        
        ' Import the XML file data into the new workbook according to the correct metadata map
        wb.XmlMaps(mapName).Import Url:=oldPath
        
        ' Save the new Excel workbook as a regular workbook without macros
        wb.SaveAs exFilePath, xlOpenXMLWorkbook
        
       ' Do a little formatting
        For Each wkSheet In wb.Worksheets
            wkSheet.Range("A1") = newName
            wkSheet.Range("A1").ColumnWidth = 40
            wkSheet.Range("A1").RowHeight = 40
            wkSheet.Range("B1").ColumnWidth = 20
            wkSheet.Range("C1").ColumnWidth = 20
            
        Next
        
        ' Do a little more cleanup and formatting, remove the instruction sheet,
        ' but leave the original macro-enabled workbook alone!
        
        If wb.FileFormat = xlOpenXMLWorkbook Then
            wb.Sheets("Instructions").Delete
            wb.Sheets("BasicInfo").Range("A1").ColumnWidth = 60
            wb.Sheets("BasicInfo").Range("B1").ColumnWidth = 45
            wb.Sheets("BasicInfo").Range("C1").ColumnWidth = 45
            wb.Sheets("ObjectPermissions").Range("D1").ColumnWidth = 20
            wb.Sheets("ObjectPermissions").Range("E1").ColumnWidth = 20
            wb.Save
        End If
        
        
        ' Now delete the XML file that Excel has imported into a workbook
        objFso.DeleteFile (oldPath)
        DoEvents
        
LoopEnd:
    
    ' Go to the next file
    Next
    
    ' Turn screen updating and alerts back on
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' Message user that the workbooks were created
    MsgBox "Excel workbooks for " & mdName & " files created in " & exDirPath, vbOKOnly
    
End Sub

  
'Some code taken from https://software-solutions-online.com/excel-vba-open-file-dialog/

' selectedSingleFile Macro
'
' A file dialog that allows a user to select a single file and returns the file path of that file

Private Function selectedSingleFile(ByVal fileType As String, ByVal fileExt As String, ByVal initVal As String) As String
    
    Dim intChoice As Integer
    Dim filter As String
    Dim fileSelector As FileDialog
    
    ' Filters out what the user can select
    
    filter = "*" & fileExt
    
    ' Only allow the user to select one file, filter the type of files suggested
    
    Set fileSelector = Application.FileDialog(msoFileDialogOpen)
    
    ' Set the properties of the file selection dialog
    
    With fileSelector
        .Filters.Clear
        .AllowMultiSelect = False
        .Title = "Select " & fileType & " file"
        .InitialFileName = initVal
        .Filters.Add fileType & "files", filter, 1
    End With

    ' Show the dialog and capture the user's selection
    
    intChoice = fileSelector.Show
    If intChoice <> 0 Then
    
        selectedSingleFile = fileSelector.SelectedItems(1)
        
    Else: selectedSingleFile = ""
    End If

End Function

'Some code taken from https://software-solutions-online.com/excel-vba-open-file-dialog/

' selectMultipleFiles Macro
'
' A file dialog that allows a user to select multiple files and returns an array of strings representing the file paths of
' the selected files


Private Function selectMultipleFiles(ByVal fileType As String, filePathArray() As String, ByVal fileExt As String) As String()
    
    Dim i, intChoice, fileCount As Integer
    Dim filter As String
    Dim mFileSelector As FileDialog
  
    ' Filters out what the user can select
    filter = "*" & fileExt & "*"
    fileType = LCase(fileType)
    
    'only allow the user to select one file, filter the type of files suggested,
    'default the choosing directory to the same one the template is in
    
    Set mFileSelector = Application.FileDialog(msoFileDialogOpen)
    
    With mFileSelector
        .AllowMultiSelect = True
        .Title = "Select " & fileType & " file"
        .Filters.Clear
        .InitialFileName = ""
        .Filters.Add fileType & " files", filter, 1
    End With
    
    'make the file dialog visible to the user
    
    intChoice = mFileSelector.Show
    
    'determine what choice the user made

    If intChoice <> 0 Then
    
        ' Initialize the array to an actual size based on the number of files selected
        
        fileCount = mFileSelector.SelectedItems.count
        ReDim filePathArray(1 To fileCount)
        
        ' Assign the filepath of each selected file to the array
        For i = 1 To fileCount
           
            filePathArray(i) = mFileSelector.SelectedItems(i)
            Next i
        
        ' Tell the user how many files were picked
        MsgBox fileCount & " files selected", vbInformation
    
    End If
    
    ' return the array
    selectMultipleFiles = filePathArray()
    

End Function

' closeAllWorkbooksExceptMain Macro
'
' Closes all open workbooks and returns control to the macro-enabled workbook


Private Sub closeAllWorkbooksExceptMain(ByVal templFilePath As String)
    
    Dim mainWb, otherWb As Workbook
    Set mainWb = Workbooks.Open(templFilePath)
    Set mainWb = Application.ActiveWorkbook
    
    Application.ScreenUpdating = False
    
    
    ' This is so all of the other workbooks created based on profiles or permission sets
    ' don't stay open and have to be manually closed off. It also means you can keep invoking
    ' the macro form as you need to
    
    For Each otherWb In Application.Workbooks
        If Not (otherWb Is mainWb) Then
            otherWb.Close SaveChanges:=True
        End If
    Next
    
    Application.ScreenUpdating = True
    
    
End Sub
