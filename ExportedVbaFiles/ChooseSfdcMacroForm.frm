VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChooseSfdcMacroForm 
   Caption         =   "Choose a Salesforce Profile or Permission Set Macro"
   ClientHeight    =   2820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6345
   OleObjectBlob   =   "ChooseSfdcMacroForm.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "ChooseSfdcMacroForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_OK_Click()
    
    Dim VsCode As Boolean
    
    
    If op_ProcessProfilesInExcel = True Then
        Call createExcelWorkbooksImportData(PR)
        Unload Me
    End If

    If op_ProcessPermSetsInExcel = True Then
        Call createExcelWorkbooksImportData(PE)
        Unload Me
    End If
    

End Sub


Private Sub btn_Cancel_Click()
    Unload Me
    
End Sub

