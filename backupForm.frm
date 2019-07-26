VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} backupForm 
   Caption         =   "Backup to PST"
   ClientHeight    =   3564
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6252
   OleObjectBlob   =   "backupForm.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "backupForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub backupBtn_Click()
    If sourceTxtBox.Value = vbNullString And destinationTxtBox.Value = vbNullString Then
        MsgBox "All fields are required!"
    ElseIf sourceTxtBox.Value = vbNullString Then
        MsgBox "Select the source location!"
    ElseIf destinationTxtBox.Value = vbNullString Then
        MsgBox "Select the target location!"
    ElseIf sourceTxtBox.Value = destinationTxtBox.Value Then
        MsgBox "Cannot create backup to the same location!"
    ElseIf fromParentFolder = toParentFolder Then
        MsgBox "Cannot create backup to the same folder!"
    Else
        Call moveData(fromParentFolder, fromNameFolder, toParentFolder, toNameFolder, toPstParentFolder, toPstNameFolder)
    End If
End Sub

Private Sub cancelBtn_Click()
    Unload backupForm
End Sub

Private Sub createPSTBtn_Click()
    Call createPSTFile
End Sub

Private Sub sourceBtn_Click()
    Call selectSource
End Sub

Private Sub targetBtn_Click()
    Call selectTarget
End Sub

Private Sub UserForm_Click()

End Sub
