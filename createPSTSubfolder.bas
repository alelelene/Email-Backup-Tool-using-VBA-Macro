Attribute VB_Name = "createPSTSubfolder"
Sub createFolder()
    Dim myolApp As Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim myFolder As Outlook.MAPIFolder
    Dim myNewFolder As Outlook.MAPIFolder
    
    Dim newProjectName As String
    Dim isNameValid As Boolean
    
    While isNameValid = False
    
    newProjectName = InputBox(prompt:=vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Enter folder name:", title:="Create folder in PST")
    Set myolApp = CreateObject("Outlook.Application")
    Set myNamespace = myolApp.GetNamespace("MAPI")
    
    
    
    Select Case StrPtr(newProjectName)
        Case 0
            Exit Sub
        Case Else
        If (newProjectName = vbNullString) Then
            MsgBox "Please enter a folder name"
            isNameValid = False
        Else
            isNameValid = True
        End If
    End Select
    
    Wend
    Set RootFolder = myNamespace.Folders(strDisplayName)
    Set myNewFolder = RootFolder.Folders.Add(newProjectName)
    toPstNameFolder = myNewFolder
    
    
End Sub
