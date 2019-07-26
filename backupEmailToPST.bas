Attribute VB_Name = "backupEmailToPST"
Public FSO As Object
Public myolApp As Outlook.Application
Public iNameSpace As NameSpace
Public fromFolder As Object
Public toFolder As Object
Public isValid As Boolean
Public fromParentFolder As Variant, fromNameFolder As Variant
Public toParentFolder As Variant, toNameFolder As Variant

Public Sub emailBackup_Click()
    backupForm.Show
End Sub

Sub selectSource()
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set myolApp = Outlook.Application
    Set iNameSpace = myolApp.GetNamespace("MAPI")
    
    isValid = False
    ' loop until a valid folder is selected
    Do While isValid = False
    
    ' select source folder
    Set fromFolder = iNameSpace.PickFolder
    If fromFolder Is Nothing Then
        GoTo ExitSub:
    End If
    Call getFromFolder(fromFolder)
    
    ' selected folder validation
    If fromParentFolder = "Mapi" Then
        MsgBox "Please select a specific folder!"
        isValid = False
    Else
        isValid = True
    End If
    Loop
ExitSub:
End Sub

Sub selectTarget()
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set myolApp = Outlook.Application
    Set iNameSpace = myolApp.GetNamespace("MAPI")
    
    isValid = False
    Do While isValid = False
    
    ' select destination folder
    Set toFolder = iNameSpace.PickFolder
    If toFolder Is Nothing Then
        GoTo ExitSub:
    End If
    Call getToFolder(toFolder)
     
     ' selected folder validation
    If toParentFolder = "Mapi" Then
        MsgBox "Please select a specific folder!"
        isValid = False
    Else
        isValid = True
    End If
    Loop
    
    isPstCreated = False
ExitSub:
End Sub

Sub getFromFolder(fromFld As MAPIFolder)
    ' gets the parent and name folder of the chosen source folder
    fromParentFolder = fromFld.Parent
    fromNameFolder = fromFld.Name
    With backupForm.sourceTxtBox
        .Text = fromParentFolder & "/" & fromNameFolder
        .Locked = True
    End With
End Sub

Sub getToFolder(toFld As MAPIFolder)
    ' gets the parent and name folder of the chosen destination folder
    toParentFolder = toFld.Parent
    toNameFolder = toFld.Name
    With backupForm.destinationTxtBox
        .Text = toParentFolder & "/" & toNameFolder
        .Locked = True
    End With
End Sub

Sub moveData(fromParentFolder As Variant, fromNameFolder As Variant, toParentFolder As Variant, toNameFolder As Variant, toPstParentFolder As Variant, toPstNameFolder As Variant)

    Dim sourceFolder As Outlook.MAPIFolder
    Dim targetFolder As Outlook.MAPIFolder
    Dim mailsCount As Long
    
    ' Source Folder
    Set sourceFolder = Outlook.Session.Folders(fromParentFolder).Folders(fromNameFolder)
    
    If (isPstCreated = False) Then
        ' Destination Folder
        Set targetFolder = Outlook.Session.Folders(toParentFolder).Folders(toNameFolder)
    Else
        Set targetFolder = Outlook.Session.Folders(toPstParentFolder).Folders(toPstNameFolder)
    End If
    
    mailsCount = sourceFolder.items.count
    
    If (mailsCount = 0) Then
        MsgBox "Source folder is empty!"
    Else
        ' move source data to the destination folder
        While mailsCount > 0
            sourceFolder.items.item(mailsCount).Move targetFolder
            mailsCount = mailsCount - 1
        Wend
        MsgBox "Backup done!"
    End If
End Sub

