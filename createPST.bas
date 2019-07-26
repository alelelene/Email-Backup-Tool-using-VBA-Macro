Attribute VB_Name = "createPST"
Option Explicit
Public strDisplayName As String
Public toPstParentFolder As Variant
Public toPstNameFolder As Variant
Public isPstCreated As Boolean

Sub createPSTFile()
    'variables
    Dim objOutlook, objNameSpace, oShell
    Dim strUserProfile, WSHNetwork, strUser, strPSTPath
    Dim intAnswer
    
    Dim objStore As Outlook.Store
    Dim objFolder As Outlook.folder
    
    Dim sFolder As String
    Dim xlObj As Excel.Application
        
    'Grab the user name
    Set WSHNetwork = CreateObject("WScript.Network")
    strUser = WSHNetwork.UserName
        
    'grab user profile
    Set oShell = CreateObject("Wscript.Shell")
    strUserProfile = oShell.ExpandEnvironmentStrings("%USERPROFILE%")
        
    'open the select folder prompt
    Set xlObj = New Excel.Application
    With xlObj.FileDialog(msoFileDialogFolderPicker)
        .title = "Select a Folder"
        .ButtonName = "Create PST here"
        .InitialFileName = strUserProfile & "\"
        If .Show = -1 Then ' if OK is pressed
            sFolder = .SelectedItems(1)
        End If
    End With
    xlObj.Quit
    Set xlObj = Nothing
        
    If sFolder <> "" Then ' if a file was chosen
    
        strPSTPath = sFolder & "\" & strUser & ShortDate(Now) & ".pst"

        'hook into MAPI and create pst
        Set objOutlook = CreateObject("Outlook.Application")
        Set objNameSpace = objOutlook.GetNamespace("MAPI")
        objNameSpace.AddStoreEx strPSTPath, 2
        
            
        strDisplayName = strUser & ShortDate(Now)
        toPstParentFolder = strDisplayName
            
        Set objOutlook = CreateObject("Outlook.Application")
        Set objNameSpace = objOutlook.GetNamespace("MAPI")
            
        For Each objStore In objNameSpace.Stores
            If objStore.FilePath = strPSTPath Then
                Set objFolder = objStore.GetRootFolder()
                objFolder.Name = strDisplayName
            End If
        Next
            
        Call createFolder
        
        'clean things up
        Set objNameSpace = Nothing
        Set objOutlook = Nothing
    End If
    
    With backupForm.destinationTxtBox
        .Text = toPstParentFolder & "/" & toPstNameFolder
        .Locked = True
    End With
    
    isPstCreated = True
    
End Sub

'function to yield date in yyyyddmm_hhmmss format
Function ShortDate(inDate)
    'convert date to yyyymmdd_hhmmss format for ease of sorting
    Dim szSecond
    Dim szMinute
    Dim szHour
    Dim szMonth
    Dim szDay
    Dim szYear
    
    szSecond = Second(inDate)
    If Len(CStr(szSecond)) = 1 Then
        szSecond = "0" & szSecond
    End If
    szMinute = Minute(inDate)
    If Len(CStr(szMinute)) = 1 Then
        szMinute = "0" & szMinute
    End If
    szHour = Hour(inDate)
    If Len(CStr(szHour)) = 1 Then
        szHour = "0" & szHour
    End If
    szMonth = Month(inDate)
    If Len(CStr(szMonth)) = 1 Then
        szMonth = "0" & szMonth
    End If
    szDay = Day(inDate)
    If Len(CStr(szDay)) = 1 Then
        szDay = "0" & szDay
    End If
    szYear = Year(inDate)
    ShortDate = szYear & szMonth & szDay & "_" & szHour & szMinute & szSecond
End Function

