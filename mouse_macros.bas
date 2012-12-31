Attribute VB_Name = "mouse_macros"
Option Explicit

'// The store, update to match yours!
Const MAIL_STORE As String = "Mailbox - Spotswood, Brent"

'// Moves all selected items in the active window to the specified Folder
Private Sub MoveSelectionToFolder(f As Folder)
    Dim win As Outlook.Explorer
    Dim selObjs() As Object
    Dim l As Long
    
    Set win = Outlook.Application.ActiveWindow
    
    '// Make an array of the currently selected items
    ReDim selObjs(1 To win.Selection.Count)
    For l = 1 To win.Selection.Count
        Set selObjs(l) = win.Selection(l)
    Next l
    
    '// Move each item. Debug output so you can see a list of what has moved and where.
    For l = LBound(selObjs) To UBound(selObjs)
        Debug.Print CStr(Now()) & " Moving item: " & selObjs(l).EntryID
        Debug.Print CStr(Now()) & "          to: " & f.FolderPath
        Call selObjs(l).Move(f)
    Next l
End Sub


'// Move all selected items to a subfolder in "Active Projects" by matching the number
Public Sub MoveSelectionToActiveProjFolder(folderNum As Integer)
    Dim activeFolders As Folders
    Dim f As Folder
    Dim l As Long
    Dim numStr As String
    
    '// The parent folder for this (customize for how you setup your own folders)
    Set activeFolders = Application.Session.Stores(MAIL_STORE).GetRootFolder().Folders("@").Folders("Projects").Folders("Actively Working").Folders
    
    '// Iterrate each subfolder, compare the left 3 characters for a match.
    numStr = Format(folderNum, "00") & "."
    For l = 1 To activeFolders.Count
        Set f = activeFolders(l)
        If (Left(f.Name, 3) = numStr) Then
            Call MoveSelectionToFolder(f)
            Exit Sub
        End If
    Next l
    
    '// If we didn't exit then we didn't match a folder. Send a notice.
    Call MsgBox("No folder is assigned to number " & CStr(folderNum) & ". To assign, rename a folder in the ""Actively Working"" folder to begin with """ & CStr(folderNum) & ". """, vbInformation, "Project Not Assigned To Number")
End Sub

'// Move all selected folders to the "Misc" folder.
'// This is a public macro that can be placed on an Outlook toolbar
Public Sub MoveSelectionToMisc()
    On Error Resume Next
    
    '// Customize the folder path for your needs
    Call MoveSelectionToFolder(Application.Session.Stores(MAIL_STORE).GetRootFolder().Folders("@").Folders("Misc"))
End Sub

'// Macros that can be placed on a toolbar and invoked to perform moves to folders
Public Sub MoveSelectionToActiveProject4(): Call MoveSelectionToActiveProjFolder(4): End Sub
Public Sub MoveSelectionToActiveProject5(): Call MoveSelectionToActiveProjFolder(5): End Sub
Public Sub MoveSelectionToActiveProject6(): Call MoveSelectionToActiveProjFolder(6): End Sub
Public Sub MoveSelectionToActiveProject7(): Call MoveSelectionToActiveProjFolder(7): End Sub
Public Sub MoveSelectionToActiveProject8(): Call MoveSelectionToActiveProjFolder(8): End Sub
Public Sub MoveSelectionToActiveProject9(): Call MoveSelectionToActiveProjFolder(9): End Sub
Public Sub MoveSelectionToActiveProject10(): Call MoveSelectionToActiveProjFolder(10): End Sub
Public Sub MoveSelectionToActiveProject11(): Call MoveSelectionToActiveProjFolder(11): End Sub







