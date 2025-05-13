Attribute VB_Name = "LoopThroughFilesInFolder"
Option Explicit

Sub LoopThroughAllFiles(folderPath As String, Optional InclSubFolder As Boolean = True)
    Dim fileName As String
    Dim fso As Object
    Dim folder As Object
    Dim subfolder As Object
    Dim file As Object
    Dim app_path_sep As String
    
    app_path_sep = Application.PathSeparator
    If Right(folderPath, 1) <> app_path_sep Then folderPath = folderPath & app_path_sep

    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    ' Call recursive procedure
    ProcessFolder folder

    MsgBox "Done!"

End Sub

Sub ProcessFolder(ByVal folder As Object, Optional InclSubFolder As Boolean = True)
    Dim subfolder As Object
    Dim file As Object

    ' Loop through each file in the current folder
    For Each file In folder.Files
        Debug.Print file.Path ' You can replace this line with your own processing
    Next file

    ' Recursively loop through each subfolder
    If Not InclSubFolder Then Exit Sub
    For Each subfolder In folder.SubFolders
        ProcessFolder subfolder
    Next subfolder
    
    
End Sub

