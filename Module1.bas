Attribute VB_Name = "Module1"
'This module contains all the declarations to use the
'Windows 95 Shell API to use the browse for folders
'dialog box.  To use the browse for folders dialog box,
'please call the BrowseForFolders function using the
'syntax: stringFolderPath=BrowseForFolders(Hwnd,TitleOfDialog)
'
'For contacting information, see other module

Option Explicit

Public Type BrowseInfo      'receive information about the folder selected by user.
     
     hWndOwner As Long      'Handle to th owner window for the dialog box.
     pIDLRoot As Long       'Pointer to an itemlist structure.
     pszDisplayName As Long 'Add. of buffer -receive the display name of folder selected.
     lpszTitle As Long      'Display above the tree view control.
     ulFlags As Long        'specifying the options for the dialog box (notify event).
     lpfnCallback As Long   'Add. of application-defined funtion.
     lParam As Long         'Application-defined value pass to callback function.
                            '(receives messages from the operating system.)
     iImage As Long         'Receive image associated with the selected folder.

End Type

Public Const BIF_RETURNONLYFSDIRS = 1
Public Const MAX_PATH = 260

Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Function BrowseForFolder(hWndOwner As Long, sPrompt As String) As String
'Creates a dialog box (select a folder and returns the selected folder's Folder object).

    'declare variables to be used
     Dim iNull As Integer
     Dim lpIDList As Long
     Dim lResult As Long
     Dim sPath As String
     Dim udtBI As BrowseInfo

    'initialise variables
     With udtBI
        .hWndOwner = hWndOwner
        .lpszTitle = lstrcat(sPrompt, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
     End With

    'Call the browse for folder API
     lpIDList = SHBrowseForFolder(udtBI)
     
    'get the resulting string path
     If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        iNull = InStr(sPath, vbNullChar)
        If iNull Then sPath = Left$(sPath, iNull - 1)
     End If

    'If cancel was pressed, sPath = ""
     BrowseForFolder = sPath

End Function

Function cPath() As String
If Right(App.Path, 1) <> "\" Then
   cPath = App.Path & "\"
End If
End Function

Function PathExists(ByVal strPathName As String) As Boolean
On Error GoTo errHandle

If Dir(strPathName, vbDirectory) <> "" Then
   PathExists = True
Else
   PathExists = False
End If

Exit Function
errHandle:
PathExists = False
End Function


Function FileExists(ByVal strPathName As String) As Boolean
On Error GoTo errHandle
    
    Open strPathName For Input As #1
    Close #1
    FileExists = True
    
Exit Function
errHandle:
FileExists = False
End Function

Function Get_WinPath() As String
Dim rtn
Dim Buffer As String 'declare the needed variables
   
   Buffer = Space(MAX_PATH)
   rtn = GetWindowsDirectory(Buffer, Len(Buffer))  'get the path
   Get_WinPath = Left(Buffer, rtn) 'parse the path to the global string
   If Right(Get_WinPath, 1) <> "\" Then
      Get_WinPath = Get_WinPath & "\"
   End If
End Function
