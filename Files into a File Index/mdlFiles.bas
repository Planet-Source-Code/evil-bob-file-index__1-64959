Attribute VB_Name = "mdlFiles"

Global FI As New clsFileIndexer

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long 'ITEMIDLIST
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Public Function BrowseFolder(hwnd As Long, szDialogTitle As String) As String
    Dim X As Long, BI As BROWSEINFO, dwIList As Long, szPath As String, wPos As Integer
    
    On Local Error Resume Next
    
    BI.hOwner = hwnd
    BI.lpszTitle = szDialogTitle
    BI.ulFlags = BIF_RETURNONLYFSDIRS
    
    dwIList = SHBrowseForFolder(BI)
    szPath = Space$(512)
    X = SHGetPathFromIDList(ByVal dwIList, ByVal szPath)
    
    If X Then
        wPos = InStr(szPath, Chr(0))
        BrowseFolder = Trim(Left$(szPath, wPos - 1))
    Else
        BrowseFolder = vbNullString
    End If
End Function

'Simply grabs the names of the files in the folder (nothing fancy)
Public Sub GetFiles(FolderPath As String, ListBox As Object)
pathname = Dir(FolderPath & "\")

Do While pathname <> ""
    ListBox.AddItem FolderPath & "\" & pathname
    pathname = Dir
Loop
End Sub

'Function to just extract the filename without the path
Public Function FixFileFormat(FileName As String) As String
Dim X As Long, y As Long
Dim Curr As String
Dim Stack As String

X = Len(FileName)
y = 0

Do Until X = y
    Curr = Mid$(FileName, X, 1)
    If Curr = "\" Then
        GoTo TheEndFix:
    Else
        Stack = Stack & Curr
    End If
X = X - 1
Loop

TheEndFix:

FixFileFormat = ReverseString(Stack)
End Function

Public Function ReverseString(iString As String) As String
Dim X As Long, y As Long
Dim Curr As String
Dim Stack As String

X = Len(iString)
y = 0

Do Until X = y
    Curr = Mid$(iString, X, 1)
    Stack = Stack & Curr
X = X - 1
Loop

ReverseString = Stack
End Function
