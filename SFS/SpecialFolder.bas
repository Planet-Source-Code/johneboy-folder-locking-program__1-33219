Attribute VB_Name = "SpecialFolder"
Const CSIDL_DESKTOP = &H0
Const CSIDL_PROGRAMS = &H2
Const CSIDL_CONTROLS = &H3
Const CSIDL_PRINTERS = &H4
Const CSIDL_PERSONAL = &H5
Const CSIDL_FAVORITES = &H6
Const CSIDL_STARTUP = &H7
Const CSIDL_RECENT = &H8
Const CSIDL_SENDTO = &H9
Const CSIDL_BITBUCKET = &HA
Const CSIDL_STARTMENU = &HB
Const CSIDL_DESKTOPDIRECTORY = &H10
Const CSIDL_DRIVES = &H11
Const CSIDL_NETWORK = &H12
Const CSIDL_NETHOOD = &H13
Const CSIDL_FONTS = &H14
Const CSIDL_TEMPLATES = &H15
Const MAX_PATH = 260
Public Type SHITEMID
    cb As Long
    abID As Byte
End Type
Public Type ITEMIDLIST
    mkid As SHITEMID
End Type
Public Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Function GetSpecialfolder(CSIDL As Long) As String
    Dim r As Long
    Dim IDL As ITEMIDLIST
    r = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    If r = NOERROR Then
        Path$ = Space$(512)
        r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
        GetSpecialfolder = Left$(Path, InStr(Path, Chr$(0)) - 1)
        Exit Function
    End If
    GetSpecialfolder = ""
End Function


Public Function CheckIfSystemFolder(FolderPath As Label, CheckIt As Label)
Dim WinPath As String, strSave As String
    strSave = String(200, Chr$(0))
    WinPath = Left$(strSave, GetWindowsDirectory(strSave, Len(strSave))) + ""
    ProgramPath = "C:\Program Files"
    
If FolderPath = GetSpecialfolder(CSIDL_STARTMENU) Or FolderPath = GetSpecialfolder(CSIDL_FAVORITES) Or FolderPath = GetSpecialfolder(CSIDL_PROGRAMS) Or FolderPath = GetSpecialfolder(CSIDL_DESKTOP) _
Or FolderPath = GetSpecialfolder(CSIDL_CONTROLS) Or FolderPath = GetSpecialfolder(CSIDL_PRINTERS) Or FolderPath = GetSpecialfolder(CSIDL_PERSONAL) Or FolderPath = GetSpecialfolder(CSIDL_STARTUP) _
Or FolderPath = GetSpecialfolder(CSIDL_RECENT) Or FolderPath = GetSpecialfolder(CSIDL_SENDTO) Or FolderPath = GetSpecialfolder(CSIDL_BITBUCKET) Or FolderPath = GetSpecialfolder(CSIDL_DESKTOPDIRECTORY) _
Or FolderPath = GetSpecialfolder(CSIDL_DRIVES) Or FolderPath = GetSpecialfolder(CSIDL_NETWORK) Or FolderPath = GetSpecialfolder(CSIDL_NETHOOD) Or FolderPath = GetSpecialfolder(CSIDL_FONTS) _
Or FolderPath = GetSpecialfolder(CSIDL_TEMPLATES) Or FolderPath = WinPath Or FolderPath = ProgramPath Then
CheckIt.Caption = "Yes"
Else
CheckIt.Caption = "No"
End If
End Function
