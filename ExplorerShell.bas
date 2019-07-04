Attribute VB_Name = "Explorer"
'****************************************************************
' Exemple issu de vbnet.mvps.org/code/browse/browseadv.htm
'****************************************************************
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2004 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public fso As Scripting.FileSystemObject

'parameters for SHBrowseForFolder
Public Type BROWSEINFO    'BI
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Public Type DLLVERSIONINFO
  cbSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformID As Long
End Type

Public Const VER_PLATFORM_WIN32s = 0          'Win32s on Windows 3.1.
Public Const VER_PLATFORM_WIN32_WINDOWS = 1   'Win32 on Windows 95,98, Me.
Public Const VER_PLATFORM_WIN32_NT = 2        'NT or 2000 - XP - Server 2003

'For Windows 95,dwMajorVersion=  4, dwMinorVersion = zero.
'For Windows 98 dwMajorVersion = 4, dwMinorVersion = 10
'For Windows Me dwMajorVersion = 4, dwMinorVersion = 90

'For Windows NT dwMajorVersion =3 or 4, dwMinorVersion = 51 or 0

'For Windows 2000 dwMajorVersion = 5, dwMinorVersion = 0
'For Windows XP dwMajorVersion = 5, dwMinorVersion = 1
'For Windows Server 2003 dwMajorVersion = 5, dwMinorVersion = 2

'windows-defined type OSVERSIONINFO
Public Type OSVERSIONINFO
  dwOSVersionSize As Long
  dwMajorVersion  As Long
  dwMinorVersion  As Long
  dwBuildNumber   As Long
  dwPlatformID    As Long
  szCSDVersion    As String * 128
End Type
                                  
Public Enum VersionEnum
  W_95
  W_98
  W_Me
  W_NT
  W_2000
  W_XP
  W_2003
End Enum

Public Declare Function DllGetVersion Lib "shlwapi" (lpVersionInformation As DLLVERSIONINFO) As Long

Public Declare Function GetVersionEx Lib "KERNEL32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

'Converts an item identifier list to a file system path.
Public Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

'Displays a dialog box that enables the user to select a shell folder.
Public Declare Function SHBrowseForFolder Lib "shell32" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

Public Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
    
Public Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)

Public Declare Sub CoInitializeEx Lib "ole32" (ByVal lpString As Any, ByVal dwCoInit As Long)

Public Declare Function LoadLibrary Lib "KERNEL32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Public Declare Function GetLastError Lib "KERNEL32" () As Long

Public Declare Function GetProcAddress Lib "KERNEL32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'Gestion de l'allocation mémoire
Public Declare Function LocalAlloc Lib "KERNEL32" (ByVal wFlags As Long, ByVal wBytes As Long) As Long
Public Declare Function LocalFree Lib "KERNEL32" (ByVal hMem As Long) As Long
Public Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'Chaines de caractères
Public Declare Function lstrlen Lib "KERNEL32" Alias "lstrlenA" (ByVal lpString As String) As Long

Public Const LMEM_FIXED = &H0
Public Const LMEM_ZEROINIT = &H40
Public Const LPTR = (LMEM_FIXED + LMEM_ZEROINIT)

'BROWSEINFO.ulFlags values:
Public Const BIF_RETURNONLYFSDIRS = &H1      'Only file system directories
Public Const BIF_DONTGOBELOWDOMAIN = &H2     'No network folders below domain level
Public Const BIF_STATUSTEXT = &H4            'Includes status area in the dialog (for callback)
Public Const BIF_RETURNFSANCESTORS = &H8     'Only returns file system ancestors
Public Const BIF_EDITBOX = &H10              'Allows user to rename selection
Public Const BIF_VALIDATE = &H20             'Insist on valid edit box result (or CANCEL)
Public Const BIF_NEWDIALOGSTYLE = &H40       'Version 5.0. Use the new user-interface.
Public Const BIF_BROWSEINCLUDEURLS = &H80    'Setting this flag provides the user with
Public Const BIF_USENEWUI = 192              'a larger dialog box that can be resized.
' ( BIF_EDITBOX | BIF_NEWDIALOGSTYLE )       'It has several new capabilities including:
                                             'dialog box, reordering, context menus, new
                                             'folders, drag and drop capability within
                                             'the delete, and other context menu commands.
                                             'To use you must call OleInitialize or
                                             'CoInitialize before calling SHBrowseForFolder.
Public Const BIF_NONEWFOLDERBUTTON = &H200   'Version 6.0
Public Const BIF_NOTRANSLATETARGETS = &H400
Public Const BIF_BROWSEFORCOMPUTER = &H1000  'Only returns computers.
Public Const BIF_BROWSEFORPRINTER = &H2000   'Only returns printers.
Public Const BIF_BROWSEINCLUDEFILES = &H4000 'Browse for everything

Public Const MAX_PATH = 260

Public Const CSIDL_DESKTOP = &H0                   '(desktop)
Public Const CSIDL_INTERNET = &H1                  'Internet Explorer (icon on desktop)
Public Const CSIDL_PROGRAMS = &H2                  'Start Menu\Programs
Public Const CSIDL_CONTROLS = &H3                  'My Computer\Control Panel
Public Const CSIDL_PRINTERS = &H4                  'My Computer\Printers
Public Const CSIDL_PERSONAL = &H5                  'My Documents
Public Const CSIDL_FAVORITES = &H6                 '(user)\Favourites
Public Const CSIDL_STARTUP = &H7                   'Start Menu\Programs\Startup
Public Const CSIDL_RECENT = &H8                    '(user)\Recent
Public Const CSIDL_SENDTO = &H9                    '(user)\SendTo
Public Const CSIDL_BITBUCKET = &HA                 '(desktop)\Recycle Bin
Public Const CSIDL_STARTMENU = &HB                 '(user)\Start Menu
Public Const CSIDL_DESKTOPDIRECTORY = &H10         '(user)\Desktop
Public Const CSIDL_DRIVES = &H11                   'My Computer
Public Const CSIDL_NETWORK = &H12                  'Network Neighbourhood
Public Const CSIDL_NETHOOD = &H13                  '(user)\nethood
Public Const CSIDL_FONTS = &H14                    'windows\fonts
Public Const CSIDL_TEMPLATES = &H15
Public Const CSIDL_COMMON_STARTMENU = &H16         '(all users)\Start Menu
Public Const CSIDL_COMMON_PROGRAMS = &H17          '(all users)\Programs
Public Const CSIDL_COMMON_STARTUP = &H18           '(all users)\Startup
Public Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19  '(all users)\Desktop
Public Const CSIDL_APPDATA = &H1A                  '(user)\Application Data
Public Const CSIDL_PRINTHOOD = &H1B                '(user)\PrintHood
Public Const CSIDL_LOCAL_APPDATA = &H1C            '(user)\Local Settings
                                                   '\Application Data (non roaming)
Public Const CSIDL_ALTSTARTUP = &H1D               'non localized startup
Public Const CSIDL_COMMON_ALTSTARTUP = &H1E        'non localized common startup
Public Const CSIDL_COMMON_FAVORITES = &H1F
Public Const CSIDL_INTERNET_CACHE = &H20
Public Const CSIDL_COOKIES = &H21
Public Const CSIDL_HISTORY = &H22
Public Const CSIDL_COMMON_APPDATA = &H23           '(all users)\Application Data
Public Const CSIDL_WINDOWS = &H24                  'GetWindowsDirectory()
Public Const CSIDL_SYSTEM = &H25                   'GetSystemDirectory()
Public Const CSIDL_PROGRAM_FILES = &H26            'C:\Program Files
Public Const CSIDL_MYPICTURES = &H27               'C:\Program Files\My Pictures
Public Const CSIDL_PROFILE = &H28                  'USERPROFILE
Public Const CSIDL_PROGRAM_FILES_COMMON = &H2B     'C:\Program Files\Common
Public Const CSIDL_COMMON_TEMPLATES = &H2D         '(all users)\Templates
Public Const CSIDL_COMMON_DOCUMENTS = &H2E         '(all users)\Documents
Public Const CSIDL_COMMON_ADMINTOOLS = &H2F        '(all users)\Start Menu\Programs
                                                   '\Administrative Tools
Public Const CSIDL_ADMINTOOLS = &H30               '(user)\Start Menu\Programs
                                                   '\Administrative Tools
Public Const CSIDL_FLAG_CREATE = &H8000&           'combine with CSIDL_ value to force
                                                   'create on SHGetSpecialFolderLocation()
Public Const CSIDL_FLAG_DONT_VERIFY = &H4000       'combine with CSIDL_ value to force
                                                   'create on SHGetSpecialFolderLocation()
Public Const CSIDL_FLAG_MASK = &HFF00              'mask for all possible flag values
'not used
'Public Const CSIDL_SYSTEMX86 = &H29               'x86 system directory on RISC
'Public Const CSIDL_PROGRAM_FILESX86 = &H2A        'x86 C:\Program Files on RISC
'Public Const CSIDL_PROGRAM_FILES_COMMONX86 = &H2C 'x86 Program Files\Common on RISC

'Constante issue de Winuser.h
Public Const WM_USER = &H400

'Constantes issues de shlobj.h
' message from browser
Public Const BFFM_INITIALIZED = 1
Public Const BFFM_SELCHANGED = 2
Public Const BFFM_VALIDATEFAILEDA = 3     ' lParam:szPath ret:1(cont),0(EndDialog)
Public Const BFFM_VALIDATEFAILEDW = 4     ' lParam:wzPath ret:1(cont),0(EndDialog)

' messages to browser
Public Const BFFM_SETSTATUSTEXTA = (WM_USER + 100)
Public Const BFFM_ENABLEOK = (WM_USER + 101)
Public Const BFFM_SETSELECTIONA = (WM_USER + 102)
Public Const BFFM_SETSELECTIONW = (WM_USER + 103)
Public Const BFFM_SETSTATUSTEXTW = (WM_USER + 104)

Private unChemin As String

Public Function VersionPoste() As VersionEnum
Dim OSV As OSVERSIONINFO

  OSV.dwOSVersionSize = Len(OSV)

  If GetVersionEx(OSV) <> 0 Then

    With OSV

      Select Case .dwPlatformID
    'dwPlatformID contains a value representing the OS
      Case VER_PLATFORM_WIN32_WINDOWS
        Select Case .dwMajorVersion
        Case 4  ' 95, 98, Millenium
          Select Case .dwMinorVersion
          Case 0
            VersionPoste = W_95
          Case 10
            VersionPoste = W_98
          Case 90
            VersionPoste = W_Me   ' Millenium
          End Select
        End Select
        
      Case VER_PLATFORM_WIN32_NT
        Select Case .dwMajorVersion
        Case 3, 4   ' 3.51 ou 4.0
          'msg = msg & " : Version " & .dwMajorVersion & "." & .dwMinorVersion
          VersionPoste = W_NT
        Case 5  ' 2000 o uServer 2003
          Select Case .dwMinorVersion
          Case 0
            VersionPoste = W_2000
          Case 1
            VersionPoste = W_XP
          Case 2
            VersionPoste = W_2003 ' Server 2003
          End Select
        End Select
      End Select
    End With
  End If

End Function

Public Function IsWinNT() As Boolean
'Plateforme NT : NT 3.51,NT 4.0, 2000, XP, Server 2003
Dim OSV As OSVERSIONINFO

  OSV.dwOSVersionSize = Len(OSV)

  If GetVersionEx(OSV) <> 0 Then

    'dwPlatformID contains a value representing
    'the OS; if VER_PLATFORM_WIN32_NT,
    'return true
     IsWinNT = (OSV.dwPlatformID = VER_PLATFORM_WIN32_NT)
  End If

End Function

Public Function IsWin2000() As Boolean
Dim OSV As OSVERSIONINFO
   
  IsWin2000 = (VersionPoste = W_2000)

End Function

Public Function Browse(ByVal BIF_FLAGS As Long, ByVal sTitre As String, ByVal Feuille As VB.Form, Optional ByVal Chemin As String) As String
Dim s As String
Static Initialized As Boolean

Dim pidl As Long
Dim BI As BROWSEINFO
Dim lpSelPath As Long
Dim Chaine As String

  If Not Initialized Then
    s = String(1, Chr(0))
    CoInitializeEx s, 0
    Initialized = True
  End If
  
   'Fill BROWSEINFO structure data
  
  If Feuille.untxtChemin Is Nothing Then
    Chaine = Chemin
  Else
    Chaine = Feuille.untxtChemin
  End If
  
  Chaine = PremierParentExistant(Chaine)

  With BI
    .hOwner = Feuille.hwnd
    .pidlRoot = 0 'CSIDL
    .lpszTitle = sTitre
    .ulFlags = BIF_FLAGS
    .pszDisplayName = Space(MAX_PATH)
    '

    If Len(Chaine) > 0 Then
      .lpfn = FARPROC(AddressOf BrowseCallbackProc)
      lpSelPath = LocalAlloc(LPTR, Len(Chaine) + 1)
      'Byval est nécessaire car le paramètre est Byref
      CopyMemory ByVal lpSelPath, ByVal Chaine, Len(Chaine) + 1
      .lParam = lpSelPath
    End If
  End With

  'show dialog returning pidl to selected item
   pidl = SHBrowseForFolder(BI)

    Call LocalFree(lpSelPath)
    
    
  If pidl <> 0 Then
  'if pidl is valid, parse & return the user's selection
    Browse = PathFromPidl(pidl)

    'pszDisplayName contains the string
    'representing the users last selection.
    'Even when SHGetPathFromIDList is empty,
    'this should return the selection, making
    'it the choice for obtaining user information
    'when selecting Printers, Control Panel etc,
    'or any of the other virtual folders that
    'does not normally return a path.
    
    If Len(Browse) > 0 Then
      With Feuille
        If Not .untxtChemin Is Nothing Then
          .untxtChemin.Text = Browse
        End If
        If Not .untxtDossier Is Nothing Then
          .untxtDossier.Text = ChaineC(BI.pszDisplayName)
        End If
      End With
    End If
  End If

  'free the pidl
  Call CoTaskMemFree(pidl)
  
End Function

Private Function PremierParentExistant(ByVal Chemin As String) As String

  If Len(Chemin) > 0 Then
    If fso.FolderExists(Chemin) Then
      PremierParentExistant = Chemin
    Else
      PremierParentExistant = PremierParentExistant(fso.GetParentFolderName(Chemin))
    End If
  End If
  
End Function

Private Function FARPROC(ByVal lpfn As Long) As Long
  FARPROC = lpfn
End Function

Public Function BrowseCallbackProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
Dim sPath As String
Dim lR As Long

  Select Case uMsg
  Case BFFM_INITIALIZED
    'Byval est nécessaire car le paramètre est Byref
'    SendMessage hWnd, IIf(IsWinNT, BFFM_SETSELECTION, BFFM_SETSELECTIONA), True, ByVal lpData
    SendMessage hwnd, BFFM_SETSELECTIONA, True, ByVal lpData
    
  
  ' Selection has changed (lParam contains pidl of selected folder)
  Case BFFM_SELCHANGED
     ' Display full path if status area if enabled
     sPath = PathFromPidl(lParam)
     lR = SendMessage(hwnd, BFFM_SETSTATUSTEXTA, False, sPath)
  End Select
End Function

Public Function PathFromPidl(ByVal pidl As Long) As String
Dim sPath As String
Dim lR As Long
   sPath = String(MAX_PATH, 0)
      'SHGetPathFromIDList returns the absolute
      'path to the selected item. No path is returned for virtual folders.
   lR = SHGetPathFromIDList(pidl, sPath)
   If lR <> 0 Then
      PathFromPidl = ChaineC(sPath)
   End If
End Function

Public Function ChaineC(ByVal Chaine As String) As String
  ChaineC = Left(Chaine, lstrlen(Chaine))
End Function
