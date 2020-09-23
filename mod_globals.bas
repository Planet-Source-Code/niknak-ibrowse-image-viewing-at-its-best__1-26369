Attribute VB_Name = "mod_globals"
Option Explicit

'********************************
'API DECLARATIONS
'********************************
    'USED TO KEEP FORM ONTOP
    Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
    'USED FOR FOLDER DIALOG
    Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
    Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
    Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
    Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

'***********************************
'PRIVATE CONSTAMTS
'***********************************
    Private Const BIF_RETURNONLYFSDIRS = 1
    Private Const MAX_PATH = 260

'***********************************
'PUBLIC CONSTAMTS
'***********************************
    Global Const HWND_TOPMOST = -1
    Global Const HWND_NOTOPMOST = -2
    Global Const SWP_NOSIZE = &H1
    Global Const SWP_NOMOVE = &H2
    Global Const SWP_NOACTIVATE = &H10
    Global Const SWP_SHOWWINDOW = &H40
    Global Const max_zoom_level = 4
    Global Const noof_supported_extensions = 7

'***********************************
'PRIVATE DATA TYPES
'***********************************
    Private Type extension
        description As String
        strdata As String
    End Type
    
    'USED FOR FOLDER DIALOG
    Private Type BrowseInfo
        hWndOwner As Long
        pIDLRoot As Long
        pszDisplayName As Long
        lpszTitle As Long
        ulFlags As Long
        lpfnCallback As Long
        lParam As Long
        iImage As Long
    End Type

'**********************************************************************
'PUBLIC DATA TYPES
'**********************************************************************
    'FUTURE ADDITIONS WILL INCLUDE VIEW TIME AND STRECH MODES ETC.
    Public Type pic
        filename As String
        width As Double
        height As Double
    End Type

'***********************************
'PUBLIC VARIABLES
'***********************************
    Public loaded As Boolean
    Public supported_extensions(noof_supported_extensions) As extension
    Public onlyfile As String

'***********************************
'PUBLIC SUBS
'***********************************
    'CONFIGURE ALL SUPPORTED FILE EXTENSIONS
    Public Sub configure_extensions()
        supported_extensions(1).description = "All Images"
        supported_extensions(1).strdata = ".bmp;*.emf;*.gif;*.ico;*.jpg;*.wmf"
        supported_extensions(2).description = "BMP Image"
        supported_extensions(2).strdata = ".bmp"
        supported_extensions(3).description = "Enhanced Metafile"
        supported_extensions(3).strdata = ".emf"
        supported_extensions(4).description = "GIF Image"
        supported_extensions(4).strdata = ".gif"
        supported_extensions(5).description = "Icon"
        supported_extensions(5).strdata = ".ico"
        supported_extensions(6).description = "JPG Image"
        supported_extensions(6).strdata = ".jpg"
        supported_extensions(7).description = "Windows Metafile"
        supported_extensions(7).strdata = ".wmf"
    End Sub
    
    'COUNT ALL VALID IMAGES IN A PATH
    Public Function count_images(spath As String) As Integer
        On Error Resume Next
        Dim fs, f, fc, f1
        Dim rndpos As Long
        Dim chkext As Integer
        Dim found As Integer
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set f = fs.GetFolder(spath)
        Set fc = f.Files
        frm_loading.Caption = "Counting supported images, please wait..."
        For Each f1 In fc
            DoEvents
            For chkext = 1 To noof_supported_extensions
                If InStr(1, f1.Name, supported_extensions(chkext).strdata, vbTextCompare) Then
                    onlyfile = f1.Name
                    found = found + 1
                End If
            Next chkext
        Next
        count_images = found
    End Function
    
    'GET FOLDER
    Public Function get_folder(hwnd As Long) As String
        Dim iNull As Integer, lpIDList As Long, lResult As Long
        Dim spath As String, udtBI As BrowseInfo
        With udtBI
            .hWndOwner = hwnd
            .lpszTitle = lstrcat("C:\", "")
            .ulFlags = BIF_RETURNONLYFSDIRS
        End With

        lpIDList = SHBrowseForFolder(udtBI)
        If lpIDList Then
            spath = String$(MAX_PATH, 0)
            SHGetPathFromIDList lpIDList, spath
            CoTaskMemFree lpIDList
            iNull = InStr(spath, vbNullChar)
            If iNull Then
                spath = Left$(spath, iNull - 1)
            End If
        End If
        get_folder = spath
    End Function
    
    'VERIFY FILE EXISTS
    Public Function verify_file(i_filename As String) As Boolean
        Dim fs
        Set fs = CreateObject("Scripting.FileSystemObject")
        If fs.fileexists(i_filename) Then
            verify_file = True
        Else
            verify_file = False
        End If
    End Function
    
    'SAVES A WINDOW POSITION
    Public Sub save_windpos(sform As Form)
        If sform.WindowState <> vbMinimized Then
            SaveSetting App.ProductName, "Windows", sform.Name & "-State", sform.WindowState
            SaveSetting App.ProductName, "Windows", sform.Name & "-Top", sform.Top
            SaveSetting App.ProductName, "Windows", sform.Name & "-Left", sform.Left
            SaveSetting App.ProductName, "Windows", sform.Name & "-Width", sform.width
            SaveSetting App.ProductName, "Windows", sform.Name & "-Height", sform.height
        End If
    End Sub
    
    'LOADS A WINDOWS SAVED POSITION
    Public Sub load_windpos(ByVal lform As Form)
        If Val(GetSetting(App.ProductName, "Windows", lform.Name & "-State")) = vbMaximized Then
            lform.WindowState = vbMaximized
        ElseIf GetSetting(App.ProductName, "Windows", lform.Name & "-State") <> "" Then
            lform.Top = Val(GetSetting(App.ProductName, "Windows", lform.Name & "-Top"))
            lform.Left = Val(GetSetting(App.ProductName, "Windows", lform.Name & "-Left"))
            lform.width = Val(GetSetting(App.ProductName, "Windows", lform.Name & "-Width"))
            lform.height = Val(GetSetting(App.ProductName, "Windows", lform.Name & "-Height"))
            If lform.Top > Screen.height Then lform.Top = 0
            If lform.Left > Screen.width Then lform.Left = 0
        End If
    End Sub
