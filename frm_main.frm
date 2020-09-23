VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm frm_main 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Ibrowse"
   ClientHeight    =   6195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6465
   Icon            =   "frm_main.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ils_images 
      Left            =   60
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":08CA
            Key             =   "open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":0E64
            Key             =   "slideshow"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":13FE
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_main.frx":1998
            Key             =   "thumbnails"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb_toolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   635
      ButtonWidth     =   2143
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ils_images"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            Key             =   "open"
            Object.ToolTipText     =   "Open an image"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Slideshow"
            Key             =   "slideshow"
            Object.ToolTipText     =   "Start a slideshow presentation"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Thumbnails"
            Key             =   "thumbnails"
            Object.ToolTipText     =   "Show thumbnails for a folder"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "exit"
            Object.ToolTipText     =   "Exit Ibrowse"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.Menu men_file 
      Caption         =   "File"
      Begin VB.Menu men_file_open 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu men_file_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu men_file_slideshow 
         Caption         =   "Slideshow"
         Shortcut        =   ^S
      End
      Begin VB.Menu men_file_thumbs 
         Caption         =   "Thumbnails"
         Shortcut        =   ^T
      End
      Begin VB.Menu men_file_sep2 
         Caption         =   "-"
      End
      Begin VB.Menu men_file_exit 
         Caption         =   "Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu men_win 
      Caption         =   "Windows"
      WindowList      =   -1  'True
      Begin VB.Menu men_win_cascade 
         Caption         =   "Cascade"
         Shortcut        =   ^C
      End
      Begin VB.Menu men_win_tileh 
         Caption         =   "Tile Horizontally"
         Shortcut        =   ^H
      End
      Begin VB.Menu men_win_tilev 
         Caption         =   "Tile Vertically"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu men_help 
      Caption         =   "Help"
      Begin VB.Menu men_help_about 
         Caption         =   "About Eyebrowse"
      End
      Begin VB.Menu men_help_bug 
         Caption         =   "Report a bug"
      End
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***********************************
'API DECLARATIONS
'***********************************
    'USED FOR COMMON DIALOG
    Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
  
'***********************************
'PRIVATE CONSTANTS
'***********************************
    Private Const SW_SHOWNORMAL = 1

'***********************************
'PRIVATE DATA TYPES
'***********************************
    'USED FOR COMMON DIALOG
    Private Type OPENFILENAME
        lStructSize As Long
        hWndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
    End Type
    
'***********************************
'FORM EVENTS
'***********************************
    'LOAD
    Private Sub MDIForm_Load()
        '---------------------------
        'RESTRICTS FORM SIZE
        startupwidth = Me.width \ Screen.TwipsPerPixelX
        startupheight = Me.height \ Screen.TwipsPerPixelY
        minX = startupwidth
        minY = startupheight
        maxX = Screen.width \ Screen.TwipsPerPixelX
        maxY = Screen.height \ Screen.TwipsPerPixelY
        Call SubClass(Me.hwnd)
        '---------------------------
        configure_extensions
        load_windpos Me
    End Sub

    'UNLOAD
    Private Sub MDIForm_Unload(Cancel As Integer)
        '---------------------------
        'RESTRICTS FORM SIZE
        UnSubClass Me.hwnd
        '---------------------------
        save_windpos Me
    End Sub

'***********************************
'MENU SUBS
'***********************************
    'OPEN IMAGE
    Private Sub men_file_open_Click()
        Dim com_dialog As OPENFILENAME
        Dim chkext As Integer
        With com_dialog
            .lStructSize = Len(com_dialog)
            .hWndOwner = Me.hwnd
            .hInstance = App.hInstance
            For chkext = 1 To noof_supported_extensions
                .lpstrFilter = .lpstrFilter + supported_extensions(chkext).description & _
                               " (*" & supported_extensions(chkext).strdata & ")" + Chr$(0) + _
                               "*" & supported_extensions(chkext).strdata + Chr$(0)
            Next chkext
            .lpstrFile = Space$(254)
            .nMaxFile = 255
            .lpstrFileTitle = Space$(254)
            .nMaxFileTitle = 255
            .lpstrInitialDir = App.Path
            .lpstrTitle = "Ibrowse - Open"
            .flags = 0
        End With
        If GetOpenFileName(com_dialog) Then
            file_open com_dialog.lpstrFile
        End If
    End Sub
       
    'SLIDESHOW
    Private Sub men_file_slideshow_Click()
        Dim spath As String
        spath = get_folder(Me.hwnd)
        If spath <> "" Then start_slideshow spath & "\"
    End Sub

    'THUMBNAILS
    Private Sub men_file_thumbs_Click()
        Dim spath As String
        spath = get_folder(Me.hwnd)
        If spath <> "" Then show_thumbs spath & "\"
    End Sub
    
    'EXIT IBROWSE
    Private Sub men_file_exit_Click()
        Dim retval As Long
        retval = MsgBox("Are you sure?", vbYesNo, "Quit Ibrowse")
        If retval = vbYes Then
            Unload Me
        End If
    End Sub
    
    'WINDOW ARRANGEMENTS
    Private Sub men_win_cascade_Click()
        Me.Arrange vbCascade
    End Sub
    Private Sub men_win_tileh_Click()
        Me.Arrange vbTileHorizontal
    End Sub
    Private Sub men_win_tilev_Click()
        Me.Arrange vbTileVertical
    End Sub

    'ABOUT IBROWSE
    Private Sub men_help_about_Click()
        Load frm_about
        frm_about.Show
    End Sub
    
    'REPORT A BUG
    Private Sub men_help_bug_Click()
        ShellExecute Me.hwnd, vbNullString, "mailto:np24@blueyonder.co.uk", vbNullString, "", SW_SHOWNORMAL
    End Sub
    
'***********************************
'TOOLBAR EVENTS
'***********************************
    Private Sub tlb_toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
        Select Case Button.Key
            Case "open"
                men_file_open_Click
            Case "slideshow"
                men_file_slideshow_Click
            Case "thumbnails"
                men_file_thumbs_Click
            Case "exit"
                men_file_exit_Click
        End Select
    End Sub

'***********************************
'IMAGE OPERATIONS
'***********************************
    'OPEN IMAGE
    Public Sub file_open(lpstrFile As String)
        Dim new_image As frm_image
        Set new_image = New frm_image
        With new_image
            .Caption = lpstrFile
            .load_image lpstrFile
            .centre_image
            .WindowState = vbMaximized
            .Show
            If .Tag = "error" Then Unload new_image
        End With
    End Sub
    
    'START SLIDESHOW
    Public Sub start_slideshow(spath As String)
        Dim new_slideshow As frm_slideshow
        Dim noofpics As Integer
        Set new_slideshow = New frm_slideshow
        Load frm_loading
        frm_loading.Show
        With new_slideshow
            noofpics = .scan_folder(spath)
            If noofpics > 1 Then
                .Caption = spath
                .nextpic
                .centre_image
                .WindowState = vbMaximized
                .Show
                .start_slideshow True
            End If
        End With
        Unload frm_loading
    End Sub
    
    'THUMBNAILS
    Private Sub show_thumbs(spath As String)
        Dim new_thumbs As frm_thumbs
        Dim noofpics As Integer
        Set new_thumbs = New frm_thumbs
        Load frm_loading
        frm_loading.Show
        With new_thumbs
            noofpics = .scan_folder(spath)
            If noofpics > 1 Then
                .Caption = spath & " - " & noofpics & " Thumbnails"
                .display_thumbs
                .Show
            End If
        End With
        Unload frm_loading
    End Sub

