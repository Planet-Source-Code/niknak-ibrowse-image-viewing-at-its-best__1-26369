VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_slideshow 
   BackColor       =   &H80000001&
   Caption         =   "<folder name>"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8340
   Icon            =   "frm_slideshow.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   8340
   Begin MSComctlLib.ImageList ils_images 
      Left            =   60
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_slideshow.frx":030A
            Key             =   "stop"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_slideshow.frx":08A4
            Key             =   "decrease"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_slideshow.frx":0E3E
            Key             =   "increase"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_slideshow.frx":13D8
            Key             =   "fullscreen"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_slideshow.frx":1972
            Key             =   "interval"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_slideshow.frx":1F0C
            Key             =   "forward"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_slideshow.frx":24A6
            Key             =   "start"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_slideshow.frx":2A40
            Key             =   "back"
         EndProperty
      EndProperty
   End
   Begin VB.Timer tim_nextpic 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   120
      Top             =   720
   End
   Begin MSComctlLib.Toolbar tlb_control 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   635
      ButtonWidth     =   1984
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ils_images"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Back"
            Key             =   "back"
            Object.ToolTipText     =   "Proceed to the next picture"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Forward"
            Key             =   "forward"
            Object.ToolTipText     =   "Go back to the previous picture"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Start"
            Key             =   "start"
            Object.ToolTipText     =   "Start the current slideshow"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stop"
            Key             =   "stop"
            Object.ToolTipText     =   "Stop the current slideshow"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Slower"
            Key             =   "slower"
            Object.ToolTipText     =   "Make the slideshow slower"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Faster"
            Key             =   "faster"
            Object.ToolTipText     =   "Make the slideshow faster"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Fullscreen"
            Key             =   "fullscreen"
            Object.ToolTipText     =   "Start fullscreen mode"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.Image img_image 
      Height          =   1815
      Left            =   3000
      MousePointer    =   2  'Cross
      Top             =   780
      Width           =   2295
   End
End
Attribute VB_Name = "frm_slideshow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'**********************************************************************
'PRIVATE VARIABLES
'**********************************************************************
    Private pics() As pic
    Private noofpics As Integer
    Private strpath As String
    Private current As Integer
    Private zoom As Integer

'***********************************
'FORM EVENTS
'***********************************
    Private Sub Form_Resize()
        centre_image
    End Sub

'***********************************
'FOLDER IMAGE GATHERING
'***********************************
    'GET FOLDER CONTENTS
    Public Function scan_folder(spath As String) As Integer
        On Error Resume Next
        strpath = spath
        Dim fs, f, fc, f1
        Dim rndpos As Long
        Dim chkext As Integer
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set f = fs.GetFolder(spath)
        Set fc = f.Files
        noofpics = count_images(spath)
        scan_folder = noofpics
        If noofpics = 1 Then
            MsgBox "There was only picture in the folder so it has been opened as normal", vbOKOnly, "Slideshow error"
            frm_main.file_open (spath & onlyfile)
            Unload Me
        Else
            frm_loading.pro_progress = 0
            frm_loading.pro_progress.Max = noofpics
            noofpics = noofpics - 1
            frm_loading.Caption = "Creating slideshow, please wait..."
            ReDim pics(noofpics) As pic
            For Each f1 In fc
                DoEvents
                frm_loading.pro_progress.Value = frm_loading.pro_progress.Value + 1
                For chkext = 1 To noof_supported_extensions
                    If InStr(1, f1.Name, supported_extensions(chkext).strdata, vbTextCompare) Then
                        Randomize
                        rndpos = Int((noofpics + 1) * Rnd)
                        While pics(rndpos).filename <> ""
                            DoEvents
                            rndpos = Int((noofpics + 1) * Rnd)
                        Wend
                        pics(rndpos).filename = f1.Name
                    End If
                Next chkext
            Next
        End If
    End Function
    
'***********************************
'IMAGE PUBLIC SUBS
'***********************************
    Public Sub centre_image()
        zoom = 0
        If (img_image.width <> pics(current).width) Or (img_image.height <> pics(current).height) Then
            img_image.width = pics(current).width
            img_image.height = pics(current).height
        End If
        img_image.Move (Me.width / 2) - (img_image.width / 2), (Me.height / 2) - (img_image.height / 2)
    End Sub
    
    'NEXT PICTURE - SLIDESHOW
    Public Sub nextpic()
        If current < noofpics Then
            current = current + 1
        Else
            current = 0
        End If
    End Sub
        
    'PREV PICTURE - SLIDESHOW
    Public Sub prevpic()
        If current = 0 Then
            current = noofpics
        Else
            current = current - 1
        End If
    End Sub
    
    'NEXT PICTURE MANUAL
    Public Sub display_pic(Optional restart As Boolean)
    On Error GoTo openerror
        If restart Then current = 0
        If verify_file(strpath & pics(current).filename) Then
            zoom = -1
            img_image.Stretch = False
            Me.Caption = strpath & " - " & current & "/" & noofpics & " - " & pics(current).filename
            img_image.Picture = LoadPicture(strpath & pics(current).filename)
            pics(current).width = img_image.width
            pics(current).height = img_image.height
            centre_image
            Exit Sub
        Else
            MsgBox "File open error", vbOKOnly, "File could not be verified"
            Exit Sub
        End If
openerror:
        MsgBox "File open error", vbOKOnly, "Ibrowse slideshow"
    End Sub

'**********************************************
'TIMER EVENT FOR NEXT PIC
'**********************************************
    Private Sub tim_nextpic_Timer()
        If Me.WindowState = vbMinimized Or frm_main.WindowState = vbMinimized Then Exit Sub
        nextpic
        display_pic
    End Sub

'**********************************************
'MENU EVENTS
'**********************************************
    'TOOLBAR BUTTONS
    Private Sub tlb_control_ButtonClick(ByVal Button As MSComctlLib.Button)
        Dim wasenabled As Boolean
        wasenabled = tim_nextpic.Enabled
        Select Case Button.Key
            Case "back"
                stop_slideshow
                prevpic
                display_pic
            Case "forward"
                stop_slideshow
                nextpic
                display_pic
            Case "start"
                start_slideshow
            Case "stop"
                stop_slideshow
            Case "slower"
                If tim_nextpic.interval < 65000 Then
                    tim_nextpic.Enabled = False
                    tim_nextpic.interval = tim_nextpic.interval + 1000
                    tim_nextpic.Enabled = wasenabled
                End If
            Case "faster"
                If tim_nextpic.interval > 1000 Then
                    tim_nextpic.Enabled = False
                    tim_nextpic.interval = tim_nextpic.interval - 1000
                    tim_nextpic.Enabled = wasenabled
                End If
            Case "fullscreen"
                stop_slideshow
                go_fullscreen
        End Select
    End Sub
    
    'FULL SCREEN MODE
    Private Sub go_fullscreen()
        Dim new_fullscreen As frm_fullscreen
        Set new_fullscreen = New frm_fullscreen
        Load frm_loading
        frm_loading.Caption = "Creating fullscreen slideshow, please wait..."
        frm_loading.Show
        With new_fullscreen
            .pass_slides strpath, noofpics
            .centre_image
            .WindowState = vbMaximized
            .Show
            .set_speed tim_nextpic.interval
            SetWindowPos .hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
            If tim_nextpic.Enabled = True Then
                .start_slideshow True
            Else
                .display_pic True
                .stop_slideshow
            End If
        End With
        Unload frm_loading
    End Sub
        
    'START SLIDESHOW
    Public Sub start_slideshow(Optional restart As Boolean)
        display_pic restart
        tim_nextpic.Enabled = True
        tlb_control.Buttons("start").Enabled = False
        tlb_control.Buttons("stop").Enabled = True
    End Sub
    
    'STOP SLIDESHOW
    Public Sub stop_slideshow()
        tim_nextpic.Enabled = False
        tlb_control.Buttons("start").Enabled = True
        tlb_control.Buttons("stop").Enabled = False
    End Sub

'***********************************
'IMAGE EVENTS
'***********************************
    Private Sub img_image_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If zoom = -1 Then Exit Sub
        If Button = vbLeftButton Then
            If zoom < max_zoom_level Then
                zoom = zoom + 1
                img_image.Visible = False
                img_image.Stretch = True
                img_image.width = img_image.width * 2
                img_image.height = img_image.height * 2
                img_image.Move (Me.width / 2) - (X * 2), (Me.width / 2) - (Y * 2)
                img_image.Visible = True
            End If
        ElseIf Button = vbRightButton Then
            If zoom > 0 Then
                zoom = zoom - 1
                img_image.Visible = False
                img_image.Stretch = True
                img_image.width = img_image.width / 2
                img_image.height = img_image.height / 2
                If zoom > 0 Then
                    img_image.Move (Me.width / 2) - (X / 2), (Me.width / 2) - (Y / 2)
                Else
                    centre_image
                End If
                img_image.Visible = True
            End If
        End If
    End Sub
