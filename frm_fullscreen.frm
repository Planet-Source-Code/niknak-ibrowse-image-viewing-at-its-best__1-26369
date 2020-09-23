VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_fullscreen 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tim_nextpic 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   960
      Top             =   1920
   End
   Begin MSComctlLib.ImageList ils_images 
      Left            =   180
      Top             =   1800
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
            Picture         =   "frm_fullscreen.frx":0000
            Key             =   "stop"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_fullscreen.frx":059A
            Key             =   "increase"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_fullscreen.frx":0B34
            Key             =   "decrease"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_fullscreen.frx":10CE
            Key             =   "close"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_fullscreen.frx":1668
            Key             =   "interval"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_fullscreen.frx":1C02
            Key             =   "forward"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_fullscreen.frx":219C
            Key             =   "start"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_fullscreen.frx":2736
            Key             =   "back"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb_control 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   635
      ButtonWidth     =   1720
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
            Caption         =   "Close"
            Key             =   "close"
            Object.ToolTipText     =   "Close full screen mode"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl_note 
      BackStyle       =   0  'Transparent
      Caption         =   "<image note>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   7275
   End
   Begin VB.Image img_image 
      Height          =   2055
      Left            =   2460
      MousePointer    =   2  'Cross
      Top             =   600
      Width           =   2415
   End
End
Attribute VB_Name = "frm_fullscreen"
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
    'RESIZE
    Private Sub Form_Resize()
        img_image.Move (Me.width / 2) - (img_image.width / 2), (Me.height / 2) - (img_image.height / 2)
        lbl_note.Move 0, lbl_note.Top, Me.width
    End Sub

    'MOUSE MOVE FOR CONTROL AUTO HIDING
    Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Y > tlb_control.height * 2 Then
            tlb_control.Visible = False
            lbl_note.Top = 0
        Else
            tlb_control.Visible = True
            lbl_note.Top = tlb_control.height
        End If
    End Sub

'***********************************
'FOLDER IMAGE GATHERING
'***********************************
    'GET FOLDER CONTENTS
    Public Sub pass_slides(spath As String, noofspics As Integer)
        On Error Resume Next
        strpath = spath
        noofpics = noofspics
        Dim fs, f, fc, f1
        Dim rndpos As Long
        Dim chkext As Integer
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set f = fs.GetFolder(spath)
        Set fc = f.Files
        frm_loading.pro_progress = 0
        frm_loading.pro_progress.Max = noofpics
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
    End Sub
    
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
    
    'SET SPEED
    Public Sub set_speed(interval As Integer)
        tim_nextpic.Enabled = False
        tim_nextpic.interval = interval
        tim_nextpic.Enabled = True
    End Sub
       
    'NEXT PICTURE MANUAL
    Public Sub display_pic(Optional restart As Boolean)
    On Error GoTo openerror
        If restart Then current = 0
        If verify_file(strpath & pics(current).filename) Then
            zoom = -1
            img_image.Stretch = False
            lbl_note.Caption = strpath & " - " & current & "/" & noofpics & " - " & pics(current).filename
            img_image.Picture = LoadPicture(strpath & pics(current).filename)
            pics(current).width = img_image.width
            pics(current).height = img_image.height
            centre_image
            Exit Sub
        End If
openerror:
        DoEvents
    End Sub

'**********************************************
'NOTE EVENTS
'**********************************************
    Private Sub lbl_note_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        tlb_control.Visible = True
        lbl_note.Top = tlb_control.height
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
            Case "close"
                Unload Me
        End Select
    End Sub
    
    'MOUSE MOVE
    Private Sub tlb_control_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        tlb_control.Visible = True
        lbl_note.Top = tlb_control.height
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
    'MOUSE DOWN
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
    
    'MOUSE MOVE
    Private Sub img_image_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        tlb_control.Visible = False
        lbl_note.Top = 0
    End Sub
