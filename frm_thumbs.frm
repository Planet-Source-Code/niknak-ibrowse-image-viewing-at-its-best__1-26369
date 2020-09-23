VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm_thumbs 
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "<folder name>"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   Icon            =   "frm_thumbs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5430
   ScaleWidth      =   6510
   Begin MSComctlLib.ImageList ils_images 
      Left            =   120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_thumbs.frx":058A
            Key             =   "noimage"
         EndProperty
      EndProperty
   End
   Begin VB.HScrollBar hsc_scroll 
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   5100
      Width           =   6375
   End
   Begin VB.Image img_thumb 
      Height          =   1575
      Index           =   11
      Left            =   4920
      MousePointer    =   2  'Cross
      Stretch         =   -1  'True
      Top             =   3420
      Width           =   1515
   End
   Begin VB.Image img_thumb 
      Height          =   1575
      Index           =   10
      Left            =   3300
      MousePointer    =   2  'Cross
      Stretch         =   -1  'True
      Top             =   3420
      Width           =   1515
   End
   Begin VB.Image img_thumb 
      Height          =   1575
      Index           =   9
      Left            =   1680
      MousePointer    =   2  'Cross
      Stretch         =   -1  'True
      Top             =   3420
      Width           =   1515
   End
   Begin VB.Image img_thumb 
      Height          =   1575
      Index           =   8
      Left            =   60
      MousePointer    =   2  'Cross
      Stretch         =   -1  'True
      Top             =   3420
      Width           =   1515
   End
   Begin VB.Image img_thumb 
      Height          =   1575
      Index           =   7
      Left            =   4920
      MousePointer    =   2  'Cross
      Stretch         =   -1  'True
      Top             =   1740
      Width           =   1515
   End
   Begin VB.Image img_thumb 
      Height          =   1575
      Index           =   6
      Left            =   3300
      MousePointer    =   2  'Cross
      Stretch         =   -1  'True
      Top             =   1740
      Width           =   1515
   End
   Begin VB.Image img_thumb 
      Height          =   1575
      Index           =   5
      Left            =   1680
      MousePointer    =   2  'Cross
      Stretch         =   -1  'True
      Top             =   1740
      Width           =   1515
   End
   Begin VB.Image img_thumb 
      Height          =   1575
      Index           =   4
      Left            =   60
      MousePointer    =   2  'Cross
      Stretch         =   -1  'True
      Top             =   1740
      Width           =   1515
   End
   Begin VB.Image img_thumb 
      Height          =   1575
      Index           =   3
      Left            =   4920
      MousePointer    =   2  'Cross
      Stretch         =   -1  'True
      Top             =   60
      Width           =   1515
   End
   Begin VB.Image img_thumb 
      Height          =   1575
      Index           =   2
      Left            =   3300
      MousePointer    =   2  'Cross
      Stretch         =   -1  'True
      Top             =   60
      Width           =   1515
   End
   Begin VB.Image img_thumb 
      Height          =   1575
      Index           =   1
      Left            =   1680
      MousePointer    =   2  'Cross
      Stretch         =   -1  'True
      Top             =   60
      Width           =   1515
   End
   Begin VB.Image img_thumb 
      Height          =   1575
      Index           =   0
      Left            =   60
      MousePointer    =   2  'Cross
      Stretch         =   -1  'True
      Top             =   60
      Width           =   1515
   End
End
Attribute VB_Name = "frm_thumbs"
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
            If noofpics > 12 Then
                hsc_scroll.Max = noofpics - 12
            Else
                hsc_scroll.Visible = False
            End If
            noofpics = noofpics - 1
            frm_loading.Caption = "Creating thumbnails, please wait..."
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
'DISPLAY THUMBS
'***********************************
    Public Sub display_thumbs()
        Dim thumbnum As Integer
        Dim first As Integer
        first = hsc_scroll
        For thumbnum = 0 To 11
            display_pic thumbnum, first
            first = first + 1
            If first > noofpics Then Exit Sub
        Next thumbnum
    End Sub

'***********************************
'IMAGE PUBLIC SUBS
'***********************************
    'NEXT PICTURE MANUAL
    Public Sub display_pic(Index As Integer, picindex As Integer)
    On Error GoTo openerror
        If verify_file(strpath & pics(picindex).filename) Then
            img_thumb(Index).Picture = LoadPicture(strpath & pics(picindex).filename)
            img_thumb(Index).ToolTipText = "Double click to open " & pics(picindex).filename
            img_thumb(Index).Tag = strpath & pics(picindex).filename
            Exit Sub
        Else
openerror:
            img_thumb(Index).Picture = ils_images.ListImages("noimage").Picture
            img_thumb(Index).Tag = ""
            Exit Sub
        End If
    End Sub
    
'***********************************
'SCROLL EVENTS
'***********************************
    Private Sub hsc_scroll_Change()
        display_thumbs
    End Sub

'***********************************
'IMAGE EVENTS
'***********************************
    Private Sub img_thumb_DblClick(Index As Integer)
        If img_thumb(Index).Tag <> "" Then frm_main.file_open img_thumb(Index).Tag
    End Sub
