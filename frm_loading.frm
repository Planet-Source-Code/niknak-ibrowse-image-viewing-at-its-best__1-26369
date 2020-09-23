VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_loading 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "<event description>"
   ClientHeight    =   615
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   615
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tim_anim 
      Interval        =   100
      Left            =   1440
      Top             =   120
   End
   Begin MSComctlLib.ImageList ils_images 
      Left            =   780
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_loading.frx":0000
            Key             =   "find4"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_loading.frx":01D3
            Key             =   "find2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_loading.frx":03B6
            Key             =   "find1"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_loading.frx":0612
            Key             =   "find3"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pro_progress 
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   60
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Image img_progress 
      Height          =   480
      Left            =   60
      Top             =   60
      Width           =   480
   End
End
Attribute VB_Name = "frm_loading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***********************************
'ANIM TIMER EVENTS
'***********************************
    Private Sub tim_anim_Timer()
        Static frame As Integer
        If frame = 4 Or frame = 0 Then
            frame = 1
        Else
            frame = frame + 1
        End If
        img_progress.Picture = ils_images.ListImages(Replace("find" & Str(frame), " ", "")).Picture
    End Sub
