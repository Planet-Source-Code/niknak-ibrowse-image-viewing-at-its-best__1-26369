VERSION 5.00
Begin VB.Form frm_about 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ibrowse"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   Icon            =   "frm_about.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   3945
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lbl_copyw 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (c) 2001, Nicholas Phillip Pateman"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   2760
      Width           =   3795
   End
   Begin VB.Label lbl_version 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   2460
      Width           =   3795
   End
   Begin VB.Image img_splash 
      BorderStyle     =   1  'Fixed Single
      Height          =   2310
      Left            =   60
      Picture         =   "frm_about.frx":08CA
      Top             =   60
      Width           =   3810
   End
End
Attribute VB_Name = "frm_about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************
'API DECLARATIONS
'***********************************
    Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long

'***********************************
'FORM EVENTS
'***********************************
    'LOAD
    Private Sub Form_Load()
        '---------------------------
        'UPDATE VERSION LABEL
        lbl_version.Caption = "Version " & App.Major & "." & App.Minor & ":" & App.Revision
        '---------------------------
        If loaded Then Exit Sub
        '---------------------------
        'CHECK FOR PREVIOUS INSTANCE
        If App.PrevInstance Then
            MsgBox "There is an instance of Ibrowse already running", vbOKOnly, "Ibrowse loading error"
            End
        End If
        '---------------------------
        'CHECK COMMAND LINE VARIABLES
        If Command = "" Then Exit Sub
        If verify_file(Command) Then
            Dim chkext As Integer
            For chkext = 1 To noof_supported_extensions
                If InStr(1, Command, supported_extensions(chkext).strdata, vbTextCompare) Then
                    Load frm_main
                    frm_main.Show
                    frm_main.file_open Command
                    Unload Me
                End If
            Next chkext
        Else
            If PathIsDirectory(Command & "\") Then
                Load frm_main
                frm_main.Show
                frm_main.start_slideshow Command & "\"
                Unload Me
            ElseIf PathIsDirectory(Command) Then
                Load frm_main
                frm_main.Show
                frm_main.start_slideshow Command
                Unload Me
            End If
        End If
    End Sub

    'UNLOAD
    Private Sub Form_Unload(Cancel As Integer)
        If Not loaded Then
            Load frm_main
            Me.Hide
            frm_main.Show
            loaded = True
        End If
    End Sub

'***********************************
'SPLASH IMAGE EVENTS
'***********************************
    Private Sub img_splash_Click()
        Unload Me
    End Sub
