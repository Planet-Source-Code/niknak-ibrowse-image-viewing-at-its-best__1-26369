VERSION 5.00
Begin VB.Form frm_image 
   BackColor       =   &H80000001&
   Caption         =   "<filename>"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frm_image.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3315
   ScaleWidth      =   4680
   Begin VB.Image img_image 
      Height          =   2655
      Left            =   1140
      MousePointer    =   2  'Cross
      Top             =   300
      Width           =   2235
   End
End
Attribute VB_Name = "frm_image"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************
'PRIVATE VARIABLES
'***********************************
    Private zoom As Integer
    Private iwidth As Double
    Private iheight As Double

'***********************************
'FORM EVENTS
'***********************************
    Private Sub Form_Resize()
        centre_image
    End Sub

'***********************************
'IMAGE PUBLIC SUBS
'***********************************
    'LOADS AN IMAGE
    Public Sub load_image(lpstrFile As String)
    On Error GoTo openerror:
        If verify_file(lpstrFile) Then
            img_image.Picture = LoadPicture(lpstrFile)
            iwidth = img_image.width
            iheight = img_image.height
            Exit Sub
        Else
            MsgBox "File open error", vbOKOnly, "File could not be verified"
            Me.Tag = "error"
            Exit Sub
        End If
openerror:
        MsgBox "File open error", vbOKOnly, "Ibrowse"
        Me.Tag = "error"
        Exit Sub
    End Sub
    
    'CENTRES IMAGE
    Public Sub centre_image()
        zoom = 0
        If (img_image.width <> iwidth) Or (img_image.height <> iheight) Then
            img_image.width = iwidth
            img_image.height = iheight
        End If
        img_image.Move (Me.width / 2) - (img_image.width / 2), (Me.height / 2) - (img_image.height / 2)
    End Sub
    
'***********************************
'IMAGE EVENTS
'***********************************
    Private Sub img_image_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
