VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmFin 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2550
   ClientLeft      =   2490
   ClientTop       =   3240
   ClientWidth     =   4335
   ControlBox      =   0   'False
   Icon            =   "FIN.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2550
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox imgOui 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   810
      Picture         =   "FIN.frx":000C
      ScaleHeight     =   555
      ScaleWidth      =   915
      TabIndex        =   1
      Top             =   1800
      Width           =   945
   End
   Begin VB.PictureBox imgNon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   2595
      Picture         =   "FIN.frx":071C
      ScaleHeight     =   555
      ScaleWidth      =   915
      TabIndex        =   0
      Top             =   1800
      Width           =   945
   End
   Begin VB.PictureBox imgNonDn 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   2700
      Picture         =   "FIN.frx":0F2A
      ScaleHeight     =   555
      ScaleWidth      =   915
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox ImgNonUp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   2820
      Picture         =   "FIN.frx":171A
      ScaleHeight     =   555
      ScaleWidth      =   915
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox imgOuiDn 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   900
      Picture         =   "FIN.frx":1F28
      ScaleHeight     =   555
      ScaleWidth      =   915
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox imgOuiUp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   1020
      Picture         =   "FIN.frx":25FC
      ScaleHeight     =   555
      ScaleWidth      =   915
      TabIndex        =   5
      Top             =   1800
      Visible         =   0   'False
      Width           =   945
   End
   Begin Threed.SSPanel SSpnlQuitter 
      Height          =   1215
      Left            =   195
      TabIndex        =   6
      Top             =   180
      Width           =   3945
      _Version        =   65536
      _ExtentX        =   6959
      _ExtentY        =   2143
      _StockProps     =   15
      Caption         =   "                 Quitter Résistor"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      Alignment       =   1
      Begin VB.Image imgvbInterrogation 
         Height          =   810
         Left            =   210
         Picture         =   "FIN.frx":2D0C
         Top             =   180
         Width           =   495
      End
   End
   Begin VB.Line Line 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   -15
      X2              =   4335
      Y1              =   1590
      Y2              =   1590
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   0
      X2              =   4335
      Y1              =   1605
      Y2              =   1605
   End
End
Attribute VB_Name = "frmFin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    Screen.MousePointer = 0
    Beep
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    DoEvents
    Select Case KeyCode
        Case 79 And Shift = 4
            End
        Case 13
            End
        Case 78 And Shift = 4
            frmFin.Hide
        Case 27
            frmFin.Hide
    End Select
End Sub

Private Sub imgNon_Click()
    imgNon.Refresh
    Screen.MousePointer = 11
    DoEvents
    
    Unload Me
    
    Screen.MousePointer = 0
End Sub

Private Sub imgNon_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then imgNon.Picture = imgNonDn.Picture
End Sub

Private Sub imgNon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Si le bouton est activé, affiche le dessin supérieur lorsque la souris glisse
    ' hors de la zone du bouton ; sinon affiche le dessin inférieur.
    Select Case Button
    Case 1
        If X <= 0 Or X > imgNon.Width Or Y < 0 Or Y > imgNon.Height Then
            imgNon.Picture = ImgNonUp.Picture
        Else
            imgNon.Picture = imgNonDn.Picture
        End If
    End Select
End Sub

Private Sub imgNon_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then imgNon.Picture = ImgNonUp.Picture
End Sub

Private Sub imgOui_Click()
    imgOui.Refresh
    End
End Sub

Private Sub imgOui_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgOui.Picture = imgOuiDn.Picture
End Sub

Private Sub imgOui_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Si le bouton est activé, affiche le dessin supérieur lorsque la souris glisse
    ' hors de la zone du bouton ; sinon affiche le dessin inférieur.
    Select Case Button
    Case 1
        If X <= 0 Or X > imgOui.Width Or Y < 0 Or Y > imgOui.Height Then
            imgOui.Picture = imgOuiUp.Picture
        Else
            imgOui.Picture = imgOuiDn.Picture
        End If
    End Select
End Sub

Private Sub imgOui_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgOui.Picture = imgOuiUp.Picture
End Sub
