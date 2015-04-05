VERSION 5.00
Begin VB.Form frmEasterEggs 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3225
   ClientLeft      =   1770
   ClientTop       =   3720
   ClientWidth     =   6405
   ClipControls    =   0   'False
   Icon            =   "EGGS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   215
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   427
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer 
      Interval        =   600
      Left            =   0
      Top             =   0
   End
   Begin VB.Image imgWerner 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2670
      Left            =   390
      Picture         =   "EGGS.frx":000C
      Top             =   285
      Width           =   1860
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bravo !"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   390
      Index           =   0
      Left            =   2760
      TabIndex        =   1
      Top             =   405
      Width           =   1050
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"EGGS.frx":424E
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Index           =   1
      Left            =   2520
      TabIndex        =   0
      Top             =   1200
      Width           =   3390
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmEasterEggs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_LostFocus()
    Unload Me
End Sub
