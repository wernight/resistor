VERSION 5.00
Begin VB.Form frmCode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Code / Décode un texte"
   ClientHeight    =   3930
   ClientLeft      =   1650
   ClientTop       =   3090
   ClientWidth     =   6495
   Icon            =   "CODE_TXT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3930
   ScaleWidth      =   6495
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picGauge 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   255
      ScaleHeight     =   465
      ScaleWidth      =   15
      TabIndex        =   6
      Top             =   2295
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.CommandButton cmdDécoder 
      Caption         =   "&Décoder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton cmdCoder 
      Caption         =   "&Coder"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox Texte 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   2280
      Width           =   6015
   End
   Begin VB.TextBox txtNbrFoisCode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      MaxLength       =   4
      TabIndex        =   5
      Text            =   "1"
      Top             =   3160
      Width           =   1815
   End
   Begin VB.Label lblNbr 
      BackStyle       =   0  'Transparent
      Caption         =   "Clé de cryptage :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4440
      TabIndex        =   4
      Top             =   2895
      Width           =   1815
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   120
      X2              =   6360
      Y1              =   1815
      Y2              =   1815
   End
   Begin VB.Line Line 
      Index           =   0
      X1              =   120
      X2              =   6360
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "    Vous avez trouvez un ""easter eggs"" qui vous permet de coder/décoder un texte !"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   780
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   6270
      WordWrap        =   -1  'True
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
      Left            =   270
      TabIndex        =   7
      Top             =   120
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Texte :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   660
   End
End
Attribute VB_Name = "frmCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strTexte As String
Private Declare Function Crypte Lib "WBCCrypteDLL" (ByVal pszTexte As String, ByVal lCle As Long) As Long
Private Declare Function Decrypte Lib "WBCCrypteDLL" (ByVal pszTexte As String, ByVal lCle As Long) As Long

Private Sub cmdCoder_Click()
    Screen.MousePointer = 11
    
    ' Crypte
    strTexte = Texte
    Crypte strTexte, Val(txtNbrFoisCode)
    Texte = strTexte
    
    Screen.MousePointer = 0
End Sub

Private Sub cmdDécoder_Click()
    Screen.MousePointer = 11
        
    ' Décrypte
    strTexte = Texte
    Decrypte strTexte, Val(txtNbrFoisCode)
    Texte = strTexte
    
    Screen.MousePointer = 0
End Sub

Private Sub SpinButton_SpinDown()
    If txtNbrFoisCode <> "1" Then
        txtNbrFoisCode = Val(txtNbrFoisCode) - 1
        DoEvents
    End If
End Sub

Private Sub SpinButton_SpinUp()
    txtNbrFoisCode = Val(txtNbrFoisCode) + 1
    DoEvents
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub
