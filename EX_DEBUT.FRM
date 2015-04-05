VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmDébut 
   BorderStyle     =   0  'None
   ClientHeight    =   5490
   ClientLeft      =   1785
   ClientTop       =   1920
   ClientWidth     =   6810
   Icon            =   "EX_DEBUT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "EX_DEBUT.frx":030A
   ScaleHeight     =   366
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   454
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel pnlVersion 
      Height          =   540
      Left            =   2640
      TabIndex        =   0
      Top             =   3720
      Width           =   3750
      _Version        =   65536
      _ExtentX        =   6615
      _ExtentY        =   953
      _StockProps     =   15
      Caption         =   "Version shareware"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   13.51
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   5
      BevelOuter      =   1
      BevelInner      =   2
      FloodShowPct    =   0   'False
      Font3D          =   2
      Autosize        =   2
   End
   Begin Threed.SSPanel pnlRésistor 
      Height          =   975
      Left            =   2640
      TabIndex        =   4
      Top             =   2640
      Width           =   3975
      _Version        =   65536
      _ExtentX        =   7011
      _ExtentY        =   1720
      _StockProps     =   15
      Caption         =   "Résistor 2.3"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   30.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   4
      BorderWidth     =   1
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      Font3D          =   4
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   450
      Left            =   2640
      TabIndex        =   5
      Top             =   4350
      Width           =   3750
      _Version        =   65536
      _ExtentX        =   6615
      _ExtentY        =   794
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   2
      BevelOuter      =   1
   End
   Begin VB.Label lblChargement 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Chargement en cours . . ."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   885
      Left            =   375
      TabIndex        =   3
      Top             =   1680
      Width           =   6060
   End
   Begin VB.Label lbl1997 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1997"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5940
      TabIndex        =   2
      Top             =   4875
      Width           =   555
   End
   Begin VB.Label lblWBC 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Création:   Werner  BEROUX"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   360
      TabIndex        =   1
      Top             =   4875
      Width           =   3255
   End
End
Attribute VB_Name = "frmDébut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdRésistor_Click()
    End
End Sub


Private Sub Form_Load()
    Me.Move (Screen.Width - Me.Width) / 2, _
        (Screen.Height - Me.Height) / 2
End Sub



Private Sub Minuterie_Timer()
    Load frmRésistor
    Load frmAideR
    frmRésistor.Show
End Sub
Private Sub pnlRésistor_Click()
    End
End Sub


