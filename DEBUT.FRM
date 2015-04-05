VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmDébut 
   BorderStyle     =   0  'None
   ClientHeight    =   5490
   ClientLeft      =   1080
   ClientTop       =   1515
   ClientWidth     =   6810
   Icon            =   "DEBUT.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   PaletteMode     =   1  'UseZOrder
   Picture         =   "DEBUT.frx":030A
   ScaleHeight     =   366
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   454
   StartUpPosition =   2  'CenterScreen
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
      Caption         =   "Version complète"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   5
      BevelOuter      =   1
      BevelInner      =   2
      Font3D          =   2
      Autosize        =   2
   End
   Begin VB.Shape shpPourcent 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00800000&
      Height          =   390
      Left            =   2670
      Top             =   4380
      Visible         =   0   'False
      Width           =   15
   End
End
Attribute VB_Name = "frmDébut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    ' Rafréchit
    Me.Refresh
    DoEvents
    
    ' Charge est affiche ou il en est
    Load frmRésistor
    shpPourcent.Visible = True
    shpPourcent.Width = 217

    
    Load frmCalcC
    shpPourcent.Width = 245
    
    ' Enlève & Affiche
    Call frmRésistor.Minuterie_Timer
    Hide
    frmRésistor.Show
    Unload frmDébut
End Sub
