VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form frmRésistor 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calcul de la valeur ohmique de résistances"
   ClientHeight    =   5535
   ClientLeft      =   915
   ClientTop       =   2040
   ClientWidth     =   9015
   DrawMode        =   2  'Blackness
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "RESISTOR.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5535
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picRésistance 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1560
      Left            =   -210
      MouseIcon       =   "RESISTOR.frx":030A
      MousePointer    =   99  'Custom
      ScaleHeight     =   104
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   631
      TabIndex        =   58
      Top             =   30
      Width           =   9465
      Begin Threed.SSPanel lblChargement 
         Height          =   525
         Left            =   3000
         TabIndex        =   59
         Top             =   555
         Width           =   3450
         _Version        =   65536
         _ExtentX        =   6085
         _ExtentY        =   926
         _StockProps     =   15
         Caption         =   "Chargement..."
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderWidth     =   4
         BevelInner      =   1
         FloodColor      =   32768
         Font3D          =   2
         Autosize        =   2
         MousePointer    =   11
      End
      Begin VB.Label lblSH 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version shareware"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   240
         TabIndex        =   60
         Top             =   0
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Image CouleurA 
         Height          =   1155
         Index           =   1
         Left            =   3360
         Stretch         =   -1  'True
         Top             =   210
         Width           =   285
      End
      Begin VB.Image CouleurA 
         Height          =   1155
         Index           =   3
         Left            =   5640
         Stretch         =   -1  'True
         Top             =   210
         Width           =   285
      End
      Begin VB.Image CouleurA 
         Height          =   1155
         Index           =   2
         Left            =   4560
         Stretch         =   -1  'True
         Top             =   210
         Width           =   285
      End
   End
   Begin Threed.SSPanel SSPanel 
      Height          =   2535
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   8775
      _Version        =   65536
      _ExtentX        =   15478
      _ExtentY        =   4471
      _StockProps     =   15
      Caption         =   "SSPanel1"
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
      Begin VB.PictureBox imgDidactel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4560
         Picture         =   "RESISTOR.frx":0614
         ScaleHeight     =   585
         ScaleWidth      =   1665
         TabIndex        =   55
         Top             =   1800
         Width           =   1695
      End
      Begin Threed.SSPanel fraAnneau 
         Height          =   630
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   6135
         _Version        =   65536
         _ExtentX        =   10821
         _ExtentY        =   1111
         _StockProps     =   15
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSRibbon opt1 
            Height          =   375
            Index           =   2
            Left            =   1320
            TabIndex        =   61
            Top             =   130
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   65
            BackColor       =   255
            GroupAllowAllUp =   0   'False
            PictureDn       =   "RESISTOR.frx":0E6C
         End
         Begin Threed.SSRibbon opt1 
            Height          =   375
            Index           =   3
            Left            =   1920
            TabIndex        =   13
            Top             =   130
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   65
            BackColor       =   33023
            GroupAllowAllUp =   0   'False
            PictureDn       =   "RESISTOR.frx":1382
         End
         Begin Threed.SSRibbon opt1 
            Height          =   375
            Index           =   4
            Left            =   2520
            TabIndex        =   14
            Top             =   130
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   65
            BackColor       =   65535
            GroupAllowAllUp =   0   'False
            PictureDn       =   "RESISTOR.frx":1898
         End
         Begin Threed.SSRibbon opt1 
            Height          =   375
            Index           =   5
            Left            =   3120
            TabIndex        =   15
            Top             =   120
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   65
            BackColor       =   32768
            GroupAllowAllUp =   0   'False
            PictureDn       =   "RESISTOR.frx":1DAE
         End
         Begin Threed.SSRibbon opt1 
            Height          =   375
            Index           =   6
            Left            =   3720
            TabIndex        =   16
            Top             =   130
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   65
            BackColor       =   16711680
            GroupAllowAllUp =   0   'False
            PictureDn       =   "RESISTOR.frx":22C4
         End
         Begin Threed.SSRibbon opt1 
            Height          =   375
            Index           =   7
            Left            =   4320
            TabIndex        =   17
            Top             =   130
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   65
            BackColor       =   8388736
            GroupAllowAllUp =   0   'False
            PictureDn       =   "RESISTOR.frx":27DA
         End
         Begin Threed.SSRibbon opt1 
            Height          =   375
            Index           =   8
            Left            =   4920
            TabIndex        =   18
            Top             =   130
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   65
            BackColor       =   8421504
            GroupAllowAllUp =   0   'False
            PictureDn       =   "RESISTOR.frx":2CF0
         End
         Begin Threed.SSRibbon opt1 
            Height          =   375
            Index           =   9
            Left            =   5520
            TabIndex        =   19
            Top             =   130
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   65
            BackColor       =   16777215
            GroupAllowAllUp =   0   'False
            PictureDn       =   "RESISTOR.frx":3206
         End
         Begin Threed.SSRibbon opt1 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   130
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   65
            BackColor       =   0
            GroupAllowAllUp =   0   'False
            PictureDn       =   "RESISTOR.frx":371C
         End
         Begin Threed.SSRibbon opt1 
            Height          =   375
            Index           =   1
            Left            =   720
            TabIndex        =   12
            Top             =   130
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   65
            BackColor       =   16512
            GroupAllowAllUp =   0   'False
            PictureDn       =   "RESISTOR.frx":3C32
         End
         Begin VB.Label lbl1erAneau 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "1er anneau"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   2520
            TabIndex        =   10
            Top             =   -30
            Width           =   1095
         End
      End
      Begin Threed.SSPanel fraAnneau 
         Height          =   630
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   6135
         _Version        =   65536
         _ExtentX        =   10821
         _ExtentY        =   1111
         _StockProps     =   15
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSRibbon opt2 
            Height          =   375
            Index           =   9
            Left            =   5520
            TabIndex        =   31
            Top             =   130
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   65
            BackColor       =   16777215
            GroupAllowAllUp =   0   'False
            PictureDn       =   "RESISTOR.frx":4154
         End
         Begin Threed.SSRibbon opt2 
            Height          =   375
            Index           =   8
            Left            =   4920
            TabIndex        =   30
            Top             =   130
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   65
            BackColor       =   8421504
            GroupAllowAllUp =   0   'False
            PictureDn       =   "RESISTOR.frx":4170
         End
         Begin Threed.SSRibbon opt2 
            Height          =   375
            Index           =   7
            Left            =   4320
            TabIndex        =   29
            Top             =   130
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   65
            BackColor       =   8388736
            GroupAllowAllUp =   0   'False
            PictureDn       =   "RESISTOR.frx":418C
         End
         Begin Threed.SSRibbon opt2 
            Height          =   375
            Index           =   6
            Left            =   3720
            TabIndex        =   28
            Top             =   130
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   65
            BackColor       =   16711680
            GroupAllowAllUp =   0   'False
            PictureDn       =   "RESISTOR.frx":41A8
         End
         Begin Threed.SSRibbon opt2 
            Height          =   375
            Index           =   5
            Left            =   3120
            TabIndex        =   27
            Top             =   130
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   65
            BackColor       =   32768
            GroupAllowAllUp =   0   'False
            PictureDn       =   "RESISTOR.frx":41C4
         End
         Begin Threed.SSRibbon opt2 
            Height          =   375
            Index           =   4
            Left            =   2520
            TabIndex        =   26
            Top             =   130
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   65
            BackColor       =   65535
            GroupAllowAllUp =   0   'False
            PictureDn       =   "RESISTOR.frx":41E0
         End
         Begin Threed.SSRibbon opt2 
            Height          =   375
            Index           =   3
            Left            =   1920
            TabIndex        =   25
            Top             =   130
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   65
            BackColor       =   33023
            GroupAllowAllUp =   0   'False
            PictureDn       =   "RESISTOR.frx":41FC
         End
         Begin Threed.SSRibbon opt2 
            Height          =   375
            Index           =   2
            Left            =   1320
            TabIndex        =   24
            Top             =   130
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   65
            BackColor       =   255
            GroupAllowAllUp =   0   'False
            PictureDn       =   "RESISTOR.frx":4218
         End
         Begin Threed.SSRibbon opt2 
            Height          =   375
            Index           =   1
            Left            =   720
            TabIndex        =   23
            Top             =   130
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   65
            BackColor       =   16512
            GroupAllowAllUp =   0   'False
            PictureDn       =   "RESISTOR.frx":4234
         End
         Begin Threed.SSRibbon opt2 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Top             =   130
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   65
            BackColor       =   0
            GroupAllowAllUp =   0   'False
            PictureDn       =   "RESISTOR.frx":4250
         End
         Begin VB.Label lbl2eAnneau 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "2e anneau"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   2520
            TabIndex        =   21
            Top             =   -30
            Width           =   1095
         End
      End
      Begin Threed.SSPanel fraAnneau 
         Height          =   630
         Index           =   2
         Left            =   120
         TabIndex        =   32
         Top             =   1800
         Width           =   4335
         _Version        =   65536
         _ExtentX        =   7646
         _ExtentY        =   1111
         _StockProps     =   15
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSRibbon opt3 
            Height          =   375
            Index           =   6
            Left            =   3720
            TabIndex        =   40
            Top             =   130
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   65
            BackColor       =   16711680
            GroupAllowAllUp =   0   'False
            PictureDn       =   "RESISTOR.frx":426C
         End
         Begin Threed.SSRibbon opt3 
            Height          =   375
            Index           =   5
            Left            =   3120
            TabIndex        =   39
            Top             =   130
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   65
            BackColor       =   32768
            GroupAllowAllUp =   0   'False
            PictureDn       =   "RESISTOR.frx":4288
         End
         Begin Threed.SSRibbon opt3 
            Height          =   375
            Index           =   4
            Left            =   2520
            TabIndex        =   38
            Top             =   130
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   65
            BackColor       =   65535
            GroupAllowAllUp =   0   'False
            PictureDn       =   "RESISTOR.frx":42A4
         End
         Begin Threed.SSRibbon opt3 
            Height          =   375
            Index           =   3
            Left            =   1920
            TabIndex        =   37
            Top             =   130
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   65
            BackColor       =   33023
            GroupAllowAllUp =   0   'False
            PictureDn       =   "RESISTOR.frx":42C0
         End
         Begin Threed.SSRibbon opt3 
            Height          =   375
            Index           =   2
            Left            =   1320
            TabIndex        =   36
            Top             =   130
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   65
            BackColor       =   255
            GroupAllowAllUp =   0   'False
            PictureDn       =   "RESISTOR.frx":42DC
         End
         Begin Threed.SSRibbon opt3 
            Height          =   375
            Index           =   1
            Left            =   720
            TabIndex        =   35
            Top             =   130
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   65
            BackColor       =   16512
            GroupAllowAllUp =   0   'False
            PictureDn       =   "RESISTOR.frx":42F8
         End
         Begin Threed.SSRibbon opt3 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   34
            Top             =   130
            Width           =   495
            _Version        =   65536
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   65
            BackColor       =   0
            GroupAllowAllUp =   0   'False
            PictureDn       =   "RESISTOR.frx":4314
         End
         Begin VB.Label lblMultiplicateur 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Multiplicateur"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2520
            TabIndex        =   33
            Top             =   -30
            Width           =   1095
         End
      End
      Begin Threed.SSPanel fraAnneau 
         Height          =   2310
         Index           =   3
         Left            =   6360
         TabIndex        =   41
         Top             =   120
         Width           =   2295
         _Version        =   65536
         _ExtentX        =   4048
         _ExtentY        =   4075
         _StockProps     =   15
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.OptionButton opt4 
            BackColor       =   &H00808080&
            Caption         =   "Argent"
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   4
            Left            =   120
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   1570
            Width           =   2055
         End
         Begin VB.OptionButton opt4 
            BackColor       =   &H00008000&
            Caption         =   "Vert"
            Enabled         =   0   'False
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   130
            Width           =   2055
         End
         Begin VB.OptionButton opt4 
            BackColor       =   &H00004080&
            Caption         =   "Marron"
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   490
            Width           =   2055
         End
         Begin VB.OptionButton opt4 
            BackColor       =   &H000000FF&
            Caption         =   "Rouge"
            Enabled         =   0   'False
            Height          =   240
            Index           =   2
            Left            =   120
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   850
            Width           =   2055
         End
         Begin VB.OptionButton opt4 
            BackColor       =   &H00008080&
            Caption         =   "Or"
            Height          =   240
            Index           =   3
            Left            =   120
            TabIndex        =   46
            Top             =   1210
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.OptionButton opt4 
            BackColor       =   &H00000000&
            Caption         =   "Noir / Sans couleur"
            Enabled         =   0   'False
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   5
            Left            =   120
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   1930
            Width           =   2055
         End
         Begin VB.Label lblTolérance 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Tolérance"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   600
            TabIndex        =   42
            Top             =   -30
            Width           =   1095
         End
      End
      Begin VB.PictureBox imgDidactelUp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4560
         Picture         =   "RESISTOR.frx":4330
         ScaleHeight     =   585
         ScaleWidth      =   1665
         TabIndex        =   56
         Top             =   1680
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.PictureBox imgDidactelDn 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   4560
         Picture         =   "RESISTOR.frx":4B88
         ScaleHeight     =   585
         ScaleWidth      =   1665
         TabIndex        =   57
         Top             =   1920
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdCalcC 
      Caption         =   "Calculer &couleur"
      Height          =   705
      Left            =   4680
      TabIndex        =   1
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton cmdCalcV 
      Caption         =   "Calculer &valeur"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   705
      Left            =   2880
      TabIndex        =   0
      Top             =   4440
      Width           =   1695
   End
   Begin VB.PictureBox imgAPropos 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   1440
      Picture         =   "RESISTOR.frx":53B2
      ScaleHeight     =   675
      ScaleWidth      =   1305
      TabIndex        =   5
      Top             =   4440
      Width           =   1335
   End
   Begin VB.PictureBox imgQuitter 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   120
      Picture         =   "RESISTOR.frx":5B46
      ScaleHeight     =   675
      ScaleWidth      =   1185
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Timer Minuterie 
      Interval        =   500
      Left            =   1020
      Top             =   5160
   End
   Begin Threed.SSPanel lblRésultat 
      Height          =   495
      Left            =   6480
      TabIndex        =   49
      Top             =   4440
      Width           =   2415
      _Version        =   65536
      _ExtentX        =   4260
      _ExtentY        =   873
      _StockProps     =   15
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Symbol"
         Size            =   18
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   2
      BevelInner      =   1
      Alignment       =   4
   End
   Begin VB.PictureBox imgQuitterUp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   120
      Picture         =   "RESISTOR.frx":63AC
      ScaleHeight     =   675
      ScaleWidth      =   1185
      TabIndex        =   3
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox imgQuitterDn 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   120
      Picture         =   "RESISTOR.frx":6C12
      ScaleHeight     =   675
      ScaleWidth      =   1185
      TabIndex        =   4
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox imgAProposUp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   1440
      Picture         =   "RESISTOR.frx":74F8
      ScaleHeight     =   675
      ScaleWidth      =   1305
      TabIndex        =   6
      Top             =   4320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox imgAProposDn 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   1440
      Picture         =   "RESISTOR.frx":7C8C
      ScaleHeight     =   675
      ScaleWidth      =   1305
      TabIndex        =   7
      Top             =   4560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin PicClip.PictureClip PictureClip 
      Left            =   3345
      Top             =   5265
      _ExtentX        =   5027
      _ExtentY        =   1905
      _Version        =   327680
      Picture         =   "RESISTOR.frx":83D0
   End
   Begin VB.Image imgRésis_GM 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1620
      Index           =   0
      Left            =   120
      Picture         =   "RESISTOR.frx":9F62
      Top             =   5280
      Visible         =   0   'False
      Width           =   9495
   End
   Begin VB.Image imgRésis_GM 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1620
      Index           =   1
      Left            =   240
      Picture         =   "RESISTOR.frx":11B3A
      Top             =   5520
      Visible         =   0   'False
      Width           =   9495
   End
   Begin VB.Image imgRésis_GM 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1620
      Index           =   2
      Left            =   360
      Picture         =   "RESISTOR.frx":19712
      Top             =   5760
      Visible         =   0   'False
      Width           =   9495
   End
   Begin VB.Image imgRésis_GM 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1620
      Index           =   3
      Left            =   480
      Picture         =   "RESISTOR.frx":210BA
      Top             =   6000
      Visible         =   0   'False
      Width           =   9495
   End
   Begin VB.Image imgRésis_GM 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1620
      Index           =   4
      Left            =   600
      Picture         =   "RESISTOR.frx":28CEC
      Top             =   6240
      Visible         =   0   'False
      Width           =   9495
   End
   Begin VB.Label lblValeurInf 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6000
      TabIndex        =   50
      Top             =   4920
      Width           =   1560
   End
   Begin VB.Label lblHeure 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   4680
      TabIndex        =   54
      Top             =   5250
      Width           =   885
   End
   Begin VB.Label lblX 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7560
      TabIndex        =   51
      Top             =   4920
      Width           =   375
   End
   Begin VB.Image picIcone 
      Height          =   480
      Index           =   5
      Left            =   11520
      Picture         =   "RESISTOR.frx":30716
      Top             =   3000
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image picIcone 
      Height          =   480
      Index           =   4
      Left            =   11520
      Picture         =   "RESISTOR.frx":30A20
      Top             =   2400
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image picIcone 
      Height          =   480
      Index           =   3
      Left            =   11520
      Picture         =   "RESISTOR.frx":30D2A
      Top             =   1800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image picIcone 
      Height          =   480
      Index           =   2
      Left            =   11520
      Picture         =   "RESISTOR.frx":31034
      Top             =   1200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image picIcone 
      Height          =   480
      Index           =   1
      Left            =   11520
      Picture         =   "RESISTOR.frx":3133E
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image picIcone 
      Height          =   480
      Index           =   0
      Left            =   11520
      Picture         =   "RESISTOR.frx":31648
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1440
      TabIndex        =   53
      Top             =   5250
      Width           =   3015
   End
   Begin VB.Label lblValeurSup 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7800
      LinkTimeout     =   0
      TabIndex        =   52
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Image imgRésis_GM 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1620
      Index           =   5
      Left            =   720
      Picture         =   "RESISTOR.frx":31952
      Top             =   6480
      Visible         =   0   'False
      Width           =   9495
   End
   Begin VB.Menu Menu_Résistor 
      Caption         =   "&Résistor"
      Begin VB.Menu mnuCalcV 
         Caption         =   "Calculer la &valeur"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCalcC 
         Caption         =   "Calculer la &couleur"
      End
      Begin VB.Menu mnuDidactel 
         Caption         =   "&Didactel"
      End
      Begin VB.Menu mnuSeparateur 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuitter 
         Caption         =   "&Quitter"
      End
   End
   Begin VB.Menu Menu_Aide 
      Caption         =   "&?"
      Begin VB.Menu mnuAide 
         Caption         =   "&Aide"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuSeparateur2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAPropos 
         Caption         =   "A &propos..."
      End
   End
End
Attribute VB_Name = "frmRésistor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nClick As Integer

Private Sub cmdCalcC_Click()
    frmCalcC.Show 1
End Sub

Public Sub cmdCalcV_Click()
Dim Resultat As String, Tolérance As String
Static Fois As Integer
    
    ' Message
    If Not Activée Then
        Fois = Fois + 1
        If Fois >= 10 Then
            Fois = 0
            Screen.MousePointer = 0
            MsgBox "Si vous utilisez frequament ce logiciel, je vous invite à vous enregistrer pour seulement 20 Frs!", vbInformation
        End If
    End If
    
    Screen.MousePointer = 11
    
    ' Calcule les anneaux
    For i% = 0 To 9
        If opt1(i%).Value Then
            Valeur = Valeur + 10 * i%
            CouleurA(1).Picture = PictureClip.GraphicCell(i%)
        End If
        If opt2(i%).Value Then
            Valeur = Valeur + i%
            CouleurA(2).Picture = PictureClip.GraphicCell(i%)
        End If
    Next
    ' Calcule le 3e anneau
    For i% = 0 To 6
        If opt3(i%).Value Then
            Valeur = Valeur * 10 ^ i%
            CouleurA(3).Picture = PictureClip.GraphicCell(i%)
            Exit For
        End If
    Next

    'Calcule le tolérance
    For i% = 0 To 5
        If opt4(i%) Then
            picRésistance.Picture = imgRésis_GM(i%)
            picRésistance.Refresh
            Select Case i%
                Case 0
                    Tolérance = "0.5"
                Case 1
                    Tolérance = "1"
                Case 2
                    Tolérance = "2"
                Case 3
                    Tolérance = "5"
                Case 4
                    Tolérance = "10"
                Case 5
                    Tolérance = "20"
            End Select
        End If
    Next
    
    Select Case Valeur
        Case Is >= 1000000
            Resultat = Str(Valeur / 1000000) & "M"
        Case Is >= 1000
            Resultat = Str(Valeur / 1000) & "k"
        Case Else
            Resultat = Str(Valeur)
    End Select
    Resultat = LTrim(Resultat)
    
    ' Affiche
    lblRésultat.Caption = Resultat & "W ±" & Tolérance & "%"

    ' Calcul le valeur la plus grande et la plus plblte.
    lblX.Caption = "<x<"
    If Right$(Resultat, 1) = "M" Or Right$(Resultat, 1) = "k" Then
        Signe$ = Right$(Resultat, 1) + "W"
    Else
        Signe$ = "W"
    End If
    Différence = Val(Resultat) * Val(Tolérance) / 100
    lblValeurInf.Caption = Str$(Val(Resultat) - Différence) + Signe$
    lblValeurSup.Caption = Str$(Val(Resultat) + Différence) + Signe$
    
    picRésistance.Refresh
    
    Screen.MousePointer = 0
End Sub

Private Sub CouleurA_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Static Précédent As Integer
    ' Test Control + Shift
    If Shift <> 3 Then
        Précédent = 0
        Exit Sub
    End If

    ' Test "easter eggs"
    Select Case Index
        Case 1
            If Précédent = 0 Then
                Précédent = 1
            Else
                Précédent = 0
            End If
        Case 3
            If Précédent = 1 Then
                Précédent = 2
            Else
                Précédent = 0
            End If
        Case 2
            If Précédent = 2 Then
                Screen.MousePointer = 11
                frmCode.Show
                Screen.MousePointer = 0
            Else
                Précédent = 0
            End If
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 68 And Shift = 4
            picRésistance_Click
        Case KeyCode = 27
            Call imgQuitter_Click
    End Select
End Sub

Private Sub Form_Load()
    ' Charge les images OK pour la sélection des couleurs de anneaux
    For i% = 0 To 9
        opt2(i%).PictureDn = opt1(i%).PictureDn
    Next
    For i% = 0 To 6
        opt3(i%).PictureDn = opt1(i%).PictureDn
    Next
    
    ' Gde Résistance = 3 anneaux noir
    picRésistance.Picture = imgRésis_GM(3).Picture
    PictureClip.Cols = 10
    For i% = 1 To 3
        CouleurA(i%).Picture = PictureClip.GraphicCell(0)
    Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmFin.Show 1
End Sub

Private Sub imgAPropos_Click()
    imgAPropos.Refresh
    Screen.MousePointer = 11
    lblChargement.Visible = True
    lblChargement.Refresh
    
    ' Affiche
    Load frmAPropos
    frmAPropos.Show 1
    
    lblChargement.Visible = False
End Sub

Private Sub imgAPropos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then imgAPropos.Picture = imgAProposDn.Picture
End Sub

Private Sub imgAPropos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Si le bouton est activé, affiche le dessin supérieur lorsque la souris glisse
    ' hors de la zone du bouton ; sinon affiche le dessin inférieur.
    Select Case Button
    Case 1
        If X <= 0 Or X > imgAPropos.Width Or Y < 0 Or Y > imgAPropos.Height Then
            imgAPropos.Picture = imgAProposUp.Picture
        Else
            imgAPropos.Picture = imgAProposDn.Picture
        End If
    End Select
End Sub

Private Sub imgAPropos_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then imgAPropos.Picture = imgAProposUp.Picture
End Sub

Private Sub imgDidactel_Click()
    imgDidactel.Refresh
    
    picRésistance_Click
End Sub

Private Sub imgDidactel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then imgDidactel.Picture = imgDidactelDn.Picture
End Sub

Private Sub imgDidactel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Si le bouton est activé, affiche le dessin supérieur lorsque la souris glisse
    ' hors de la zone du bouton ; sinon affiche le dessin inférieur.
    Select Case Button
    Case 1
        If X <= 0 Or X > imgDidactel.Width Or Y < 0 Or Y > imgDidactel.Height Then
            imgDidactel.Picture = imgDidactelUp.Picture
        Else
            imgDidactel.Picture = imgDidactelDn.Picture
        End If
    End Select
End Sub

Private Sub imgDidactel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then imgDidactel.Picture = imgDidactelUp.Picture
End Sub

Private Sub imgQuitter_Click()
    imgQuitter.Refresh
    Screen.MousePointer = 11
    DoEvents
    
    ' Affiche
    Load frmFin
    frmFin.Show 1
End Sub

Private Sub imgQuitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then imgQuitter.Picture = imgQuitterDn.Picture
End Sub

Private Sub imgQuitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Si le bouton est activé, affiche le dessin supérieur lorsque la souris glisse
    ' hors de la zone du bouton ; sinon affiche le dessin inférieur.
    Select Case Button
    Case 1
        If X <= 0 Or X > imgQuitter.Width Or Y < 0 Or Y > imgQuitter.Height Then
            imgQuitter.Picture = imgQuitterUp.Picture
        Else
            imgQuitter.Picture = imgQuitterDn.Picture
        End If
    End Select
End Sub

Private Sub imgQuitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then imgQuitter.Picture = imgQuitterUp.Picture
End Sub

Private Sub picRésistance_Click()
    ' Version Sharewarte
    If Not Activée Then
        MsgBox "L'aide sur les résistances (Didactel) n'est pas disponible dans la version shareware.", vbInformation, "Résistor"
        Exit Sub
    End If
    
    ' Version Complète
    Screen.MousePointer = 11
    lblChargement.Visible = True
    DoEvents
    
    Load frmAideR
    frmAideR.Show
    
    lblChargement.Visible = False
    Screen.MousePointer = 0
End Sub

Public Sub Minuterie_Timer()
Dim DateNow As String
Static num As Integer
    ' Affiche le film
    num = num + 1: If num = 6 Then num = 0
    frmRésistor.Icon = picIcone(num)
    
    ' Afficher l'heure
    If lblHeure = CStr(Time) Then Exit Sub
    lblHeure = Time
    
    ' Affiche la date
    DateNow = Format$(DateValue(Date), "dddd d mmmm yyyy")
    If lblDate = CStr(Date) Then Exit Sub
    lblDate = DateNow
End Sub

Private Sub mnuAide_Click()
    MsgBox "   Pour connaître la valeur correspondant à un code couleur," & Enter & _
        "entrez d 'abord ce code couleur grâce aux boutons d 'options," & Enter & _
        "puis cliquez sur «Calculer Valeur»."
End Sub

Private Sub mnuAPropos_Click()
    imgAPropos_Click
End Sub

Private Sub mnuCalcC_Click()
    cmdCalcC_Click
End Sub

Private Sub mnuCalcV_Click()
    cmdCalcV_Click
End Sub

Private Sub mnuDidactel_Click()
    picRésistance_Click
End Sub

Private Sub mnuQuitter_Click()
    Call imgQuitter_Click
End Sub

Private Static Sub FoisClick()
    If nClick = 7 Then
        cmdCalcV.Enabled = True
        mnuCalcV.Enabled = True
    End If
End Sub

Private Sub opt1_Click(Index As Integer, Value As Integer)
    nClick = nClick Or 1
    FoisClick
End Sub

Private Sub opt2_Click(Index As Integer, Value As Integer)
    nClick = nClick Or 2
    FoisClick
End Sub

Private Sub opt3_Click(Index As Integer, Value As Integer)
    nClick = nClick Or 4
    FoisClick
End Sub
