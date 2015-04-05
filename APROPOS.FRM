VERSION 5.00
Begin VB.Form frmAProposAv 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "A propos..."
   ClientHeight    =   5055
   ClientLeft      =   1455
   ClientTop       =   2490
   ClientWidth     =   6615
   FillColor       =   &H00C0C0C0&
   FillStyle       =   0  'Solid
   Icon            =   "APROPOS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "APROPOS.frx":030A
   ScaleHeight     =   337
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   441
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Minuterie 
      Interval        =   10
      Left            =   120
      Top             =   3960
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Werner"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   1
      Left            =   -1050
      TabIndex        =   1
      Top             =   3960
      Width           =   1050
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Création:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   360
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   -300
      Width           =   6615
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "    Version x.xx"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      TabIndex        =   4
      Top             =   4560
      Width           =   6615
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1997"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   3
      Left            =   6750
      TabIndex        =   3
      Top             =   3960
      Width           =   675
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BEROUX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   2
      Left            =   6750
      TabIndex        =   0
      Top             =   3960
      Width           =   1305
   End
   Begin VB.Label lblInfos 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   120
      TabIndex        =   5
      Top             =   135
      Width           =   6375
   End
End
Attribute VB_Name = "frmAProposAv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Fois As Integer
Private Sub Form_Activate()
    ' Centre
    Me.Move (Screen.Width - Me.Width) / 2, _
        (Screen.Height - Me.Height) / 2
    
    Fois = 0
    
    Screen.MousePointer = 0
End Sub
Private Sub Form_Click()
    Rem Easter Eggs
    ' Test Couleurs
    If frmRésistor!opt1(4) And frmRésistor!opt2(7) And frmRésistor!opt3(1) Then
        ' Affiche
        Unload Me
        frmEasterEggs.Show
    End If
    
    Rem Infos
    If lblInfos = "" Then
        'Message
        If Activée Then
            lblInfos.Top = 40
            Msg$ = "    Il est interdit de changer le code de ce programme ou de le copier!!!"
        Else
            Msg$ = "    Il est interdit de changer le code de ce programme !!!" & Enter & Enter & _
                "    Pour commander la version complète, envoyez-moi 20 F suivient de vos Nom, Prénom, Adresse, [Tel.] [FAX] [E-Mail], Votre CODE: " & CodeActivation
        End If
        Msg$ = Msg$ & Enter & Enter & "L'auteur ne peut pas être tenu pour responsable des problèmes liés à l'utilisation de ce programme." & Enter & "    Merci de me faire part de vos remarques." & Enter & Enter & _
        "     BEROUX Werner" & Enter & "     Chemin de Labadier " & Enter & "     30400 Villeneuve-les-Avignon." & Enter & "     Tel && FAX: 04 90 25 96 91" & Enter & "     E-Mail : WernerBeroux@Compuserve.com" & Enter & Enter & _
        "    Pour en savoir plus, consultez LISEZMOI.COM."
        
        'Affiche
        lblInfos = Msg$
    Else
        'Quitte
        Screen.MousePointer = 11
        DoEvents
        
        Unload Me
        
        Screen.MousePointer = 0
    End If
End Sub
Private Sub Form_Load()
    ' Donne la version
    lblVersion = "Version " & App.Major & "." & App.Minor & App.Revision
End Sub
Private Sub lblInfos_Click()
    ' Affiche les Infos
    Form_Click
End Sub
Private Sub Minuterie_Timer()
    Fois = Fois + 10
    Select Case Fois
        Case Is < 241
            Label(0).Top = Fois
        Case 250 To 370
            Label(1).Left = Int((Fois - 320) * 1.5)
        Case 380 To 510
            Label(2).Left = 430 - ((Fois - 380) * 2)
        Case 520 To 560
            Label(3).Left = Int(425 - ((Fois - 520) * 3))
    End Select
End Sub
