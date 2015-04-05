VERSION 5.00
Begin VB.Form frmDébutSh 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Résistor - Code :  1234_4788"
   ClientHeight    =   2790
   ClientLeft      =   3405
   ClientTop       =   3645
   ClientWidth     =   4995
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "DebutSh.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2790
   ScaleWidth      =   4995
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReg 
      Caption         =   "&Commander"
      Default         =   -1  'True
      Height          =   345
      Left            =   1747
      TabIndex        =   1
      Top             =   2280
      Width           =   1485
   End
   Begin VB.Timer Minuterie 
      Interval        =   1000
      Left            =   0
      Top             =   -120
   End
   Begin VB.CommandButton cmdActiver 
      Caption         =   "Donner la clé"
      Height          =   345
      Left            =   3330
      TabIndex        =   2
      Top             =   2280
      Width           =   1485
   End
   Begin VB.CommandButton cmdOK 
      Enabled         =   0   'False
      Height          =   345
      Left            =   165
      TabIndex        =   3
      Top             =   2280
      Width           =   1485
   End
   Begin VB.Line Line 
      Index           =   3
      X1              =   150
      X2              =   4770
      Y1              =   675
      Y2              =   675
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   165
      X2              =   4800
      Y1              =   690
      Y2              =   690
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"DebutSh.frx":030A
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Index           =   0
      Left            =   180
      TabIndex        =   4
      Top             =   870
      Width           =   4650
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTitre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Résistor x.xx Version Shareware"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   1245
      TabIndex        =   0
      Top             =   240
      Width           =   2835
   End
   Begin VB.Image icoVerbes 
      Height          =   480
      Left            =   375
      Picture         =   "DebutSh.frx":03CE
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lblTitre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Résistor x.xx Version Shareware"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   1215
      TabIndex        =   5
      Top             =   225
      Width           =   2835
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"DebutSh.frx":06D8
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   780
      Index           =   1
      Left            =   195
      TabIndex        =   6
      Top             =   885
      Width           =   4650
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblJours 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vous êtes au Xe jour de vos 30 jours d'évaluation."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   225
      TabIndex        =   7
      Top             =   1875
      Width           =   4515
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblJours 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Vous êtes au Xe jour de vos 30 jours d'évaluation."
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
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   1890
      Width           =   4515
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmDébutSh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Sub cmdActiver_Click()
    Screen.MousePointer = 11
    
    Load frmCléActive
    Hide
    frmCléActive.Show

    Screen.MousePointer = 0
End Sub

Private Sub cmdOK_Click()
    If cmdOK.Caption = "Quitter" Then End
    Screen.MousePointer = 11
    
    ' Affiche
    frmDébut!pnlVersion.Visible = Activée
    frmDébut!pnlVersion.Refresh
    Hide
    frmDébut.Show 1         'Charge & Affiche
    
    If Activée <> 0 Then
        'Active la Version Complète
        For i% = 0 To 5
            frmRésistor!opt4(i%).Enabled = True
            frmCalcC!optT(i%).Enabled = True
        Next
        For i% = 0 To 6 Step 3
            frmCalcC!optOhm(i%).Enabled = True
        Next
    Else
        frmRésistor!lblSH.Visible = True
        frmRésistor.Caption = frmRésistor.Caption & " - #Non Enregistré#"
        'Affiche shareware et ces fonctionalitées
        MsgBox "Les fonctions en gris clair sauf 'Calculer valeur' ne sont pas disponibles dans la version shareware." & Enter & "Vous ne disposez pas du 'Didactel'.", vbInformation, "Résistor - Version Shareware"
    End If
    
    ' Décharge
    Unload Me
    frmRésistor!lblChargement.Visible = False
    Screen.MousePointer = 0
End Sub

Private Sub cmdReg_Click()
    frmReg.Show 1
End Sub

Private Sub Form_Load()
Dim CodeActive As String, Nom As String, CodeDonné As String
Dim nCmdPos As Integer, bCmdFait(2) As Boolean
Dim strWindows As String
Dim FichierDate As String, strDateCree As String, DateNow As Date, nJour As Integer
    On Error Resume Next

    Enter = Chr(13) & Chr(10)
    
    #If Complète Then
        Activée = True
        cmdOK_Click
    #Else
        ' INI
        FichierCode = App.Path & "\Resistor.cle"
        FichierDate = App.Path & "\Resistor.dat"
        
        ' Donne la version
        lblTitre(0) = App.EXEName & " " & App.Major & "." & App.Minor & App.Revision & " Version Shareware"
        lblTitre(1) = lblTitre(0)
        
        ' Créé un Code
        If Dir(FichierCode, vbHidden) = "" Then
            Randomize
            Do
                CodeActivation = Int(Rnd * 10000) & "_" & Int(Rnd * 10000)
            Loop While Len("Vrb_" & Coder(CodeActivation)) < 10
            Open FichierCode For Output As #1
            Print #1, CodeActivation
            Close
            SetAttr FichierCode, vbReadOnly Or vbHidden
        End If
        
        ' Sort les infos
        SetAttr FichierCode, vbNormal
        Open FichierCode For Input As #1
        Line Input #1, CodeActivation
        Line Input #1, Nom
        Line Input #1, CléDonné
        Close
        SetAttr FichierCode, vbReadOnly + vbHidden
        Clé = "Res_" & Coder(CodeActivation & "_" & Mid(Nom, 1, 1) & Mid(Nom, 4, 1) & Mid(Nom, 5, 1))
        frmDébutSh.Caption = "Résistor - Code :  " & CodeActivation
        
        ' Test la clé
        If CléDonné = Clé And Len(Clé) >= 10 Then
            Activée = True
            cmdOK_Click
            Exit Sub
        End If
        
        ' Déplace les boutons
        nCmdPos = cmdOK.Left
        Randomize
        For i% = 0 To 2
            Do
                nRnd% = Int(3 * Rnd)
            Loop While bCmdFait(nRnd%)
            bCmdFait(nRnd%) = True
            
            Select Case nRnd%
            Case 0
                cmdOK.Left = nCmdPos
                nCmdPos = nCmdPos + cmdOK.Width + 120
            Case 1
                cmdReg.Left = nCmdPos
                nCmdPos = nCmdPos + cmdReg.Width + 120
            Case 2
                cmdActiver.Left = nCmdPos
                nCmdPos = nCmdPos + cmdActiver.Width + 120
            End Select
        Next
        
        ' Windows Path
        strWindows = String(255, Chr(0))
        nLen% = GetWindowsDirectory(strWindows, Len(strWindows))
        strWindows = Left(strWindows, nLen%)
        
        ' Enregistre le Jours (1ere exécution)
        DateNow = FileDateTime(strWindows & "\USER.DAT")
        If Dir(FichierDate, vbHidden) = "" Then
            Open FichierDate For Output As #1
            Print #1, Trim(Str(DateNow))
            Close
        End If
        SetAttr FichierDate, vbReadOnly Or vbHidden
        
        ' Charge la date
        SetAttr FichierDate, vbNormal
        Open FichierDate For Input As #1
        Line Input #1, strDateCree
        Close
        SetAttr FichierDate, vbReadOnly + vbHidden
        
        ' Test & Nbr de jours écoulés
        If DateNow < strDateCree Or _
            Left(strDateCree, 9) <> Left(FileDateTime(FichierDate), 9) Then
            nJour = 0
        Else
            nJour = DateNow - FileDateTime(FichierDate) + 1
        End If
        
        ' En fonction de la date...
        Select Case nJour
        Case Is < 1
            MsgBox ("Erreur: Quelqu'un à trafiqué la date du système")
            Stop
        Case Is <= 30
            lblMsg(0) = "Ceci est une version d'évaluation de Résistor." & Enter & "Vous pouvez l'utiliser pendant 30 jours. Si après cette période de 30 jours vous voulez continuer d'utiliser Résistor, vous devez acheter une licence."
            cmdOK.Caption = "Essayer"
            cmdOK.Enabled = True
            Minuterie.Enabled = False
        Case Is <= 60
            lblMsg(0) = "Ceci est une version d'évaluation de Résistor." & Enter & "Vous avez dépassé votre période de 30 jours d'essai." & Enter & "Vous devez acheter une licence si vous continuez à vous servir de Résistor. Cliquez sur 'Commander Résistor'."
            lblJours(1).BackColor = QBColor(12)
            cmdOK.Caption = "5"
        Case Else
            lblMsg(0) = "Ceci est une version d'évaluation de Résistor." & Enter & "Vous avez dépassé la période d'essai depuis" & Str(nJour - 30) & " jours." & Enter & "Vous devez acheter une licence, elle vous permettra d'utiliser Résistor pendant une période illémitée."
            lblJours(1).BackColor = QBColor(12)
            cmdOK.Caption = "Quitter"
            cmdOK.Enabled = True
            Minuterie.Enabled = False
        End Select
        lblMsg(1) = lblMsg(0)
        lblJours(0) = "Vous êtes au " & nJour & "e jour de vos 30 jours d'évaluation."
        lblJours(1) = lblJours(0)
    #End If
End Sub

Private Sub Minuterie_Timer()
    ' Attend avant OK=Possible
    If Val(cmdOK.Caption) > 0 Then
        cmdOK.Caption = cmdOK.Caption - 1
    Else
        Minuterie.Enabled = False
        cmdOK.Caption = "Essayer"
        cmdOK.Enabled = True
    End If
End Sub
