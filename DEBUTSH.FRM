VERSION 5.00
Begin VB.Form frmD�butSh 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "R�sistor - Code :  1234_4788"
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
      Caption         =   "Donner la cl�"
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
      Caption         =   "R�sistor x.xx Version Shareware"
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
      Caption         =   "R�sistor x.xx Version Shareware"
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
      Caption         =   "Vous �tes au Xe jour de vos 30 jours d'�valuation."
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
      Caption         =   "Vous �tes au Xe jour de vos 30 jours d'�valuation."
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
Attribute VB_Name = "frmD�butSh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Sub cmdActiver_Click()
    Screen.MousePointer = 11
    
    Load frmCl�Active
    Hide
    frmCl�Active.Show

    Screen.MousePointer = 0
End Sub

Private Sub cmdOK_Click()
    If cmdOK.Caption = "Quitter" Then End
    Screen.MousePointer = 11
    
    ' Affiche
    frmD�but!pnlVersion.Visible = Activ�e
    frmD�but!pnlVersion.Refresh
    Hide
    frmD�but.Show 1         'Charge & Affiche
    
    If Activ�e <> 0 Then
        'Active la Version Compl�te
        For i% = 0 To 5
            frmR�sistor!opt4(i%).Enabled = True
            frmCalcC!optT(i%).Enabled = True
        Next
        For i% = 0 To 6 Step 3
            frmCalcC!optOhm(i%).Enabled = True
        Next
    Else
        frmR�sistor!lblSH.Visible = True
        frmR�sistor.Caption = frmR�sistor.Caption & " - #Non Enregistr�#"
        'Affiche shareware et ces fonctionalit�es
        MsgBox "Les fonctions en gris clair sauf 'Calculer valeur' ne sont pas disponibles dans la version shareware." & Enter & "Vous ne disposez pas du 'Didactel'.", vbInformation, "R�sistor - Version Shareware"
    End If
    
    ' D�charge
    Unload Me
    frmR�sistor!lblChargement.Visible = False
    Screen.MousePointer = 0
End Sub

Private Sub cmdReg_Click()
    frmReg.Show 1
End Sub

Private Sub Form_Load()
Dim CodeActive As String, Nom As String, CodeDonn� As String
Dim nCmdPos As Integer, bCmdFait(2) As Boolean
Dim strWindows As String
Dim FichierDate As String, strDateCree As String, DateNow As Date, nJour As Integer
    On Error Resume Next

    Enter = Chr(13) & Chr(10)
    
    #If Compl�te Then
        Activ�e = True
        cmdOK_Click
    #Else
        ' INI
        FichierCode = App.Path & "\Resistor.cle"
        FichierDate = App.Path & "\Resistor.dat"
        
        ' Donne la version
        lblTitre(0) = App.EXEName & " " & App.Major & "." & App.Minor & App.Revision & " Version Shareware"
        lblTitre(1) = lblTitre(0)
        
        ' Cr�� un Code
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
        Line Input #1, Cl�Donn�
        Close
        SetAttr FichierCode, vbReadOnly + vbHidden
        Cl� = "Res_" & Coder(CodeActivation & "_" & Mid(Nom, 1, 1) & Mid(Nom, 4, 1) & Mid(Nom, 5, 1))
        frmD�butSh.Caption = "R�sistor - Code :  " & CodeActivation
        
        ' Test la cl�
        If Cl�Donn� = Cl� And Len(Cl�) >= 10 Then
            Activ�e = True
            cmdOK_Click
            Exit Sub
        End If
        
        ' D�place les boutons
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
        
        ' Enregistre le Jours (1ere ex�cution)
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
        
        ' Test & Nbr de jours �coul�s
        If DateNow < strDateCree Or _
            Left(strDateCree, 9) <> Left(FileDateTime(FichierDate), 9) Then
            nJour = 0
        Else
            nJour = DateNow - FileDateTime(FichierDate) + 1
        End If
        
        ' En fonction de la date...
        Select Case nJour
        Case Is < 1
            MsgBox ("Erreur: Quelqu'un � trafiqu� la date du syst�me")
            Stop
        Case Is <= 30
            lblMsg(0) = "Ceci est une version d'�valuation de R�sistor." & Enter & "Vous pouvez l'utiliser pendant 30 jours. Si apr�s cette p�riode de 30 jours vous voulez continuer d'utiliser R�sistor, vous devez acheter une licence."
            cmdOK.Caption = "Essayer"
            cmdOK.Enabled = True
            Minuterie.Enabled = False
        Case Is <= 60
            lblMsg(0) = "Ceci est une version d'�valuation de R�sistor." & Enter & "Vous avez d�pass� votre p�riode de 30 jours d'essai." & Enter & "Vous devez acheter une licence si vous continuez � vous servir de R�sistor. Cliquez sur 'Commander R�sistor'."
            lblJours(1).BackColor = QBColor(12)
            cmdOK.Caption = "5"
        Case Else
            lblMsg(0) = "Ceci est une version d'�valuation de R�sistor." & Enter & "Vous avez d�pass� la p�riode d'essai depuis" & Str(nJour - 30) & " jours." & Enter & "Vous devez acheter une licence, elle vous permettra d'utiliser R�sistor pendant une p�riode ill�mit�e."
            lblJours(1).BackColor = QBColor(12)
            cmdOK.Caption = "Quitter"
            cmdOK.Enabled = True
            Minuterie.Enabled = False
        End Select
        lblMsg(1) = lblMsg(0)
        lblJours(0) = "Vous �tes au " & nJour & "e jour de vos 30 jours d'�valuation."
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
