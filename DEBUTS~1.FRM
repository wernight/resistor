VERSION 5.00
Begin VB.Form frmActivationAv 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2325
   ClientLeft      =   1830
   ClientTop       =   1305
   ClientWidth     =   6000
   ControlBox      =   0   'False
   Icon            =   "debutshAv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2325
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdReg 
      Caption         =   "&Enregistrement"
      Default         =   -1  'True
      Height          =   330
      Left            =   2520
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.Timer Minuterie 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdActiver 
      Caption         =   "Activer"
      Height          =   330
      Left            =   4815
      TabIndex        =   3
      Top             =   585
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "10"
      Enabled         =   0   'False
      Height          =   330
      Left            =   4815
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   180
      Width           =   975
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"debutshAv.frx":030A
      Height          =   585
      Index           =   2
      Left            =   1320
      TabIndex        =   5
      Top             =   1560
      Width           =   4500
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   1335
      X2              =   5785
      Y1              =   1455
      Y2              =   1455
   End
   Begin VB.Line Line 
      Index           =   0
      X1              =   1320
      X2              =   5760
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright� 1998  WBC"
      Height          =   195
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   1665
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R�sistor x.xx Version Shareware"
      Height          =   195
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   2265
   End
   Begin VB.Image icoR�sistor 
      Height          =   480
      Left            =   240
      Picture         =   "debutshAv.frx":03C7
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmActivationAv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdActiver_Click()
    Screen.MousePointer = 11
    
    Load frmCl�Active
    Hide
    frmCl�Active.Show

    Screen.MousePointer = 0
End Sub

Private Sub cmdOK_Click()
    Screen.MousePointer = 11
    
    ' Centre & Charge
    frmD�but.Move (Screen.Width - frmD�but.Width) / 2, _
        (Screen.Height - frmD�but.Height) / 2
    
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
    Unload frmActivation
    frmR�sistor!lblChargement.Visible = False
    Screen.MousePointer = 0
End Sub

Private Sub cmdReg_Click()
    frmReg.Show 1
End Sub

Private Sub Form_Load()
Dim Cl�Donn� As String
    On Error Resume Next

    ' Centre
    Move (Screen.Width - Me.Width) / 2, _
        (Screen.Height - Me.Height) / 2 - 800
    
    Enter = Chr(13) & Chr(10)
    
    ' Si c'est la version Compl�te
    #If Compl�te <> 0 Then
        Activ�e = True
        cmdOK_Click
    #Else
        ' Donne la version
        Label(0) = "R�sistor " & App.Major & "." & App.Minor & App.Revision & " Version Shareware"
        
        ' Donne le fichier qui contient la Cl�
        FichierCode = App.Path & "\RESISTOR.cle"
        
        If Dir(FichierCode, vbHidden) = "" Then
NewCode:
            'Si: fichier n'�xiste pas, cr�� un Code
            Randomize
            CodeActivation = Rnd * 10000
            Open FichierCode For Output As 1
            Print #1, CodeActivation
            Close
            SetAttr FichierCode, vbReadOnly + vbHidden
        Else
            ' Sort les infos
            SetAttr FichierCode, vbNormal
            Open FichierCode For Input As 1
            Line Input #1, CodeActivation
            Line Input #1, Cl�Donn�
            Close
            SetAttr FichierCode, vbReadOnly + vbHidden
            Cl� = "R" & Coder(CodeActivation)
            
            ' Test la cl�
            If Cl�Donn� = Cl� Then
                Activ�e = True
                cmdOK_Click
            End If
        End If
    #End If
End Sub

Private Sub Minuterie_Timer()
    Static Temps As Integer

    ' Attend avant OK=Possible
    Temps = Temps + 1
    If Temps < 10 Then
        cmdOK.Caption = 10 - Temps
    Else
        cmdOK.Caption = "OK"
        cmdOK.Enabled = True
    End If
End Sub
