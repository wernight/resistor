VERSION 4.00
Begin VB.Form frmActivation 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2190
   ClientLeft      =   1785
   ClientTop       =   5340
   ClientWidth     =   6000
   ControlBox      =   0   'False
   Height          =   2595
   Icon            =   "Active.frx":0000
   Left            =   1725
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   Top             =   4995
   Width           =   6120
   Begin VB.CommandButton cmdActiver 
      Caption         =   "Activer"
      Height          =   330
      Left            =   4815
      TabIndex        =   1
      Top             =   585
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   330
      Left            =   4815
      TabIndex        =   0
      Top             =   180
      Width           =   975
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"Active.frx":030A
      Height          =   780
      Index           =   2
      Left            =   1320
      TabIndex        =   4
      Top             =   1200
      Width           =   4500
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   1320
      X2              =   5760
      Y1              =   1090
      Y2              =   1090
   End
   Begin VB.Line Line 
      Index           =   0
      X1              =   1320
      X2              =   5760
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright� 1997  WBC"
      Height          =   195
      Index           =   1
      Left            =   1320
      TabIndex        =   3
      Top             =   480
      Width           =   1665
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R�sistor Version Shareware"
      Height          =   195
      Index           =   0
      Left            =   1320
      TabIndex        =   2
      Top             =   240
      Width           =   1950
   End
   Begin VB.Image icoR�sistor 
      Height          =   480
      Left            =   240
      Picture         =   "Active.frx":03E6
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmActivation"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Dim Shared Active As Boolean
Private Sub cmdActiver_Click()
    Screen.MousePointer = 11
    
    Load frmMotActive
    Hide
    frmMotActive.Show

    Screen.MousePointer = 0
End Sub
Private Sub cmdOK_Click()
    Screen.MousePointer = 11
    
    ' Centre & Charge
    frmD�but.Move (Screen.Width - frmD�but.Width) / 2, _
        (Screen.Height - frmD�but.Height) / 2
    
    ' Affiche version compl�te
    frmD�but!pnlVersion.Visible = Active
    
    ' Affiche
    Hide
    frmD�but.Show 1         'Charge & Affiche
    ' D�charge
    Unload frmActivation
    
    If Active Then
        Activer
    Else
        ' Affiche shareware et ces fonctionalit�es
        If frmR�sistor!lblsh.Visible Then MsgBox "    Les fonctions en gris clair sauf ""Calculer valeur"" ne sont pas disponibles dans la version shareware.", vbInformation, "R�sistor - Version shareware"
    End If
    
    frmR�sistor!lblChargement.Visible = False
    Screen.MousePointer = 0
End Sub
Private Sub Form_Load()
Dim CodeActive As String, CodeDonn� As String

    ' Centre
    Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2 - 800
    
    On Error Resume Next
    
    ' Si c'est la version Compl�te
    #If Compl�te = True Then
        Active = True
        cmdOK_Click
        
    #Else
        ' Si le fichier existe, sort la cl�
        FichierCode = App.Path & "\RESISTOR.CLE"
        
        If Dir(FichierCode, vbHidden) <> "" Then
            ' Sort les infos
            SetAttr FichierCode, vbNormal
            Open FichierCode For Input As 1
            Line Input #1, CodeActive
            Line Input #1, CodeDonn�
            Close
            SetAttr FichierCode, vbReadOnly + vbHidden
            
            ' Test la cl�
            If (CodeDonn� = "R" & Coder(CodeActive) Or CodeDonn� = "Wernight") And CodeDonn� <> "" Then
                Active = True
                cmdOK_Click
            End If
        End If
    #End If
End Sub
Private Sub Activer()
    ' Active la version compl�te
    frmR�sistor!lblsh.Visible = False
    For i% = 0 To 5
        frmR�sistor!opt4(i%).Enabled = True
        frmCalcC!optT(i%).Enabled = True
    Next
    For i% = 0 To 6 Step 3
        frmCalcC!optOhm(i%).Enabled = True
    Next
End Sub
