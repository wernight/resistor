Attribute VB_Name = "Codage"
Public Enter As String
Public FichierCode As String    'Fichier contenant le Code, Cl�, ...
Public Cl� As String            'Cl� � donner
Public CodeActivation As String 'Code que l'on donne � l'utilisateur
Public Activ�e As Boolean       'Si c'est la version Compl�te
Dim Code(3, 20) As String

Private Declare Function Crypte Lib "WBCCrypteDLL" (ByVal pszTexte As String, ByVal lCle As Long) As Long

Public Function Coder(Texte As String) As String
    Coder = Texte
    Crypte Coder, 1910248771
End Function
