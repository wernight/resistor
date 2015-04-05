Attribute VB_Name = "Codage"
Public Enter As String
Public FichierCode As String    'Fichier contenant le Code, Clé, ...
Public Clé As String            'Clé à donner
Public CodeActivation As String 'Code que l'on donne à l'utilisateur
Public Activée As Boolean       'Si c'est la version Complète
Dim Code(3, 20) As String

Private Declare Function Crypte Lib "WBCCrypteDLL" (ByVal pszTexte As String, ByVal lCle As Long) As Long

Public Function Coder(Texte As String) As String
    Coder = Texte
    Crypte Coder, 1910248771
End Function
