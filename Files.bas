Attribute VB_Name = "Files"
' Copie avec sous r�pertoires / Taille d'un r�pertoire et sous r�pertoires
' Copyright (c) W. BEROUX 1999
' www.alc-wbc.com
' wbc@alc-wbc.com

Public Function DirSize(Directory As String, Optional Attrib As Integer = vbArchive Or vbHidden Or vbDirectory) As Long
    Dim f As String, CurSubPath As String
    
    On Error GoTo DirSizeErr
    
    ' Calcule la taille du fichier ?
    If GetAttr(Directory) <> vbDirectory Then
        DirSize = FileLen(Directory)
        Exit Function
    End If
    
    ' V�rifit le r�pertoire
    If Right(Directory, 1) <> "\" Then Directory = Directory & "\"
    If Dir(Directory, vbDirectory) = "" Then
        MsgBox "Le r�pertoire est invalide.", vbExclamation
        DirSize = False
        Exit Function
    End If
    
    ' Calcule la taille du r�pertoire
    f = Dir(Directory, Attrib)
    If f = "" Then
        MsgBox "Erreur: R�pertoire invalide", vbExclamation
        DirSize = False
        Exit Function
    End If
    Do While True
        ' Donne le fichier
        If f = "." Or f = ".." Then GoTo NextDo1
        ' Fin
        If f = "" And CurSubPath = "" Then
            Exit Do
        ElseIf f = "" Then
            ' Se place au r�p. d'avant et au fichier suivant
            For i% = Len(CurSubPath) - 1 To 2 Step -1
                If Mid(CurSubPath, i%, 1) = "\" Then GoTo RepFound1
            Next
            i% = 0
RepFound1:
            PathToFind$ = Mid(CurSubPath, i% + 1, Len(CurSubPath) - i% - 1)
            CurSubPath = Left(CurSubPath, i%)
            f = Dir(Directory & CurSubPath, Attrib)
            f = Dir
            Do
                f = Dir
            Loop Until f = PathToFind$ And GetAttr(Directory & CurSubPath & f) = vbDirectory
        Else
            ' R�pertoire
            If GetAttr(Directory & CurSubPath & f) = vbDirectory Then
                ' Se place ds ce r�p.
                CurSubPath = CurSubPath & f & "\"
                f = Dir(Directory & CurSubPath, Attrib)
            ' Fichier
            Else
                ' Ajoute � la taille totale
                DirSize = DirSize + FileLen(Directory & CurSubPath & f)
            End If
        End If
NextDo1:
        DoEvents
        ' Donne le fichier (suite)
        f = Dir
    Loop

    Exit Function
DirSizeErr:
    If Err Then MsgBox "Erreur " & Err & ": " & Error(Err) & ".", vbExclamation
    DirSize = False
End Function

Public Function XCopy(Source As String, Destination As String, Optional AttribToCopy As Integer = vbArchive Or vbHidden Or vbDirectory) As Boolean
    Dim f As String, TtLen As Long, CopiedLen As Long, CurSubPath As String
    Dim OriAttrib As Integer
    
    On Error GoTo XCopyErr
    
    ' V�rifit le r�pertoire de destination
    If Right(Destination, 1) <> "\" Then Destination = Destination & "\"
    If Dir(Left(Destination, 3), AttribToCopy) = "" Then
        MsgBox "Le r�pertoire d'installation est invalide.", vbExclamation
        XCopy = False
        Exit Function
    End If
    For i% = 4 To Len(Destination)
        i% = InStr(i%, Destination, "\", vbTextCompare)
        If i% = 0 Then Exit For
        If Dir(Left(Destination, i%), vbDirectory) = "" Then MkDir Left(Destination, i%)
    Next
    
    ' Calcule la taille du fichier ?
    If GetAttr(Source) <> vbDirectory Then
        FileCopy Source, Destination
        XCopy = True
        Exit Function
    End If
    
    ' V�rifit les r�pertoires
    If Right(Source, 1) <> "\" Then Source = Source & "\"
    If Dir(Source, AttribToCopy) = "" Then
        MsgBox "Le r�pertoire source est invalide.", vbExclamation
        XCopy = False
        Exit Function
    End If
    
    ' Calcule la taille a copier
    TtLen = DirSize(Source, AttribToCopy)
    If TtLen = -1 Then
        XCopy = False
        Exit Function
    End If
    
    ' Copie
    CurSubPath = ""
    f = Dir(Source, AttribToCopy)
    If f = "" Then
        MsgBox "Erreur: R�pertoire invalide", vbExclamation
        XCopy = True
        Exit Function
    End If
    Do While True
        ' Donne le fichier
        If f = "." Or f = ".." Then GoTo NextDo2
        ' Fin
        If f = "" And CurSubPath = "" Then
            Exit Do
        ElseIf f = "" Then
            ' Se place au r�p. d'avant et au fichier suivant
            For i% = Len(CurSubPath) - 1 To 2 Step -1
                If Mid(CurSubPath, i%, 1) = "\" Then GoTo RepFound2
            Next
            i% = 0
RepFound2:
            PathToFind$ = Mid(CurSubPath, i% + 1, Len(CurSubPath) - i% - 1)
            CurSubPath = Left(CurSubPath, i%)
            f = Dir(Source & CurSubPath, AttribToCopy)
            f = Dir
            Do
                f = Dir
            Loop Until f = PathToFind$ And GetAttr(Source & CurSubPath & f) = vbDirectory
        Else
            ' R�pertoire
            If GetAttr(Source & CurSubPath & f) = vbDirectory Then
                ' Cr�� le r�p. dans la destination
                If Dir(Destination & CurSubPath & f, vbDirectory) = "" Then MkDir Destination & CurSubPath & f
                ' Se place ds ce r�p.
                CurSubPath = CurSubPath & f & "\"
                f = Dir(Source & CurSubPath, AttribToCopy)
            ' Fichier
            Else
                ' Copie
                If Len(FileDateTime(Destination & CurSubPath & f)) = 0 Then
                Else
                    If FileDateTime(Source & CurSubPath & f) < FileDateTime(Destination & CurSubPath & f) Then
                        GoTo CopyEnd
                    Else
                        SetAttr Destination & CurSubPath & f, vbNormal
                        Kill Destination & CurSubPath & f
                    End If
                End If
                OriAttrib = GetAttr(Source & CurSubPath & f)
                FileCopy Source & CurSubPath & f, Destination & CurSubPath & f
                SetAttr Destination & CurSubPath & f, OriAttrib And (Not vbReadOnly)
CopyEnd:
            End If
        End If
NextDo2:
        DoEvents
        ' Donne le fichier (suite)
        f = Dir
    Loop
    
    XCopy = True
    Exit Function
XCopyErr:
    If Err = 53 Then Resume Next
    If Err Then MsgBox "Erreur " & Err & ": " & Error(Err) & ".", vbExclamation
    XCopy = False
    Exit Function
End Function
