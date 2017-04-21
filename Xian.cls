VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Xian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' -----------------------------------------------------------
' (C) 2017 - Renaud Coustellier GNU Public License
' -----------------------------------------------------------
Public Function version() As String
    version = "DELF_DALF V1.00 Avril 2017"
End Function



Public Function PrepareFeuilleExamen(Optional examen As oDELF_DALF)
    '--- Cette fonction fabrique une feuille d'examen vierge
    '--- L'utilisateur doit indiquer :
    '--- o Le Niveau A1 A2 B1 B2 C1 C2
    '--- o Le nombre d'élèves
    '--- o le nombre de jurys
    '--- o les membres du jurys

    Dim oBook As New Excel.Workbook
    Dim oSheet As New Excel.Worksheet
    Dim oParam As New Excel.Worksheet

    Dim msg As String
    
    Set oBook = Application.ActiveWorkbook
    
    
    Dim tablo() As String
    
    '--- Ok, en avant
    Set oSheet = oBook.Sheets.Add()
    oSheet.Name = examen.Niveau
    
    '--- Taille des colonnes
    Columns("A:A").ColumnWidth = 4
    Columns("B:B").ColumnWidth = 28.86
    Columns("C:C").ColumnWidth = 13.43
    
    Columns("D:D").ColumnWidth = 17.43
    Columns("E:G").ColumnWidth = 14.86
    Columns("H:H").ColumnWidth = 8.71
    
    Columns("I:I").ColumnWidth = 19.86
    Columns("J:J").ColumnWidth = 8.71
    
    Dim lastColumn As String
    lastColumn = Chr(Asc("O") + examen.nbJurys - 5)
    
    
    Columns("K:" & lastColumn).ColumnWidth = 17.43
    
    oSheet.Cells(2, 4) = "EPREUVES ORALES INDIVIDUELLES " + examen.Niveau + Repo + " - " + Format(examen.DateExamen, "dddd mmm yy")
    oSheet.Cells(4, 2) = "Coordinateur : " + examen.Coordinateur
    
    oSheet.Cells(4, 9) = Str(Trim(examen.NbCandidats)) + " candidats"

    Rows("7:7").RowHeight = 30.75
    
    '--- Création des drop down jurys
    Dim range1 As Range, rng As Range
    
    Dim i As Integer
    
    For i = 1 To examen.nbJurys
        Lettre = Chr(74 + i)
        Set rng = oSheet.Range(Lettre + "4")
        
        With rng.Validation
            .Delete 'delete previous validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AdminJury"
        End With
        If examen.MembreJury(i, 1) <> "" Then
            oSheet.Cells(4, 10 + i) = examen.MembreJury(i, 1)
        End If
        
        Set rng = oSheet.Range(Lettre + "5")
        With rng.Validation
            .Delete 'delete previous validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AdminJury"
        End With
        If examen.MembreJury(i, 2) <> "" Then
            oSheet.Cells(5, 10 + i) = examen.MembreJury(i, 2)
        End If
        oSheet.Cells(6, 10 + i) = "Jury" + Str(i)
    Next i
    
    '--- Création des titres
    oSheet.Cells(6, 1) = "N°"
    oSheet.Cells(6, 2) = "Numéro Candidat"
    oSheet.Cells(6, 3) = "Sexe"
    oSheet.Cells(6, 4) = "Nom"
    oSheet.Cells(6, 5) = "Prénom"
    oSheet.Cells(6, 6) = "Hanzi"
    oSheet.Cells(6, 7) = "Préparation"
    oSheet.Cells(6, 9) = "Passation"
    oSheet.Cells(7, 7) = "Horaires"
    oSheet.Cells(7, 9) = "Horaires"
    oSheet.Cells(7, 8) = "Durée en mn"
    oSheet.Cells(7, 10) = "Durée en mn"
    
    '--- Merge des cellules
    For i = 1 To 8
        Select Case i
            Case 1
                sr = "A6:A7"
            Case 2
                sr = "B6:B7"
            Case 3
                sr = "C6:C7"
            Case 4
                sr = "D6:D7"
            Case 5
                sr = "E6:E7"
            Case 6
                sr = "F6:F7"
            Case 7
                sr = "G6:H6"
            Case 8
                sr = "I6:J6"
        End Select
        
        Set rng = oSheet.Range(sr)
        rng.Select
        With Selection
            .MergeCells = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
    Next i
    
    '--- Le cadre pour les titres
    cadreFin ("A6:" & lastColumn & "7")
    
    Set rng = oSheet.Range("G7:J7")
    rng.Select

    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    
    Dim myParam As New oParam
    
    myParam.LectureNiveauTarif
    DureePrepa = myParam.GetParam(examen.Niveau, "Durées préparation")
    'DureePassation = myParam.GetParam(examen.Niveau, "Durées passation")
    
    Dim horaireDebut As String
    
    horaireDebut = myParam.GetParam(examen.Niveau, "Horaire examens")
    tablo() = Split(horaireDebut, "-")
    horaireDebut = tablo(0)
    horaireDebut = Replace(horaireDebut, "h", ":")
    If Right(horaireDebut, 1) = ":" Then
        horaireDebut = horaireDebut + "00"
    End If
    
    horaireDebut = Format(CDate(horaireDebut), "HH:MM") + " - " + Format(CDate(horaireDebut) + 1 / 1440 * Val(DureePrepa), "HH:MM")

    
    oSheet.Cells(8, 7) = horaireDebut
    
    If examen.NbCandidats \ examen.nbJurys <> examen.NbCandidats / examen.nbJurys Then
        nbCandidatParJury = examen.NbCandidats \ examen.nbJurys + 1
    Else
        nbCandidatParJury = examen.NbCandidats / examen.nbJurys
    End If
    
    
    '--- numéros candidats
    For i = 1 To examen.NbCandidats
        oSheet.Cells(7 + i, 1) = Str(i)
        oSheet.Cells(7 + i, 8).Formula = "=DureePreparation()"
        oSheet.Cells(7 + i, 10).Formula = "=DureePassation()"
        oSheet.Cells(7 + i, 9).Formula = "=heurePassation(G" & Trim(Str(i + 7)) & ")"
        If i > 1 And i + 7 - (examen.nbJurys) > 7 Then
            oSheet.Cells(7 + i, 7).Formula = "=HeurePreparation(G" & Trim(Str(i + 7 - (examen.nbJurys))) & ")"
        End If
    Next i


    '---
    If examen.nbJurys > 1 Then
        '--- Fusion des cellules par nbJurys
        Application.DisplayAlerts = False
        For i = 1 To examen.NbCandidats Step examen.nbJurys
            For j = 1 To examen.nbJurys
                ' ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-7],RC[-6])
                '--- Horaires
                sr = "G" + Trim(i + 7) + ":G" + Trim(i + 7 + examen.nbJurys - 1)
                Set rng = oSheet.Range(sr)
                rng.Select
                With Selection
                    .MergeCells = True
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .WrapText = True
                End With
                
                  sr = "H" + Trim(i + 7) + ":H" + Trim(i + 7 + examen.nbJurys - 1)
                Set rng = oSheet.Range(sr)
                rng.Select
                With Selection
                    .MergeCells = True
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .WrapText = True
                End With
                
                sr = "I" + Trim(i + 7) + ":I" + Trim(i + 7 + examen.nbJurys - 1)
                Set rng = oSheet.Range(sr)
                rng.Select
                With Selection
                    .MergeCells = True
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .WrapText = True
                End With
                
                  sr = "J" + Trim(i + 7) + ":J" + Trim(i + 7 + examen.nbJurys - 1)
                Set rng = oSheet.Range(sr)
                rng.Select
                With Selection
                    .MergeCells = True
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .WrapText = True
                End With
                
                '--- Candidats
                sr = Chr(74 + j) + Trim(i + 7) + ":" + Chr(74 + j) + Trim(i + 7 + examen.nbJurys - 1)
                Set rng = oSheet.Range(sr)
                rng.Select
                With Selection
                    .MergeCells = True
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .WrapText = True
                End With
                '--- Concataine
                'oSheet.Cells(i + 7, j + 10).Formula = "=D" & Trim(Str(i + 6 + j)) & " & ""-"" & E" + Trim(Str(i + 6 + j))
                oSheet.Cells(i + 7, j + 10).Formula = "=CONCATENATE(INDIRECT(""D" & Trim(Str(i + 6 + j)) & """)," & """-""" & ",INDIRECT(""E" + Trim(Str(i + 6 + j)) + """" & "))"
                
            Next j
            '--- Ajoute la couleur
            If i Mod 2 = 1 Then ' une fois sur deux
                sr = "A" + Trim(Str(i + 7)) + ":" & lastColumn & Trim(Str(i + 6 + examen.nbJurys))
                Range(sr).Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent4
                    .TintAndShade = 0.599993896298105
                    .PatternTintAndShade = 0
                End With
            End If
        Next i
        
    Else
        '--- Todo cas ou il n'y a qu'un jury !
        For i = 1 To examen.NbCandidats
            oSheet.Cells(i + 7, 11).Formula = "=D" & Trim(Str(i + 7)) & " & ""-"" & E" + Trim(Str(i + 7))
            '--- Ajoute la couleur
            If i Mod 2 = 1 Then ' une fois sur deux
                sr = "A" + Trim(Str(i + 7)) + ":O" + Trim(Str(i + 7))
                Range(sr).Select
                With Selection.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorAccent4
                    .TintAndShade = 0.599993896298105
                    .PatternTintAndShade = 0
                End With
            End If
        Next i
            
        
    End If
    
    Application.DisplayAlerts = True
    
    
    '--- Concataine les jurys
    For i = 1 To examen.nbJurys
        Letter = ConvertToLetter(i + 10)
        oSheet.Cells(7, 10 + i).Formula = "=" & Letter & Trim(Str(4)) & " & ""  / "" & " & Letter & Trim(Str(5))
    Next i
    
        
    '--- Le cadre pour le tableau
    sr = "A8:" & lastColumn + Trim(Str(7 + examen.NbCandidats))
    cadreFin (sr)
    '--- Centrer
    sr = "G6:" & lastColumn + Trim(Str(7 + examen.NbCandidats))
    Set rng = oSheet.Range(sr)
    rng.Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    '--- Supprimer les colonnes en trop
    'sr = ""
    'If examen.nbJurys = 1 Then
    '    sr = "L:O"
    'ElseIf examen.nbJurys = 2 Then
    '    sr = "M:O"
    'ElseIf examen.nbJurys = 3 Then
    '    sr = "N:O"
    ''ElseIf examen.nbJurys = 4 Then
    '    sr = "O:O"
    'End If
    '
    'If sr <> "" Then
    '    Set rng = oSheet.Range(sr)
    '    rng.Select
    '    Selection.Delete Shift:=xlToLeft
    'End If
    'On Error Resume Next
    'Columns("N:O").Select
    'Selection.Delete
    'On Error GoTo 0
    '---------------------------------------------------------
    '------------------------------ Liste des absents
    '---------------------------------------------------------
    
    aPos = 12 + examen.NbCandidats
    
    oSheet.Cells(aPos, 2) = "Liste des absents"
    oSheet.Cells(aPos + 1, 2) = "Commentaires"
    oSheet.Cells(aPos + 1, 3) = "N° Candidat"
    oSheet.Cells(aPos + 1, 4) = "Sexe"
    oSheet.Cells(aPos + 1, 5) = "Nom"
    oSheet.Cells(aPos + 1, 6) = "Prénom"
    oSheet.Cells(aPos + 1, 7) = "Hanzi"
    
    'sR = Trim(Str(aPos + 1)) + ":" + Trim(Str(aPos + 1))
    
    aMax = aPos + 2 + Int(examen.NbCandidats / 2) + 1
    
    '--- Le cadre pour le tableau
    sr = "B" + Trim(Str(aPos)) + ":G" + Trim(Str(aMax))
    cadreFin (sr)
    '--- Merger
    sr = "B" + Trim(Str(aPos)) + ":G" + Trim(Str(aPos))
    Set rng = oSheet.Range(sr)
    rng.Select
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    '--- Centrer
    sr = "B" + Trim(Str(aPos + 1)) + ":G" + Trim(Str(aPos + 1))
    Set rng = oSheet.Range(sr)
    rng.Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    '--- Feuille de rémunération des épreuves
    
    '---------------------------------------------------------
    '------------------------------ Rémunérations Surveillance
    '---------------------------------------------------------
    
    rpos = aMax + 3
    oSheet.Cells(rpos, 2) = "Rémunérations surveillance"
    oSheet.Cells(rpos, 3) = "Personnes"
    oSheet.Cells(rpos, 4) = "Surveillance"
    
    rmax = rpos + 10 '+ examen.nbJurys * 2 + 5 + 1
    
    '--- Le cadre pour le tableau
    sr = "B" + Trim(Str(rpos)) + ":D" + Trim(Str(rmax))
    cadreFin (sr)
    
    '--- On pose la liste des fonctions administratives / colonne totaux
    For i = 1 To 10
        Set rng = oSheet.Range("B" + Trim(Str(rpos + i)))
        With rng.Validation
            .Delete 'delete previous validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AdminTaches"
        End With
        Set rng = oSheet.Range("C" + Trim(Str(rpos + i)))
        With rng.Validation
            .Delete 'delete previous validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AdminPerso"
        End With
        
        '--- Petite formule
        Formula = "=if(C" & Trim(Str(rpos + i)) & " <> 0,CoutSurveillance(""" + examen.Niveau + """,B" & Trim(Str(rpos + i)) & "),"""")"
        oSheet.Cells(rpos + i, 4).Formula = Formula
    Next i
    
    
    '--- Total du tableau surveillabnce
    sr = "D" + Trim(Str(rpos)) + ":D" + Trim(Str(rmax))
    
    oSheet.Cells(rpos + i, 4).Formula = "=sum(" + sr + ")"
    oSheet.Cells(rpos + i, 2).Formula = "Total"
    
    'Dim rng As Range
    Set rng = Range("D" & Trim(Str(rpos + i)))
    rng.Name = "Total" & oSheet.Name & "Surveillance"
    
    
    sr = "B" + Trim(Str(rmax + 1)) + ":D" + Trim(Str(rmax + 1))
    cadreFin (sr)
    
    Range(sr).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    
    
    
    '----------------------------------------------------------
    '--- Gestion des jurys
    '----------------------------------------------------------
    previ = rmax + 3
    
    oSheet.Cells(previ, 4) = "Epreuve Individuelle"
    oSheet.Cells(previ, 6) = "Epreuve Collective"
    oSheet.Cells(previ, 8) = "Total Ep. Ind."
    oSheet.Cells(previ, 9) = "Total Ep. Col."
    oSheet.Cells(previ, 10) = "Total"
    
    previ = previ + 1
    oSheet.Cells(previ, 4) = "Nombre Candidats"
    oSheet.Cells(previ, 5) = "Nombre Sujets"
    oSheet.Cells(previ, 6) = "Nombre copies"
    oSheet.Cells(previ, 7) = "Sujets"
    
    
    '--- Le cadre pour le tableau
    sr = "D" + Trim(Str(previ - 1)) + ":j" + Trim(Str(previ))
    cadreFin (sr)
    
    For i = 1 To examen.nbJurys * 2 Step 2 ' to 10
        oSheet.Cells(previ + i, 2) = "Jury" + Str(((i + 1) \ 2))
        oSheet.Cells(previ + i + 1, 2) = "Jury" + Str(((i + 1) \ 2))
    Next i
    
    For i = 1 To examen.nbJurys * 2
        '--- Création de la liste déroulante jury
        Set rng = oSheet.Range("C" + Trim(Str(previ + i)))
        With rng.Validation
            .Delete 'delete previous validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AdminJury"
        End With
        If (i + 1) Mod 2 = 0 Then
            Letter = ConvertToLetter(Int(i + 1) / 2 + 10)
            oSheet.Cells(previ + i, 3).Formula = "=if(" & Letter & "4<>0," & Letter & "4,"" "")"
            oSheet.Cells(previ + i + 1, 3).Formula = "=if(" & Letter & "5<>0," & Letter & "5,"" "")"
        End If
                
        '--- petit formule
        'CoutEpreuveIndividuelle("C1";C53;D53;E53)
        Formula = "=if(C" & Trim(Str(previ + i)) & " <> 0,CoutEpreuveIndividuelle(""" + examen.Niveau + """,C"
        Formula = Formula & Trim(Str(previ + i)) & ",D" & Trim(Str(previ + i))
        Formula = Formula & ",E" & Trim(Str(previ + i)) & "),"""")"
        oSheet.Cells(previ + i, 8).Formula = Formula
                
        '--- petit formule
        'CoutEpreuveCollective("C1";C53;D53;E53)
        Formula = "=if(C" & Trim(Str(previ + i)) & " <> 0,CoutEpreuveCollective(""" + examen.Niveau + """,C"
        Formula = Formula & Trim(Str(previ + i)) & ",F" & Trim(Str(previ + i))
        Formula = Formula & ",G" & Trim(Str(previ + i)) & "),"""")"
        oSheet.Cells(previ + i, 9).Formula = Formula
        
        '--- La somme
        oSheet.Cells(previ + i, 10).Formula = "=H" & Trim(Str(previ + i)) & " + I" & Trim(Str(previ + i))
                
    Next i
    
    '--- Le cadre pour le tableau
    sr = "B" + Trim(Str(previ + 1)) + ":j" + Trim(Str(previ + i - 1))
    cadreFin (sr)
    
    rmax = previ + i
        
    '--- Totaux bas du tableau
    oSheet.Cells(rmax, 2) = "Totaux Jurys"
    sr = "J" + Trim(Str(previ + 1)) + ":j" + Trim(Str(previ + i - 1))
    oSheet.Cells(rmax, 10).Formula = "=sum(" + sr + ")"
    
    Set rng = Range("J" & Trim(Str(rmax)))
    rng.Name = "Total" & oSheet.Name & "Jurys"
    
    With Selection.Font
        .Name = "Comic Sans MS"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
    '--- Cadre
    sr = "B" + Trim(Str(rmax)) + ":j" + Trim(Str(rmax))
    cadreFin (sr)
    Range(sr).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    
    '-------------------------------------------
    '--------- TOTAL Surveillance et Jurys
    '____________________________________________
    
    rmax = rmax + 2
    
     '--- Totaux bas du tableau
    oSheet.Cells(rmax, 2) = "Total Examen (Surveillance + Jury)"
    oSheet.Cells(rmax, 4) = "=Total" & oSheet.Name & "Surveillance + Total" & oSheet.Name & "Jurys"
    
    'Dim rng As Range
    Set rng = Range("D" & Trim(Str(rmax)))
    rng.Name = "Total" & oSheet.Name & "SurveillanceJury"
    
    
    sr = "B" + Trim(Str(rmax)) + ":D" + Trim(Str(rmax))
    '--- Cadre
    '
    cadreFin (sr)
    Range(sr).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Name = "Comic Sans MS"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
    oSheet.Calculate
    oSheet.Visible = xlSheetVisible
    'oSheet.Cells(rMax + 1, 4) = "=sum(D" + Trim(Str(rPos + 2)) + ":D" + Trim(Str(rMax)) + ")"
    'oSheet.Cells(rMax + 1, 5) = "=sum(E" + Trim(Str(rPos + 2)) + ":E" + Trim(Str(rMax)) + ")"
    'oSheet.Cells(rMax + 1, 6) = "=sum(F" + Trim(Str(rPos + 2)) + ":F" + Trim(Str(rMax)) + ")"
    
    '--- Ligne des totaux
    MsgBox "Génération de la feuille terminée !", vbInformation, "Préparation de la feuille d'examen"

End Function

Private Function cadreFin(sRange As String)
    Dim oSheet As New Excel.Worksheet
    Dim rng As Range
    
    Set oSheet = Application.ActiveSheet
    Set rng = oSheet.Range(sRange)
    
    rng.Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    
End Function




Public Function CalculerCoutPrevisionnel()
    '----------------------------------------------------------------------------------------------------------------------------
    '--- Cette fonction calcule le cout prévisionnel de l'examen en utilisant :
    ' - la feuille d'examen en cours A1 ou A2... C2
    ' - la feuille de paramètres
    '-----------------------------------------------------------------------------------------------------------------------------
    
    Dim oBook As New Excel.Workbook
    Dim oSheet As New Excel.Worksheet
    Dim oParam As New Excel.Worksheet
    
    Repo = MsgBox("Vous allez générer les calcul de coût prévisionnel de l'examen." & vbCrLf & vbCrLf & "Continuer ?", vbQuestion + vbYesNo, "Calcul coût prévisionnel !")
    
    If Repo <> 6 Then
        Exit Function
    End If
    
    DoEvents
    
    Set oBook = Application.Workbooks(1)
    Set oSheet = Application.ActiveSheet
    
    Select Case Trim(oSheet.Name)
    Case "A1", "A2", "B1", "B2", "C1", "C2"
        '--- Trouvé !
        lbFlag = True
    Case Else
        lbFlag = False
    End Select
    
    If lbFlag = False Then
        MsgBox "La feuille active n'est pas une feuille d'examen (A1 ou A2 ou ...C1) !", vbInformation, "Calcul prévionnel"
        Exit Function
    End If
    
    '--- Recherche des param-tres
    Set oParam = oBook.Sheets("Parametres")
    
    Dim Admin As New Dictionary
    Dim jury As New Dictionary
    
    '--- Selon le type, on va chercher la bonne ligne "horaires" dans la feuille paramètres
     
    Set oParam = oBook.Sheets("Parametres")
    If oSheet.Name = "A1" Then
        pPos = 2
    ElseIf oSheet.Name = "A2" Then
        pPos = 3
    ElseIf oSheet.Name = "B1" Then
        pPos = 4
    ElseIf oSheet.Name = "B2" Then
        pPos = 5
    ElseIf oSheet.Name = "C1" Then
        pPos = 6
    ElseIf oSheet.Name = "C2" Then
        pPos = 7
    End If

    '--- Création du dictionnaire tarif administratif
    For i = 13 To 20
        sIndex = oParam.Cells(1, i)
        sVal = oParam.Cells(pPos, i)
        Admin.Add sIndex, sVal
    Next i
    
    '--- Création du tarif jury
    For i = 22 To 25
        sIndex = oParam.Cells(1, i)
        sVal = oParam.Cells(pPos, i)
        jury.Add sIndex, sVal
    Next i
    
    '--- Recherche du tableau des rémunérations
    
    posTab = 0
    For i = 20 To 60
        If Left(oSheet.Cells(i, 2), 12) = "Rémunération" Then
            posTab = i
            Exit For
        
        End If
    Next i
       
    If posTab = 0 Then
        MsgBox "Impossible de trouver le tableau des rémunérations entre les cellules 20 et 60 contenant la chaine ", vbInformation, "Calcul prévisionnel"
        Exit Function
    End If
 
    '---
    Dim tablo() As String
    Dim j As Integer
    For i = posTab + 1 To posTab + 30
        If oSheet.Cells(i, 2) = "Totaux" Then ' Or oSheet.Cells(i, 2) = "" Then
            Exit For
        End If
        sKey = oSheet.Cells(i, 2)
    
        If Admin.Exists(sKey) Then
            oSheet.Cells(i, 4) = Admin.Item(sKey)
        ElseIf UCase(Left(sKey, 4)) = "JURY" Then
            If Len(oSheet.Cells(i, 3)) > 1 Then
                '--- IL y a un jury :
                '--- Recherche nombre candidats
                nbCandidat = 0
                j = Val(Mid(sKey, 5)) + 10
                
                For k = 8 To 50
                    'On Error Resume Next
                    If Len(oSheet.Cells(k, j)) > 5 Then
                        
                        nbCandidat = nbCandidat + 1
                        'tablo() = Split(oSheet.Cells(k, 9), "-")
                        'HeureFin = tablo(1)
                    'Else
                    '    Exit For
                    End If
                Next k
                '--- Calculs en fonction du nombre de candidats
                If InStr(UCase(sKey), "COMPL") = 0 Then
                    oSheet.Cells(i, 4).Formula = "=" & jury.Items(0) & " + " & jury.Items(1) & " * " & nbCandidat
                    '--- Est ce que l'on a saisi un  nombre de copies ?
                    If oSheet.Cells(i, 7) = "" Then
                        oSheet.Cells(i, 5).Formula = "=" & jury.Items(2) & " + " & jury.Items(3) & " * " & nbCandidat
                    Else
                        oSheet.Cells(i, 5).Formula = "=" & jury.Items(2) & " + " & jury.Items(3) & " * G" & Trim(Str(i))
                    End If
                Else
                    '--- Est ce que l'on a saisi un  nombre de copies ?
                    If oSheet.Cells(i, 7) <> "" Then
                        oSheet.Cells(i, 5).Formula = "=" & jury.Items(2) & " + " & jury.Items(3) & " * G" & Trim(Str(i))
                    End If
                End If
                
                
                
                
                oSheet.Cells(i, 6).Formula = "=if(D" + Trim(Str(i)) + " + E" + Trim(Str(i)) + "<>0,D" + Trim(Str(i)) + " + E" + Trim(Str(i)) + ","" "")"
                
            End If
        Else
            '--- C'est un ligne vide mais on ne sait pas ce que c'est
             'oSheet.Cells(i, 4) = "erreur"
        End If
        
        
    Next i
 
 
    MsgBox "Calcul terminé pour cette feuille !", vbInformation, "Calcul coût prévisionnel"
    
 
    Set oBook = Nothing
    Set oSheet = Nothing
    Set oParam = Nothing
    
End Function
Public Function GenererPaye(MaPaye As oPaye)
    '----------------------------------------------------------------------------------------------------------------------------
    '--- Cette fonction génère le classeur des Bulletins de Paye, enfin, a priori
    '-----------------------------------------------------------------------------------------------------------------------------
    Dim oBook As New Excel.Workbook
    Dim oSheet As New Excel.Worksheet
    Dim oBookModele As New Excel.Workbook
    Dim oBookPaye As New Excel.Workbook
    
    Dim j As Integer
    
    Dim startPaye As Integer
    Dim endPaye As Integer
    
    Dim i As Integer
    Dim montant As Integer
    
    
    
    Dim cherche As String
    Dim p As Integer
    Dim Personne As String
    Dim lastPos  As Integer
    
    Dim aPayer As Boolean   '--- Personne à payer ou pas, lecture dans la colonne E "non payé" de la table de paramètres
    
    Dim nbSujetsLus As Integer
    
    Dim myParamPaye As New oParam
    
    myParamPaye.LectureNiveauTarif
    
    Dim fso As New FileSystemObject
    Dim FilePath As String
    FilePath = Application.ActiveWorkbook.path
    
    Set oBook = Application.ActiveWorkbook
    For p = 1 To MaPaye.Personnes.count
        '--- Pour chaque personne
        Personne = MaPaye.Personnes.Items(p - 1)
        lastPos = 12
        
        
        For i = 1 To MaPaye.examen.count
            '--- Pour chaque examen
            
            Set oSheet = oBook.Worksheets(MaPaye.examen.Items(i - 1))
            
            For j = 10 To 250
                '--- Recherche du début de la paye
                cherche = oSheet.Cells(j, 2)
                If Left(cherche, 12) = "Rémunération" Then
                    '--- tableau des rémunérations trouvé !!!
                    startPaye = j + 1
                    Exit For
                End If
            Next j
            
            For j = startPaye + 1 To 300
                '--- Recherche de la fin de la paye
                cherche = oSheet.Cells(j, 2)
                If Left(cherche, 6) = "Totaux" Then
                    '--- tableau des rémunérations trouvé !!!
                    endPaye = j - 1
                    Exit For
                End If
            Next j
            
            
            nbSujetsLus = 1
            Dim typeRemuneration As Integer ' 1 = Administratif / 2 = jury
            
            
            If startPaye <> 0 And endPaye <> 0 Then
                '--- On va chercher dans la zone trouvée les personnes de MyPaye.Personne
                For j = startPaye To endPaye
                    cherche = oSheet.Cells(j, 3)
                    If cherche = Personne Then
                    
                        aPayer = myParamPaye.PersonneAPayer(Personne)
                    
                        '--- On a trouvé la personne à payer
                        If Left(UCase(oSheet.Cells(j, 2)), 4) = "JURY" Then
                            typeRemuneration = 2    '--- jury
                        Else
                            typeRemuneration = 1    '--- admin
                        End If
                        
                        montant = 0
                        
                        If typeRemuneration = 1 Then
                            montant = myParamPaye.GetParam(oSheet.Name, (oSheet.Cells(j, 2)))
                            'oSheet.Cells(j, 4) = montant
                            oSheet.Cells(j, 4).Formula = "=if(C" & Trim(Str(j)) & " <> 0,CoutSurveillance(""" + oSheet.Name + """,B" & Trim(Str(j)) & "),"""")"
                            'If Val(oSheet.Cells(j, 4)) <> 0 Then
                            '    montant = Val(oSheet.Cells(j, 4))
                            'End If
                        ElseIf typeRemuneration = 2 Then
                            montant = Val(oSheet.Cells(j, 5)) + Val(oSheet.Cells(j, 6)) + Val(oSheet.Cells(j, 7)) + Val(oSheet.Cells(j, 8))
                        End If
                        
                        If montant <> 0 And aPayer = True Then
                            '--- On a trouvé des montants !
                            '--- Ouvre le modele qui est dans le dossier courant et qui se nomme BulletinPayeModele
                            If p = 1 And lastPos = 12 Then
                                '--- c'est la première personne et la première ligne
                                ' on ouvre le classeur modèle pour la première copie
                                On Error Resume Next
                                Set oBookModele = Application.Workbooks.Open(Application.ActiveWorkbook.path + "\BulletinPayeModele.xlsx")
                                If Err.Number <> 0 Then
                                    MsgBox "Impossible de trouver le document 'BulletinPayeModele.xlsx' dans le dossier courant, la génération est annulée.", vbInformation, "Génération Bulletin Paye"
                                    Exit Function
                                End If
        
                                On Error GoTo 0
                                '--- Sélectionne la feuille de modele
                                Set oSheetCopy = oBookModele.Sheets("Modèle")
                                '--- Ouvre un nouveau classeur en mémoire
                                If p < 2 Then
                                    Set oBookPaye = Application.Workbooks.Add
                                End If
                                
                                Dim oSheetPaye As New Excel.Worksheet
                
                                oSheetCopy.Copy oBookPaye.Sheets(Sheets.count)  '--- Copie
                                'oBookConvocations.Sheets(2).Delete
                                '--- Remplissage de la feuille
                                Set oSheetPaye = oBookPaye.Sheets("Modèle")
                                'oSheetJury.Cells(8, 2) = "CONVOCATION AUX " + oConvoc.Cells(j, 3)
                                'oSheetJury.Cells(9, 4) = oConvoc.Cells(j, 4)
                                oSheetPaye.Name = Personne
                                'oSheetPaye.Cells(4, 1) = personne
                                
                            ElseIf lastPos = 12 Then
                                '--- On change de personne
                                oSheetCopy.Copy oBookPaye.Sheets(Sheets.count)  '--- Copie
                                Set oSheetPaye = oBookPaye.Sheets("Modèle")
                                oSheetPaye.Name = Personne
                            End If
                            oSheetPaye.Cells(4, 4) = Personne
                            oSheetPaye.Cells(lastPos, 3) = MaPaye.examen.Items(i - 1) 'examen
                            
                            oSheetPaye.Cells(i + 2, 5) = oSheet.Cells(2, 4)
                            
                            If typeRemuneration = 2 Then
                                '---------------------
                                '--- C'est un JURY
                                '---------------------
                                EiCandidat = Val(oSheet.Cells(j, 4))
                                EiSujet = Val(oSheet.Cells(j, 5))
                                EcCopie = Val(oSheet.Cells(j, 6))
                                EcSujet = Val(oSheet.Cells(j, 7))
                                
                                '--- C'est un Jury
                                If EiCandidat <> 0 Then
                                    '--- Ligne Oral
                                    ' Lecture sujet   Passation oraux
                                    oSheetPaye.Cells(lastPos, 4) = "Jury Epreuve Individuelle"
                                    oSheetPaye.Cells(lastPos, 5) = EiCandidat
                                    oSheetPaye.Cells(lastPos, 6) = EiSujet
                                    '--------
                                    oSheetPaye.Cells(lastPos, 7).Formula = "=" & Val(myParamPaye.GetParam(oSheet.Name, "Lecture sujet")) & "*" & EiSujet & " + " & Val(myParamPaye.GetParam(oSheet.Name, "Passation oraux")) & " * " & EiCandidat
                                    
                                    Formula = "=if(C" & Trim(Str(j)) & " <> 0,CoutEpreuveIndividuelle(""" + oSheet.Name + """,C"
                                    Formula = Formula & Trim(Str(j)) & ",D" & Trim(Str(j))
                                    Formula = Formula & ",E" & Trim(Str(j)) & "),"""")"
                                    
                                    oSheet.Cells(j, 8).Formula = Formula '"=" & Val(myParamPaye.GetParam(oSheet.Name, "Lecture sujet")) & "*" & EiSujet & " + " & Val(myParamPaye.GetParam(oSheet.Name, "Passation oraux")) & " * " & EiCandidat
                                    
                                    oSheet.Cells(j, 10).Formula = "=I" & Trim(Str(j)) & "+H" & Trim(Str(j))
                                    '--------
                                    'oSheetPaye.Cells(lastPos, 7).Formula = "=if(C" & Trim(Str(j)) & " <> 0,CoutSurveillance(""" + examen.Niveau + """,B" & Trim(Str(j)) & "),"""")"
                                    lastPos = lastPos + 1
                                    
                                End If
                                If EcCopie <> 0 Then
                                    '--- Ligne écrit
                                    ' Connaissance sujet ép collective    Correction copies
                                    oSheetPaye.Cells(lastPos, 3) = MaPaye.examen.Items(i - 1) 'examen
                                    oSheetPaye.Cells(lastPos, 4) = "Jury Epreuve Collective"
                                    oSheetPaye.Cells(lastPos, 5) = EcCopie
                                    oSheetPaye.Cells(lastPos, 6) = EcSujet
                                    oSheetPaye.Cells(lastPos, 7).Formula = "=" & Val(myParamPaye.GetParam(oSheet.Name, "Connaissance sujet ép collective")) & " * " & EcSujet & " + " & Val(myParamPaye.GetParam(oSheet.Name, "Correction copies")) & " * " & EcCopie
                                    
                                    Formula = "=if(C" & Trim(Str(j)) & " <> 0,CoutEpreuveCollective(""" + oSheet.Name + """,C"
                                    Formula = Formula & Trim(Str(j)) & ",F" & Trim(Str(j))
                                    Formula = Formula & ",G" & Trim(Str(j)) & "),"""")"
                                    
                                    oSheet.Cells(j, 9).Formula = Formula '"=" & Val(myParamPaye.GetParam(oSheet.Name, "Connaissance sujet ép collective")) & " * " & EcSujet & " + " & Val(myParamPaye.GetParam(oSheet.Name, "Correction copies")) & " * " & EcCopie
                                    oSheet.Cells(j, 10).Formula = "=I" & Trim(Str(j)) & "+H" & Trim(Str(j))
                                    lastPos = lastPos + 1
                                End If
                                
                                
                            Else
                                '---------------------
                                '--- Tâche administrative
                                '---------------------
                                oSheetPaye.Cells(lastPos, 4) = oSheet.Cells(j, 2) 'sujet
                                oSheetPaye.Cells(lastPos, 7) = myParamPaye.GetParam(oSheet.Name, (oSheet.Cells(j, 2)))
                                lastPos = lastPos + 1
                            End If
                            
                            oSheetPaye.Cells(31, 1) = MaPaye.Message1
                            oSheetPaye.Cells(32, 1) = MaPaye.Message2
                        ElseIf montant <> 0 And aPayer = False Then
                            '--- On a corrigé des copie ou autre mais on est une personne qui n'est pas payée
                            oSheet.Cells(j, 8) = 0
                            oSheet.Cells(j, 9) = 0
                        End If
                    End If
                    
                Next j
            End If
        
        
        Next i
        
        
        If MaPaye.GeneratePDF = True Then
            '--- Export PDF
            Filename = FilePath & "\BP_" & Trim(Personne) & "pdf"
        
            If fso.FileExists(Filename) Then
                fso.DeleteFile (Filename)
            End If
            oSheetPaye.ExportAsFixedFormat xlTypePDF, Filename
        End If
        
    Next p
    
    On Error Resume Next
    oBookModele.Close
    On Error GoTo 0
    
    Set oBook = Nothing
    Set oSheet = Nothing
    Set oBookModele = Nothing
    Set oBookPaye = Nothing
    
    MsgBox "Opération terminée", vbInformation, "DALF DELF"
    
 
End Function

Public Function GenererConvocation(MyConvo As oConvocations)
    '----------------------------------------------------------------------------------------------------------------------------
    '--- Cette fonction génère le classeur des convocations par jury qui permet ensuite d'imprimer les convocations (2 par deux)
    '-----------------------------------------------------------------------------------------------------------------------------
    Dim oBook As New Excel.Workbook
    Dim oSheet As New Excel.Worksheet
    Dim oConvoc As New Excel.Worksheet
    Dim oParam As New Excel.Worksheet
    Dim oBookModele As New Excel.Workbook
    Dim oBookConvocations As New Excel.Workbook
 
    Repo = MsgBox("Vous allez générer le document des convocation pour les enseignants ?" & vbCrLf & vbCrLf & "Continuer ?", vbQuestion + vbYesNo, "Génération convocations !")
    ' MsgBox "Test" + Application.ActiveWorkbook.Path
    
    
    If Repo <> 6 Then
        Exit Function
    End If
    
    DoEvents
    
    Set oBook = Application.Workbooks(1)

    '--- Recherche feuille Convocations
    On Error Resume Next
    Set oConvoc = oBook.Sheets("Convocations")
    If Err.Number <> 0 Then
        MsgBox "Impossible de trouver la feuille Convocations. Vérifiez que vous n'avez qu'un seul classeur ouvert !", vbCritical, "Erreur Générer Convocation"
        Exit Function
    End If
    Err.Clear
    On Error GoTo 0
    
    Dim fso As New FileSystemObject
    Dim FilePath As String
    FilePath = Application.ActiveWorkbook.path
    
    '--- Nettoyage feuille Convocations
    If MyConvo.GenererLaFeuille = True Then
        oConvoc.Cells.Range("A2", "J100").Clear
    End If
    

    Dim lbFlag
    Dim okFlag As Boolean
    okFlag = False
        
    indexJury = 2
    
    For i = 1 To oBook.Sheets.count
        '--- Recherche des feuilles d'examen présentes dans le document en cours
        '--- Attention la recherche est stricte, pas d'espace dans les noms des documents !
        Set oSheet = oBook.Sheets(i)
        Select Case Trim(oSheet.Name)
        Case "A1", "A2", "B1", "B2", "C1", "C2"
            '--- Trouvé !
            If MyConvo.examen.Exists(oSheet.Name) Then
                okFlag = True
                lbFlag = True
            Else
                lbFlag = False
            End If
        Case Else
            lbFlag = False
        End Select
        
        Dim tablo() As String
        
        If lbFlag = True Then
            '--- On a trouvé une feuille d'examen
            '--- Génération de la feuille de convocation
            Epreuve = oSheet.Cells(2, 4)     ' Epreuve
            tablo() = Split(Epreuve, "-")
            Epreuve = tablo(0)
            LaDate = tablo(1)
            
            '--- Selon le type, on va chercher la bonne ligne "horaires" dans la feuille paramètres
            Set oParam = oBook.Sheets("Parametres")
            If oSheet.Name = "A1" Then
                pPos = 2
            ElseIf oSheet.Name = "A2" Then
                pPos = 3
            ElseIf oSheet.Name = "B1" Then
                pPos = 4
            ElseIf oSheet.Name = "B2" Then
                pPos = 5
            ElseIf oSheet.Name = "C1" Then
                pPos = 6
            ElseIf oSheet.Name = "C2" Then
                pPos = 7
            End If
            
            HoraireConvoc = oParam.Cells(pPos, 8)
            HoraireExam = oParam.Cells(pPos, 9)
            If InStr(HoraireExam, "-") Then
                tablo() = Split(HoraireExam, "-")
            Else
                MsgBox "Impossible de trouver l'horaire d'examen dans la feuille 'parametres' en cellule " + Str(pPos) + "-9 ! Vous devez corriger pour générer les convocations.", vbInformation
                Exit Function
            End If
            
            DebutEpreuve = tablo(0)
            DuréePrépa = oParam.Cells(pPos, 10)
            duréePassation = oParam.Cells(pPos, 11)
            
            
            For j = 11 To 16
                '--- Recherche des jurys
                Nom = oSheet.Cells(7, j)
                If Len(Nom) > 5 Then
                    
                    If InStr(Nom, "/") Then
                        tablo() = Split(Nom, "/")
                    Else
                        mess = "Impossible de trouver le séparateur de nom '/' dans la cellule  7-" + Str(j) + " de la feuille " + oSheet.Name + vbCrLf + vbCrLf
                        MsgBox mess + "Génération des convocations impossible !", vbInformation, "Génération Convocations"
                        Exit Function
                    End If
                    '--- Constructions de la feuille des convocations contenant la liste des jurys
                    oConvoc.Cells(indexJury, 1) = tablo(0) 'oSheet.Cells(7, j)    ' Nom
                    oConvoc.Cells(indexJury + 1, 1) = tablo(1)
                    oConvoc.Cells(indexJury, 3) = Epreuve    ' Epreuve
                    oConvoc.Cells(indexJury + 1, 3) = Epreuve
                    oConvoc.Cells(indexJury, 4) = LaDate    ' Date
                    oConvoc.Cells(indexJury + 1, 4) = LaDate
                    oConvoc.Cells(indexJury, 5) = HoraireConvoc    ' Heure Convocation
                    oConvoc.Cells(indexJury + 1, 5) = HoraireConvoc    '
                    oConvoc.Cells(indexJury, 6) = DebutEpreuve    ' Debut Eprv
                    oConvoc.Cells(indexJury + 1, 6) = DebutEpreuve
                    oConvoc.Cells(indexJury, 7) = duréePassation    ' Debut Eprv
                    oConvoc.Cells(indexJury + 1, 7) = duréePassation  ' Debut Eprv
                    oConvoc.Cells(indexJury, 10) = j - 10    ' Debut Eprv
                    oConvoc.Cells(indexJury + 1, 10) = j - 10  ' Debut Eprv
                    
                    '--- Recherche habilitation
                    For k = 2 To 100
                        If Trim(tablo(0)) = Trim(oParam.Cells(k, 1)) Then
                            oConvoc.Cells(indexJury, 2) = oParam.Cells(k, 3)
                            Exit For
                        End If
                    Next k
                    For k = 2 To 100
                        If Trim(tablo(1)) = Trim(oParam.Cells(k, 1)) Then
                            oConvoc.Cells(indexJury + 1, 2) = oParam.Cells(k, 3)
                            Exit For
                        End If
                    Next k
                    
                    '--- Recherche nombre candidats
                    
                    nbCandidat = 0
                    For k = 8 To 50
                        On Error Resume Next
                        If Len(oSheet.Cells(k, j)) > 5 Then
                            
                            nbCandidat = nbCandidat + 1
                            tablo() = Split(oSheet.Cells(k, 9), "-")
                            HeureFin = tablo(1)
                        'Else
                        '    Exit For
                        End If
                    Next k
                    On Error GoTo 0
                    oConvoc.Cells(indexJury, 9) = nbCandidat
                    oConvoc.Cells(indexJury, 8) = HeureFin
                    oConvoc.Cells(indexJury + 1, 9) = nbCandidat
                    oConvoc.Cells(indexJury + 1, 8) = HeureFin
                    
                    
                Else
                    Exit For
                End If
                indexJury = indexJury + 2
                
                'oConvoc.Activate
                
                
                
            Next j
            
        End If
        
    Next i
    
    
        
    If okFlag = False Then
        MsgBox "Aucune feuille d'examen n'a été trouvée (les feuilles se nomment A1 A2 B1 B2 C1 C2) et donc il est impossible de générer les documents de convocation.", vbInformation, "Générer Convocations"
        Exit Function
    End If
        
    NbJury = indexJury - 2
    
    '--- Maintenant que toutes les données sont dans la feuille de "base de données" convocation,
    '--- On génère les feuilles de convocaction pour chaque jury
        
    '--- Ouvre le modele qui est dans le dossier courant et qui se nomme ConvocationsModele
    On Error Resume Next
    Set oBookModele = Application.Workbooks.Open(Application.ActiveWorkbook.path + "\ConvocationsModele.xlsx")
    If Err.Number <> 0 Then
        MsgBox "Impossible de trouver le document 'ConvocationsModele.xlsx' dans le dossier courant, la génération est annulée.", vbInformation, "Génération Convocation"
        Exit Function
    End If
    
    
    '--- Pour chaque jury on copie la feuille modèle et on remplis la feuille
    
    
    On Error GoTo 0
    '--- Sélectionne la feuille de modele
    Set oSheet = oBookModele.Sheets("Modèle")
    '--- Ouvre un nouveau classeur en mémoire
    Set oBookConvocations = Application.Workbooks.Add
    
    Dim oSheetJury As New Excel.Worksheet
    For i = 1 To NbJury / 2
        '--- Pour chaque groupe de jury (composé de 2 personnes à convoquer)
        j = i * 2
        oSheet.Copy oBookConvocations.Sheets(Sheets.count)  '--- Copie
        'oBookConvocations.Sheets(2).Delete
        '--- Remplissage de la feuille
        Set oSheetJury = oBookConvocations.Sheets(i)
        oSheetJury.Cells(8, 2) = "CONVOCATION AUX " + oConvoc.Cells(j, 3)
        oSheetJury.Cells(9, 4) = oConvoc.Cells(j, 4)
        
        oSheetJury.Cells(16, 3) = oConvoc.Cells(j, 1)   'nom
        oSheetJury.Cells(17, 3) = oConvoc.Cells(j, 2)
        oSheetJury.Cells(18, 3) = oConvoc.Cells(j, 4)
        oSheetJury.Cells(19, 3) = oConvoc.Cells(j, 5)
        oSheetJury.Cells(20, 3) = oConvoc.Cells(j, 6)
        oSheetJury.Cells(21, 3) = oConvoc.Cells(j, 7)
        oSheetJury.Cells(22, 3) = oConvoc.Cells(j, 8)
        oSheetJury.Cells(23, 3) = oConvoc.Cells(j, 9)
        oSheetJury.Cells(24, 3) = oConvoc.Cells(j, 10)
        
        oSheetJury.Cells(29, 3) = oConvoc.Cells(j + 1, 1) 'nom
        oSheetJury.Cells(30, 3) = oConvoc.Cells(j + 1, 2)
        oSheetJury.Cells(31, 3) = oConvoc.Cells(j + 1, 4)
        oSheetJury.Cells(32, 3) = oConvoc.Cells(j + 1, 5)
        oSheetJury.Cells(33, 3) = oConvoc.Cells(j + 1, 6)
        oSheetJury.Cells(34, 3) = oConvoc.Cells(j + 1, 7)
        oSheetJury.Cells(35, 3) = oConvoc.Cells(j + 1, 8)
        oSheetJury.Cells(36, 3) = oConvoc.Cells(j + 1, 9)
        oSheetJury.Cells(37, 3) = oConvoc.Cells(j + 1, 10)
        oSheetJury.Name = "Convoc" + Str(i)
        
        If MyConvo.GenererPDF = True Then
            typeExamen = Right(Trim(oConvoc.Cells(j, 3)), 2)
            '--- Export PDF
            'oSheetJury.Activate
            Filename = FilePath & "\Convo_" & typeExamen & Trim(oConvoc.Cells(j, 1) & oConvoc.Cells(j + 1, 1)) & "pdf"

            If fso.FileExists(Filename) Then
                fso.DeleteFile (Filename)
            End If
            oSheetJury.ExportAsFixedFormat xlTypePDF, Filename
        End If
                
        
        
    Next i
    
    '--- Et voila, terminé !
    
    
    'dBook.Worksheets(Mois).Name = ColMois.Item(Mois)
    Application.DisplayAlerts = False
    oBookModele.Close (False)
    Application.DisplayAlerts = True
    Set oBookModele = Nothing
    
    MsgBox "La génération des convocations est terminée !", vbInformation, "Génération Convocations"
    
End Function

Public Function ChercheZoneNom(NamedRange As String) As Variant
    Dim oBook As New Excel.Workbook
    Set oBook = Application.Workbooks(1)
    
    Dim NamedRangeAsArray() As Variant
    NamedRangeAsArray() = Range(oBook.Names(NamedRange).RefersToRange.Address)
    'For i = LBound(NamedRangeAsArray()) To UBound(NamedRangeAsArray())
    '    MsgBox NamedRangeAsArray(i, 1)
    'Next i
    ChercheZoneNom = NamedRangeAsArray
End Function

Public Function VerifierCodeSaisis()
    ' Recherche dans la colonne CODE la conformité des codes avec ceux définis dans la feuille
    Dim oBook As New Excel.Workbook
    Dim oSheet As New Excel.Worksheet
    Dim oConsolidation As New Excel.Worksheet
 
    If Application.Workbooks.count = 0 Then
        MsgBox "Contrôle impossible : la feuille de saisie de caisse doit être la feuille active !", vbCritical, "Saisie Caisse"
        Exit Function
    End If

    Repo = MsgBox("Cette opération vérifie que les codes saisis sont corrects !" & vbCrLf & vbCrLf & "Confirmez vous la vérification ?", vbQuestion + vbYesNo, "Contrôle des codes !")
    
    If Repo <> 6 Then
        Exit Function
    End If
    
    DoEvents
    
    Set oBook = Application.Workbooks(1)

    Set oSheet = oBook.ActiveSheet
    
    Dim NomClasseur As String
    
    NomClasseur = oSheet.Name
    
    Dim Continue As Boolean
    Select Case UCase(NomClasseur)
    Case "JANVIER", "FEVRIER", "MARS", "AVRIL", "MAI", "JUIN", "JUILLET", "AOUT", "SEPTEMBRE", "OCTOBRE", "NOVEMBRE", "DECEMBRE"
        Continue = True
    Case Else
        Continue = False
    End Select
   
    If Continue = False Then
        MsgBox "Impossible de contrôler les codes, la feuille en cours n'est pas une feuille de saisie mensuelle.", vbCritical, "Saisie Caisse"
        Exit Function
    End If
    
        
    colCode = oSheet.Cells(2, 11)
    If UCase(colCode) <> "CODE" Then
        MsgBox "Impossible de contrôler les codes, la feuille ne contient pas la mention code en cellule 2,11 - K2", vbCritical, "Saisie Caisse"
        Exit Function
    End If
        
    
    '--- Recherche des codes dans la feuille Consolidation
    Dim ColCodes As New Collection
    
    Set oConsolidation = oBook.Sheets("Consolidation")
    
    Dim lsCode As String
    For i = 4 To 204
        lsCode = oConsolidation.Cells(i, 1)
        If (Left(lsCode, 1) = "R" Or Left(lsCode, 1) = "D") And Mid(lsCode, 3, 1) = "." Then
            ColCodes.Add lsCode
        End If
    Next i
    
    '--- Comparaison des codes de la consolidation avec ceux de la saisie
    Dim codeOk As Boolean
    For i = 4 To 209
        lsCode = oSheet.Cells(i, 11)
        oSheet.Cells(i, 12) = ""
        codeOk = False
        If lsCode <> "" Then
            For j = 1 To ColCodes.count
                If lsCode = ColCodes(j) Then
                    codeOk = True
                    Exit For
                End If
            Next j
            If codeOk = True Then
                oSheet.Cells(i, 12) = "Ok"
            Else
                oSheet.Cells(i, 12) = "Erreur !"
            End If
        End If
        
        
    Next i
    
End Function

Private Function ConvertToLetter(iCol As Integer) As String
   Dim iAlpha As Integer
   Dim iRemainder As Integer
   iAlpha = Int(iCol / 27)
   iRemainder = iCol - (iAlpha * 26)
   If iAlpha > 0 Then
      ConvertToLetter = Chr(iAlpha + 64)
   End If
   If iRemainder > 0 Then
      ConvertToLetter = ConvertToLetter & Chr(iRemainder + 64)
   End If
   
   Select Case ConvertToLetter
   Case "A["
        ConvertToLetter = "BA"
    Case "B\"
        ConvertToLetter = "CB"
   
   End Select
   
End Function

Private Function ChercheNbCandidats(NomJury As String, NomExamen As String)
    '--- Recherche le nombre de candidats Oral pour :
    '---    un membre du jury
    '---    un examen
    
    Dim oExam As New Excel.Worksheet
    
    Set oExam = Application.Workbooks(1).Worksheets(NomExamen)
    
    Dim nbCandidat As Integer
    
    ChercheNbCandidats = 0
    
    nbCandidat = 0
                
    Dim i As Integer
    
    Dim found As Boolean
    found = False
    
    For i = 11 To 15
        If oExam.Cells(4, i) = NomJury Or oExam.Cells(5, i) = NomJury Then
            found = True
            Exit For
        End If
    Next i
    
    If found = False Then
        Exit Function
    End If
    
    Dim vide As Integer
    
    For k = 8 To 50
        'On Error Resume Next
        If Len(oExam.Cells(k, i)) > 5 Then

            nbCandidat = nbCandidat + 1
        Else
            vide = vide + 1
            If vide > 10 Then
                '--- Si 10 lignes vides, on sort
                Exit For
            End If
        End If
    Next k
    
    ChercheNbCandidats = nbCandidat
        
End Function

' Excel macro to export all VBA source code in this project to text files for proper source control versioning
' Requires enabling the Excel setting in Options/Trust Center/Trust Center Settings/Macro Settings/Trust access to the VBA project object model
Public Sub ExportVisualBasicCode()
    '-------------------------------------------------------
    '--- this original code thanks to Steve Jansen
    '--- https://gist.github.com/steve-jansen/7589478
    '--- use it to "ghithub" the project
    '-------------------------------------------------------
    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24
    
    Dim VBComponent As Object
    Dim count As Integer
    Dim path As String
    Dim directory As String
    Dim extension As String
    Dim fso As New FileSystemObject
    
    If Application.Workbooks.count < 1 Then
        MsgBox "Merci d'ouvrir un classer avant d'exporter le code", vbExclamation
        Exit Sub
    End If
    
    directory = ActiveWorkbook.path & "\VisualBasic"
    count = 0
    
    If Not fso.FolderExists(directory) Then
        Call fso.CreateFolder(directory)
    End If
    Set fso = Nothing
    
    For Each VBComponent In Application.VBE.VBProjects.Item(1).VBComponents 'ActiveWorkbook.VBProject.VBComponents
        Select Case VBComponent.Type
            Case ClassModule, Document
                extension = ".cls"
            Case Form
                extension = ".frm"
            Case Module
                extension = ".bas"
            Case Else
                extension = ".txt"
        End Select
            
                
        On Error Resume Next
        Err.Clear
        
        path = directory & "\" & VBComponent.Name & extension
        Call VBComponent.Export(path)
        
        If Err.Number <> 0 Then
            Call MsgBox("Failed to export " & VBComponent.Name & " to " & path, vbCritical)
        Else
            count = count + 1
            Debug.Print "Exported " & Left$(VBComponent.Name & ":" & Space(Padding), Padding) & path
        End If

        On Error GoTo 0
    Next
    
    Application.StatusBar = "Successfully exported " & CStr(count) & " VBA files to " & directory
    Application.OnTime Now + TimeSerial(0, 0, 10), "ClearStatusBar"
End Sub

