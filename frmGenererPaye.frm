VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGenererPaye 
   Caption         =   "DELF-DALF [Génération Bulletin de Paye]"
   ClientHeight    =   9525
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   8280
   OleObjectBlob   =   "frmGenererPaye.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGenererPaye"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myXian As New Xian
Dim dPersonnes As New Dictionary


'---------------------------------------------------
'--- Formulaire pour générer les bulletins de Paye
'--- (C) Renaud Coustellier 2017
'--- pour l'Alliance Française de Xian
'----------------------------------------------------

Private Sub btnAnnuler_Click()
    Unload Me
End Sub

Private Sub btnAucun_Click()
    Dim i
    Dim cherche As String
    Dim vArray As Variant

    For i = 1 To dPersonnes.count
        cherche = dPersonnes.Items(i - 1)
        
        If InStr(cherche, "|") <> 0 Then
            vArray = Left(cherche, Len(cherche) - 1)
            dPersonnes.Item(Left(cherche, Len(cherche) - 1)) = vArray
        End If
    Next i
    
    Affichage
    

End Sub

Private Sub btnListMoins_Click()
    Dim Personne As String
    Dim vArray As Variant

    ReDim vArray(0 To 2)
    
    If Me.lstSelection <> "" Then
        'If Me.lstSelection.ListCount < 2 Then
        Personne = Me.lstSelection.Text
        vArray = dPersonnes.Item(Personne)
        vArray = Personne
        dPersonnes.Item(Personne) = vArray
        'Me.lstSelection.AddItem addJury
    End If
    Affichage

End Sub

Private Sub btnListPlus_Click()
    Dim Personne As String
    Dim vArray As Variant

    ReDim vArray(0 To 2)
    
    If Me.lstJury.Text <> "" Then
        'If Me.lstSelection.ListCount < 2 Then
        Personne = Me.lstJury.Text
        vArray = dPersonnes.Item(Personne)
        vArray = Personne & "|"
        dPersonnes.Item(Personne) = vArray
        'Me.lstSelection.AddItem addJury
    End If
    Affichage
End Sub

Private Sub btnListTous_Click()
    Dim i
    Dim cherche As String
    Dim vArray As Variant

    For i = 1 To dPersonnes.count
        cherche = dPersonnes.Items(i - 1)
        If InStr(cherche, "|") = 0 Then
            vArray = dPersonnes.Items(i - 1) & "|"
            dPersonnes.Item(cherche) = vArray
        End If
    Next i
    
    Affichage
    
End Sub

Private Sub btnOk_Click()
    Dim MaPaye As New oPaye
    Dim cherche As String

    If Me.chkA1.Value = True Then MaPaye.examen.Add "A1", "A1"
    If Me.chkA2.Value = True Then MaPaye.examen.Add "A2", "A2"
    If Me.chkB1.Value = True Then MaPaye.examen.Add "B1", "B1"
    If Me.chkB2.Value = True Then MaPaye.examen.Add "B2", "B2"
    If Me.chkC1.Value = True Then MaPaye.examen.Add "C1", "C1"
    If Me.chkC2.Value = True Then MaPaye.examen.Add "C2", "C2"
    
    Dim i As Integer
    
    For i = 1 To dPersonnes.count
        cherche = dPersonnes.Items(i - 1)
        If InStr(cherche, "|") <> 0 Then
            
            MaPaye.Personnes.Add Left(cherche, Len(cherche) - 1), Left(cherche, Len(cherche) - 1)
        End If
    Next i

    MaPaye.ModelePath = Me.txtModele
    MaPaye.Message1 = Me.txtPaye1
    MaPaye.Message2 = Me.txtPaye2
    
    If Me.chkPDF.Value = True Then
        MaPaye.GeneratePDF = True
    End If
    
    myXian.GenererPaye MaPaye
End Sub

Private Sub chkA1_Change()
    RecherchePersonnes
End Sub

Private Sub chkA2_Change()
    RecherchePersonnes
End Sub

Private Sub chkB1_Change()
    RecherchePersonnes
End Sub

Private Sub chkB2_Change()
    RecherchePersonnes
End Sub

Private Sub chkC1_Change()
    RecherchePersonnes
End Sub

Private Sub chkC2_Change()
    RecherchePersonnes
End Sub




Private Sub Frame6_Click()

End Sub

Private Sub lstJury_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btnListPlus_Click
End Sub

Private Sub lstSelection_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btnListMoins_Click
End Sub

Private Sub UserForm_Initialize()
    Dim oBook As New Excel.Workbook
    Dim oSheet As New Excel.Worksheet
    Dim oParam As New Excel.Worksheet

    
    Me.chkA1.Enabled = False
    Me.chkA2.Enabled = False
    Me.chkB1.Enabled = False
    Me.chkB2.Enabled = False
    Me.chkC1.Enabled = False
    Me.chkC2.Enabled = False
   
    
    Set oBook = Application.Workbooks(1)
    Dim i As Integer
    For i = 1 To oBook.Sheets.count
        '--- Recherche des feuilles d'examen présentes dans le document en cours
        '--- Attention la recherche est stricte, pas d'espace dans les noms des documents !
        Set oSheet = oBook.Sheets(i)
        Select Case Trim(oSheet.Name)
        Case "A1", "A2", "B1", "B2", "C1", "C2"
            '--- Trouvé !
            Me.Controls("chk" & oSheet.Name).Enabled = True
            Me.Controls("chk" & oSheet.Name).Value = True
            
        End Select
    Next i
    Me.txtModele = Application.ActiveWorkbook.path + "\BulletinPayeModele.xlsx"
    
    RecherchePersonnes
    
    Set oBook = Nothing
    Set oSheet = Nothing
    Set oParam = Nothing
    
End Sub


Private Sub RecherchePersonnes()
    '--- Recherche la liste des personnes en fonction des examens sélectionnés
    Dim oBook As New Excel.Workbook
    Dim oSheet As New Excel.Worksheet
    Dim oParam As New Excel.Worksheet

    dPersonnes.RemoveAll
    'dPaye.RemoveAll
    
    Dim dExamens As New Dictionary
    
    dExamens.RemoveAll
    dExamens.Add "A1", "A1"
    dExamens.Add "A2", "A2"
    dExamens.Add "B1", "B1"
    dExamens.Add "B2", "B2"
    dExamens.Add "C1", "C1"
    dExamens.Add "C2", "C2"
    
    
    Dim i As Integer
    Dim j As Integer
    
    Dim startPaye As Integer
    Dim endPaye As Integer
    
    Dim montant1 As Integer
    Dim montant2 As Integer
    
    
    Dim cherche As String
    For i = 1 To 6
        '--- Trouvé !
        If Me.Controls("chk" & dExamens.Items(i - 1)).Value = True Then
            '--- On va chercher les Personnes
            Set oBook = Application.ActiveWorkbook
            On Error Resume Next
            Set oSheet = oBook.Worksheets(dExamens.Items(i - 1))
            If Err.Number <> 0 Then
                MsgBox "Impossible de trouver une feulle d'examen : le classeur Excel d'examen doit être le classeur actif !", vbCritical, "Génération paye : erreur"
                
                Exit Sub
            End If
            
            On Error GoTo 0
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
            
            If startPaye <> 0 And endPaye <> 0 Then
                '--- On va chercher dans la zone trouvée qui l'on doit payer et combien !
                For j = startPaye To endPaye
                    cherche = oSheet.Cells(j, 3)
                    If cherche <> "" Then
                        'montant1 = 0
                        'montant2 = 0
                        'If Val(oSheet.Cells(j, 4)) <> 0 Then
                        '    montant1 = Val(oSheet.Cells(j, 4))
                        'End If
                        'If Val(oSheet.Cells(j, 5)) <> 0 Then
                        '    montant2 = Val(oSheet.Cells(j, 5))
                        'End If
                        
                        'If montant1 + montant2 <> 0 Then
                            If Not dPersonnes.Exists(cherche) Then
                                dPersonnes.Add cherche, cherche '+ "|" + Str(dIndex), cherche + "|" + dExamens.Items(i - 1) + "|" + Str(montant1) + "|" + Str(montant2)
                            End If
                        'End If
                    End If
                Next j
            End If
        End If
        
    Next i
    
    Set oBook = Nothing
    Set oSheet = Nothing
    Set oParam = Nothing

    
    Affichage
    
End Sub

Public Sub Affichage()
    Dim i As Integer
    
    Me.lstJury.Clear
    Me.lstSelection.Clear
    
    Dim cherche As String
    Dim tablo() As String
    
    For i = 1 To dPersonnes.count
    
        cherche = dPersonnes.Items(i - 1)
        If InStr(cherche, "|") Then
            tablo() = Split(cherche, "|")
            Me.lstSelection.AddItem tablo(0)
        Else
            Me.lstJury.AddItem cherche
        End If
    Next i
        
    Me.lblPersonnes = Me.lstJury.ListCount & " personnes"
    Me.lblPersonnesAPayer = Me.lstSelection.ListCount & " personnes à payer"
   
End Sub
