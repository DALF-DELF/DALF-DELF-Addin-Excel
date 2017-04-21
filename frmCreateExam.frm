VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCreateExam 
   Caption         =   "DELF-DALF [Création examen]"
   ClientHeight    =   9708
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   7440
   OleObjectBlob   =   "frmCreateExam.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCreateExam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myXian As New Xian
Dim dJury As New Dictionary
'---------------------------------------------------
'--- Formulaire pour générer les Feuilles d'examen
'--- (C) Renaud Coustellier 2017
'--- pour l'Alliance Française de Xian
'----------------------------------------------------

Private Sub btJuryPlus_Click()
    If Val(Me.txtNbJurys.Text) < 10 Then
        Me.txtNbJurys.Text = Str(Val(Me.txtNbJurys.Text) + 1)
        JurysTabs (1)
    End If
    
End Sub

Private Sub btnJuryMoins_Click()
    If Val(Me.txtNbJurys.Text) > 1 Then
        Me.txtNbJurys.Text = Str(Val(Me.txtNbJurys.Text) - 1)
        JurysTabs (-1)
    End If
    
End Sub
Private Sub JurysTabs(nbTabs As Integer)
    If nbTabs = 1 Then
        Me.tsJurys.Tabs.Add
        Me.tsJurys.Tabs(Me.tsJurys.Tabs.count - 1).Caption = "Jury" + Str(Me.tsJurys.Tabs.count)
    Else
        Me.tsJurys.Tabs.Remove (Me.tsJurys.Tabs.count - 1)
    End If
End Sub

Private Sub btnListMoins_Click()
    Dim addJury As String
    Dim vArray As Variant

    ReDim vArray(0 To 2)
    
    'If Me.lstSelection.ListCount < 2 Then
        addJury = Me.lstSelection.Text
        vArray = dJury.Item(addJury)
        vArray = addJury & "#" '& tsJurys.SelectedItem.Caption
        dJury.Item(addJury) = vArray
        'Me.lstSelection.AddItem addJury
    'End If
        
    AfficheListe


End Sub

Private Sub btnListPlus_Click()
    
    Dim addJury As String
    Dim vArray As Variant

    ReDim vArray(0 To 2)
    
    If Me.lstSelection.ListCount < 2 Then
        addJury = Me.lstJury.Text
        vArray = dJury.Item(addJury)
        vArray = addJury & "#" & tsJurys.SelectedItem.Caption
        dJury.Item(addJury) = vArray
        'Me.lstSelection.AddItem addJury
    End If
        
    AfficheListe
    
End Sub

Private Sub CommandButton1_Click()
    Dim Repo  As Variant
    Dim msg As String
        
    msg = "Voulez-vous créer le formulaire de l'examen ?"
    Repo = MsgBox(msg, vbQuestion + vbYesNo, "Créer formulaire")
    
    If Repo = 6 Then
        
    
        '--- On génère la feuille tout de suite
        'Dim myXian As New Xian

        Dim myExam As New oDELF_DALF
        
        Dim bOkExam As Boolean
        bOkExam = False
        
        If Me.optA1.Value = True Then
            myExam.Niveau = "A1"
            bOkExam = True
        ElseIf Me.optA2.Value = True Then
            myExam.Niveau = "A2"
            bOkExam = True
        ElseIf Me.optB1.Value = True Then
            myExam.Niveau = "B1"
            bOkExam = True
        ElseIf Me.optB2.Value = True Then
            myExam.Niveau = "B2"
            bOkExam = True
        ElseIf Me.optC1.Value = True Then
            myExam.Niveau = "C1"
            bOkExam = True
        ElseIf Me.optC2.Value = True Then
            myExam.Niveau = "C2"
            bOkExam = True
        End If
          
        If bOkExam = False Then
            MsgBox "Merci de sélectionner un examen !", vbCritical, "Erreur, pas d'examen sélectionné !"
            Exit Sub
        End If
          
        
        
        myExam.NbCandidats = Val(Me.txtNbCandidats)
        myExam.nbJurys = Val(Me.txtNbJurys)
        myExam.DateExamen = Me.txtDateExam
        myExam.Coordinateur = Me.txtCoordinateur
        Set myExam.Jurys = dJury
             
        myXian.PrepareFeuilleExamen myExam
           
    End If
    
    Unload Me
    
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub


Private Sub lstJury_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btnListPlus_Click
End Sub

Private Sub lstSelection_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btnListMoins_Click
End Sub
Private Sub optA1_Click()
    updateJurys
End Sub

Private Sub optA2_Click()
    updateJurys
End Sub
Private Sub optB1_Click()
    updateJurys
End Sub
Private Sub optB2_Click()
    updateJurys
End Sub

Private Sub optC1_Click()
    updateJurys "C1"
End Sub

Private Sub optC2_Click()
    updateJurys "C1"
End Sub

Private Function updateJurys(Optional C1 As String) As Boolean
    '--- recherche des jurys et Remplissage de la liste des jurys
    Dim jury As Variant
    Dim i As Integer
    jury = myXian.ChercheZoneNom("Jury")
    
    dJury.RemoveAll
    
    For i = 1 To UBound(jury, 1)
        If C1 = "" Then
            If Not dJury.Exists(jury(i, 1)) Then
                dJury.Add jury(i, 1), jury(i, 1) & "#"
            End If
        Else
            If Not dJury.Exists(jury(i, 1)) And InStr(jury(i, 3), "C2") Then
                dJury.Add jury(i, 1), jury(i, 1) & "#"
            End If
        End If
    Next i
    AfficheListe
    
End Function


Private Sub tsJurys_Change()
        
    AfficheListe
End Sub

Private Sub UserForm_Initialize()
    Dim oBook As New Excel.Workbook
    Dim oSheet As New Excel.Worksheet
    Dim oParam As New Excel.Worksheet

    Me.optA1.Enabled = True
    Me.optA2.Enabled = True
    Me.optB1.Enabled = True
    Me.optB2.Enabled = True
    Me.optC1.Enabled = True
    Me.optC2.Enabled = True
    
    Set oBook = Application.Workbooks(1)
    Dim i As Integer
    For i = 1 To oBook.Sheets.count
        '--- Recherche des feuilles d'examen présentes dans le document en cours
        '--- Attention la recherche est stricte, pas d'espace dans les noms des documents !
        Set oSheet = oBook.Sheets(i)
        Select Case Trim(oSheet.Name)
        Case "A1", "A2", "B1", "B2", "C1", "C2"
            '--- Trouvé !
            Me.Controls("opt" & oSheet.Name).Enabled = False
            Me.Controls("opt" & oSheet.Name).Value = False
        End Select
    Next i

    Do While Me.tsJurys.Tabs.count > 1
        Me.tsJurys.Tabs.Remove 0
    Loop
    'Me.tsJurys.Value     .Caption = "11"
    Me.tsJurys.Tabs(0).Caption = "Jury 1"
    
    '--- Recherche de la liste des jurys
    
    '--- Recherche des param-tres
    Set oParam = oBook.Sheets("Parametres")
    oParam.Activate
    Dim Admin As New Dictionary
    'Dim jury As New Dictionary
    
    '--- recherche des jurys et Remplissage de la liste des jurys
    Dim jury As Variant
    jury = myXian.ChercheZoneNom("Jury")
    
    For i = 1 To UBound(jury, 1)
        If Not dJury.Exists(jury(i, 1)) Then
            dJury.Add jury(i, 1), jury(i, 1) & "#"
        End If
    Next i
    AfficheListe
    
    Me.txtDateExam = Format(Now, "dd/mm/yyyy")

End Sub

Private Sub AfficheListe()
    '----------------------------------------------------------------------
    'Affiche la liste des jurys dans la liste
    ' le dictionnaire est composé de la clé et de l'entrée "nom jury" + séparateur # + le jury n* éventuel
    ' s'il n'y a pas de numéro de jury alors ce membre du jury est disponible
    '-------------------------------------------------------------------------
    Dim i As Integer
    Dim tablo() As String
    
    Me.lstJury.Clear
    Me.lstSelection.Clear
    
    '--- Affiche la liste
    For i = 0 To dJury.count - 1
        tablo() = Split(dJury.Items(i), "#")
        If tablo(1) = "" Then
            Me.lstJury.AddItem tablo(0)
        ElseIf tablo(1) = Me.tsJurys.SelectedItem.Caption Then
            Me.lstSelection.AddItem tablo(0)
        End If
    Next i
    
    
    
    
End Sub


