Attribute VB_Name = "General"
'---------------------------------------------------------------------------
' Renaud Coustellier pour AF XIAN
'---------------------------------------------------------------------------
' Version .99. du Février 2016 : Statut - Beta
'---------------------------------------------------------------------------
' Les outils ci-présents "as is"
'---------------------------------------------------------------------------
' Attention : pour fonctionner vous devez bla bla bla
'---------------------------------------------------------------------------
' Get information from this address :
' http://www.ozgrid.com/VBA/DesciptionToUDF.htm

Option Private Module

'-------------------------------------------------------------
' Constantes liées au menu
'-------------------------------------------------------------
Global Const NOM_ADDIN = "AFXian-RC"
Global Const NOM_MENU_EXCEL = "Worksheet Menu Bar"
Global Const NOM_MENU_GENERAL = "DELF_DALF"
Global Const TEXTE_MENU_GENERAL = "&DELF_DALF"


Global Const NOM_SOUSMENU_10 = "Préparer une feuille d'examen..."
Global Const TEXTE_SOUSMENU_10 = "Préparer une feuille d'examen..."

Global Const NOM_SOUSMENU_11 = "Générer les convocations..."
Global Const TEXTE_SOUSMENU_11 = "Générer les convocations..."

Global Const NOM_SOUSMENU_12 = "Calculer Coût prévisionnel..."
Global Const TEXTE_SOUSMENU_12 = "Calculer Coût prévisionnel..."


Global Const NOM_SOUSMENU_13 = "Générer les bulletins de paye..."
Global Const TEXTE_SOUSMENU_13 = "Générer les bulletins de paye..."

'Global Const NOM_SOUSMENU_13 = "Importer objet (depuis source vers destination)..."
'Global Const TEXTE_SOUSMENU_13 = "Importer objet (depuis source vers destination)..."

'Global Const NOM_SOUSMENU_14 = "Corriger la NC CB..."
'Global Const TEXTE_SOUSMENU_14 = "Corriger la NC CB..."

'Global Const NOM_SOUSMENU_15 = "Rechercher dans les Tables..."
'Global Const TEXTE_SOUSMENU_15 = "Rechercher dans les Tables..."

'Global Const NOM_SOUSMENU_16 = "Generation des Scripts(vues, procédures)..."
'Global Const TEXTE_SOUSMENU_16 = "Generation des Scripts(vues, procédures)..."

Global Const NOM_SOUSMENU_20 = "Aide DELF_DALF..."
Global Const TEXTE_SOUSMENU_20 = "Aide DELF_DALF..."
Global Const NOM_SOUSMENU_22 = "A Propos..."
Global Const TEXTE_SOUSMENU_22 = "A Propos..."

Global Const version = "Version 0.99 Février 2017"

'-------------------------------------------------------------
' Constantes de paramétrages
'-------------------------------------------------------------
Global Const TITRE_FORMDATALINK = "DELF_DALF-RC"
'---------------------------------------------------------------


Sub auto_open()
    On Error Resume Next
    Dim SMMenu As Object
    Dim oXL As Application
    Dim Index
    Dim TagSearchMenu
    Dim DataLinkMenu
    'Set ClassAlarmsManager = cAlarmsManager
   
    Set oXL = Application
    Set SMMenu = oXL.CommandBars(NOM_MENU_EXCEL).FindControl(Tag:=NOM_MENU_GENERAL)
    If SMMenu Is Nothing Then
        Index = oXL.CommandBars(NOM_MENU_EXCEL).Controls.count
        Set SMMenu = oXL.CommandBars(NOM_MENU_EXCEL).Controls.Add( _
                Type:=msoControlPopup, Before:=Index, Temporary:=True)
        With SMMenu
            .Visible = True
            .Caption = TEXTE_MENU_GENERAL
            .Tag = NOM_MENU_GENERAL
        End With
        
        Set TagSearchMenu = SMMenu.Controls.Add()
        
        With TagSearchMenu
            .Caption = TEXTE_SOUSMENU_10
            .Tag = NOM_SOUSMENU_10
            .OnAction = "Preparer_Feuille_Test_Click"
        End With
        
        Set TagSearchMenu = SMMenu.Controls.Add()
        
        With TagSearchMenu
            .Caption = TEXTE_SOUSMENU_11
            .Tag = NOM_SOUSMENU_11
            .OnAction = "Generer_Convocations_Click"
        End With

        Set TagSearchMenu = SMMenu.Controls.Add()
        
        With TagSearchMenu
            .Caption = TEXTE_SOUSMENU_12
            .Tag = NOM_SOUSMENU_12
            .OnAction = "Calculer_CoutPrevisionnel_Click"
        End With
        Set DataLinkMenu = SMMenu.Controls.Add()
        
        'DataLinkMenu.BeginGroup = True
        With DataLinkMenu
            .Caption = TEXTE_SOUSMENU_13
            .Tag = NOM_SOUSMENU_13
            .OnAction = "Generer_Paye_Click"
        End With
        
'        Set DataLinkMenu = SMMenu.Controls.Add()
        'DataLinkMenu.BeginGroup = True
'        With DataLinkMenu
'            .Caption = TEXTE_SOUSMENU_13
'            .Tag = NOM_SOUSMENU_13
'            .OnAction = "ImporterObjetMDB_Click"
'        End With
'
'        'Set DataLinkMenu = SMMenu.Controls.Add()
'
'        DataLinkMenu.BeginGroup = True
'        With DataLinkMenu
'            .Caption = TEXTE_SOUSMENU_14
'            .Tag = NOM_SOUSMENU_14
'            .OnAction = "CorrigerNCCB_Click"
'        End With
'
'        DataLinkMenu.BeginGroup = True
'        With DataLinkMenu
'            .Caption = TEXTE_SOUSMENU_15
'            .Tag = NOM_SOUSMENU_15
'            .OnAction = "ListeDesTables_Click"
'        End With
'
'        Set DataLinkMenu = SMMenu.Controls.Add()
'        With DataLinkMenu
'            .Caption = TEXTE_SOUSMENU_16
'            .Tag = NOM_SOUSMENU_16
'            .OnAction = "GenerationScript_Click"
'        End With
'
'
        Set DataLinkMenu = SMMenu.Controls.Add()
        DataLinkMenu.BeginGroup = True
        With DataLinkMenu
            .Caption = TEXTE_SOUSMENU_20
            .Tag = NOM_SOUSMENU_20
            .OnAction = "AfficheAide_Click"

        End With

        Set DataLinkMenu = SMMenu.Controls.Add()
        'DataLinkMenu.BeginGroup = True
        With DataLinkMenu
            .Caption = TEXTE_SOUSMENU_22
            .Tag = NOM_SOUSMENU_22
            .OnAction = "APropos_Click"
        End With
        oXL.AddIns(NOM_ADDIN).Installed = True
    End If
End Sub
Private Sub APropos_Click()
    frmAbout.Show
End Sub
Private Sub AfficheAide_Click()
    Dim retour As Long
    Dim fso As New FileSystemObject
    Dim AideFileName As String
    'Dim MyClasseIni As New cIni
    
    AideFileName = "\\AllianceFrance\Public\Documents informatique\Applications\AFXLA\AideXLA.htlm"
    retour = VBA.Shell("C:\Program Files\Internet Explorer\IEXPLORE.EXE " & AideFileName, vbMaximizedFocus)

End Sub
Private Sub Preparer_Feuille_Test_Click()
    
    frmCreateExam.Show
    
End Sub
Private Sub Generer_Convocations_Click()
    frmConvocations.Show
End Sub

Private Sub Generer_Paye_Click()
    frmGenererPaye.Show
End Sub


Private Sub Calculer_CoutPrevisionnel_Click()
    'Dim MyTool As New Xian
    'MyTool.CalculerCoutPrevisionnel
    'Set MyTool = Nothing

End Sub


Public Function CoutEpreuveIndividuelle(exam As String, Personne As String, NbCandidats As Integer, NbSujets As Integer) As Integer
    '--- Renvoie le salaire pour un examen, un nb de candidats, un nb de sujets
    Dim param As New oParam
    param.LectureNiveauTarif
    
    If param.PersonneAPayer(Personne) = True Then
    
        CoutEpreuveIndividuelle = param.GetParam(exam, "Lecture sujet") * NbSujets + param.GetParam(exam, "Passation oraux") * NbCandidats
    Else
        CoutEpreuveIndividuelle = 0
    End If
    Set param = Nothing
End Function

Public Function CoutEpreuveCollective(exam As String, Personne As String, NbCandidats As Integer, NbSujets As Integer) As Integer
    '--- Renvoie le salaire pour un examen, un nb de candidats, un nb de sujets
    Dim param As New oParam
    param.LectureNiveauTarif
    
    If param.PersonneAPayer(Personne) = True Then
        CoutEpreuveCollective = param.GetParam(exam, "Connaissance sujet ép collective") * NbSujets + param.GetParam(exam, "Correction copies") * NbCandidats
    Else
        CoutEpreuveCollective = 0
    End If
    Set param = Nothing
End Function

Public Function HeurePreparation(heureDepart As String, Optional DureePassation As Integer) As String
    ' Calcule l'heure de préparation = l'heure de début de préparation + la durée de passation
    Dim tablo() As String
    'Durées préparation  Durées passation

    
    Dim hDepart As Date
    Dim hFin As Date
    
    If heureDepart = "" Then
        Exit Function
    End If
    
    Dim param As New oParam
    param.LectureNiveauTarif
    
    Dim oSheet As New Excel.Worksheet
    
    Set oSheet = Application.ActiveSheet
    
    If duréeePrepa = 0 Then
        duréeePrepa = Val(param.GetParam(oSheet.Name, "Durées préparation"))
    End If
    
    If DureePassation = 0 Then
        DureePassation = Val(param.GetParam(oSheet.Name, "Durées passation"))
    End If
    
    
    
    tablo() = Split(heureDepart, "-")
    tablo(0) = Replace(tablo(0), "h", ":")
    
    
    hDepart = CDate(tablo(0)) + (1 / 1440 * DureePassation)
    
    hFin = hDepart + (1 / 1440 * duréeePrepa)
    
    HeurePreparation = Format(hDepart, "HH:MM") & " - " & Format(hFin, "HH:MM")
End Function



Public Function HeurePassation(heureDepart As String, Optional duréeePrepa As Integer, Optional DureePassation As Integer) As String
    Dim tablo() As String
    'Durées préparation  Durées passation

    If heureDepart = "" Then
        HeurePassation = ""
        Exit Function
    End If
    
    Dim hDepart As Date
    Dim hFin As Date
    
    
    Dim param As New oParam
    param.LectureNiveauTarif
    
    Dim oSheet As New Excel.Worksheet
    
    Set oSheet = Application.ActiveSheet
    
    If duréeePrepa = 0 Then
        duréeePrepa = Val(param.GetParam(oSheet.Name, "Durées préparation"))
    End If
    
    If DureePassation = 0 Then
        DureePassation = Val(param.GetParam(oSheet.Name, "Durées passation"))
    End If
    
    
    
    tablo() = Split(heureDepart, "-")
    tablo(0) = Replace(tablo(0), "h", ":")
    
    
    hDepart = CDate(tablo(0))
    
    hFin = hDepart + (1 / 1440 * duréeePrepa) + (1 / 1440 * DureePassation)
    
    HeurePassation = Format(hDepart + (1 / 1440 * duréeePrepa), "HH:MM") & " - " & Format(hFin, "HH:MM")
    
End Function
Public Function DureePreparation(Optional exam As String) As Integer
    'renvoie la durée de préparation
    Dim param As New oParam
    
    Dim oSheet As New Excel.Worksheet
    
    Set oSheet = Application.ActiveSheet
    If exam = "" Then
        exam = oSheet.Name
    End If
    
    param.LectureNiveauTarif
    DureePreparation = param.GetParam(exam, "Durées préparation") ' Durées passation
    Set oSheet = Nothing
    Set param = Nothing


End Function

Public Function DureePassation(Optional exam As String) As Integer
    'renvoie la durée de préparation
    Dim param As New oParam
    
    Dim oSheet As New Excel.Worksheet
    
    Set oSheet = Application.ActiveSheet
    If exam = "" Then
        exam = oSheet.Name
    End If
    
    param.LectureNiveauTarif
    DureePassation = param.GetParam(exam, "Durées passation") ' Durées passation
    Set oSheet = Nothing
    Set param = Nothing


End Function


Public Function CoutSurveillance(exam As String, sujet As String) As Integer
    '--- Cette fonction renvoie le tarif d'un param?tre de surveillance
    Dim param As New oParam
    param.LectureNiveauTarif
    
    CoutSurveillance = param.GetParam(exam, sujet)
    Set param = Nothing
    
End Function

Public Function testOk() As String
    testOk = "ok"
End Function

