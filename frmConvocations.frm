VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConvocations 
   Caption         =   "DELF-DALF [Convocations]"
   ClientHeight    =   6444
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   9180
   OleObjectBlob   =   "frmConvocations.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmConvocations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'--- Formulaire pour générer les Convocations
'--- (C) Renaud Coustellier 2017
'--- pour l'Alliance Française de Xian
'----------------------------------------------------


Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub btnAnnuler_Click()
    Unload Me
End Sub



Private Sub btnOk_Click()
    Dim myCon As New oConvocations
    
    If Me.chkA1.Value = True Then myCon.examen.Add "A1", "A1"
    If Me.chkA2.Value = True Then myCon.examen.Add "A2", "A2"
    If Me.chkB1.Value = True Then myCon.examen.Add "B1", "B1"
    If Me.chkB2.Value = True Then myCon.examen.Add "B2", "B2"
    If Me.chkC1.Value = True Then myCon.examen.Add "C1", "C1"
    If Me.chkC2.Value = True Then myCon.examen.Add "C2", "C2"
        
    myCon.GenererLaFeuille = Me.opt1.Value
    myCon.ModelePath = Me.txtModele.Text
    
    If Me.chkPDF.Value = True Then
        myCon.GenererPDF = True
    End If
    
    Dim MyTool As New Xian
    
    MyTool.GenererConvocation myCon
    
    Set MyTool = Nothing
    
    Unload Me
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
   
    If Application.Workbooks.count = 0 Then
        MsgBox "Aucun classeur disponible ", vbExclamation, "Générer les convocations"
        Unload Me
        Exit Sub
    End If
    
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

    Me.txtModele = Application.ActiveWorkbook.path + "\ConvocationsModele.xlsx"


End Sub
