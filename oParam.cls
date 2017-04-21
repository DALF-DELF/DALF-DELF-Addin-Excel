VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "oParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public A1 As New Dictionary
Public A2 As New Dictionary
Public B1 As New Dictionary
Public B2 As New Dictionary
Public C1 As New Dictionary
Public C2 As New Dictionary


Public Function GetParam(exam, titre)
    Select Case exam
    Case "A1"
        GetParam = A1.Item((titre))
    Case "A2"
        GetParam = A2.Item((titre))
    Case "B1"
        GetParam = B1.Item((titre))
    Case "B2"
        GetParam = B2.Item((titre))
    Case "C1"
        GetParam = C1.Item((titre))
    Case "C2"
        GetParam = C2.Item((titre))
    End Select
    
    
End Function
Public Function PersonneAPayer(Personne As String) As Boolean
    '--- Lecture de la feuille de paramètre pour savoir si la personne est A PAYER
    '--- Par exemple le directeur corrige les copies mais n'est pas payé
    
    Dim oBook As New Excel.Workbook
    Dim oSheet As New Excel.Worksheet
    
    Set oBook = Application.Workbooks(1)
    Set oSheet = oBook.Sheets("Parametres")
    
    PersonneAPayer = True
    
    For i = 2 To 200
        If oSheet.Cells(i, 1) = Personne Then
            If oSheet.Cells(i, 5) = 1 Then
                PersonneAPayer = Fale
            End If
            Exit For
        End If
    Next i

    Set oBook = Nothing
    Set oSheet = Nothing
    
End Function


Public Function LectureNiveauTarif()
    '--- Lecture de A1 à C2
    '--- Sur toutes les colonnes
    
    Dim oBook As New Excel.Workbook
    Dim oSheet As New Excel.Worksheet
    Dim oParam As New Excel.Worksheet
    
    Dim dict As New Dictionary
    
    
    Set oBook = Application.Workbooks(1)
    Set oSheet = oBook.Sheets("Parametres")
    
    LectureNiveauTarif = False
    
    For i = 8 To 50
        If oSheet.Cells(1, i) <> "" Then
            '--- On a trouvé un titre de colonne
            titre = oSheet.Cells(1, i)
            For j = 2 To 7
                If oSheet.Cells(j, i) <> "" Then
                    '--- il y a quelque chose dans la colonne
                    Select Case j
                    Case 2
                        A1.Add titre, oSheet.Cells(j, i)
                    Case 3
                        A2.Add titre, oSheet.Cells(j, i)
                    Case 4
                        B1.Add titre, oSheet.Cells(j, i)
                    Case 5
                        B2.Add titre, oSheet.Cells(j, i)
                    Case 6
                        C1.Add titre, oSheet.Cells(j, i)
                    Case 7
                        C2.Add titre, oSheet.Cells(j, i)
                    End Select
                End If
            Next j
        End If
    Next i

    Set oBook = Nothing
    Set oSheet = Nothing
    

End Function
