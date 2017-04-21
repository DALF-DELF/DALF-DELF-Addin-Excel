VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "oDELF_DALF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Niveau As String
Public DateExamen As Date
Public NbCandidats As Integer
Public nbJurys As Integer
Public Jurys As New Dictionary
Public Coordinateur As String


Public Function MembreJury(NumJury As Integer, NumMembre As Integer) As String
    '--- Retourne le nom du numMembre (entre 1 et 2) du numéro de jury
    Dim i As Integer
    Dim found As Integer
    Dim tablo() As String
    
    found = 0
    MembreJury = ""
    For i = 0 To Jurys.count - 1
        tablo() = Split(Jurys.Items(i), "#")
        If NumJury = Val(Mid(tablo(1), 5)) Then
            found = found + 1
            If found = NumMembre Then
                MembreJury = tablo(0)
                Exit For
            End If
        End If
    Next i
    
    
    
End Function

