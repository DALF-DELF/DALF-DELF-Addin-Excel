VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "oConvocations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--- Objet Convocations
Public examen As New Dictionary         '--- Liste des feuilles � g�n�rer
Public GenererLaFeuille As Boolean      '--- Genere la feuille de convocations ou utilise les donn�es existantes
Public ModelePath As String             '--- Le chemin d'acc�s du mod�le de document
Public GenererPDF As Boolean            '--- Est-ce que l'on g�n�re un PDF ou pas ?
