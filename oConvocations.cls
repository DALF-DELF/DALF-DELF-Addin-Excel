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
Public examen As New Dictionary         '--- Liste des feuilles à générer
Public GenererLaFeuille As Boolean      '--- Genere la feuille de convocations ou utilise les données existantes
Public ModelePath As String             '--- Le chemin d'accès du modèle de document
Public GenererPDF As Boolean            '--- Est-ce que l'on génère un PDF ou pas ?
