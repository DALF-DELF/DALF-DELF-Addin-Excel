VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "oPaye"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--- Objet Convocations
Public examen As New Dictionary         '--- Liste des examen
Public Personnes As New Dictionary      '--- Liste des personnes
Public ModelePath As String             '--- Le chemin d'accès du modèle de document
Public Message1 As String
Public Message2 As String
Public GeneratePDF As Boolean


