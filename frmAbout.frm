VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAbout 
   Caption         =   "DALF_DELF"
   ClientHeight    =   5904
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7632
   OleObjectBlob   =   "frmAbout.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btnOk_Click()
    Unload Me
End Sub

Private Sub Label2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim Repo
    
    Repo = MsgBox("Exporter le code du projet", vbQuestion + vbYesNo, "DALF DELF")
    If Repo = 6 Then
        Dim myXian As New Xian
        myXian.ExportVisualBasicCode
    End If
    
End Sub
