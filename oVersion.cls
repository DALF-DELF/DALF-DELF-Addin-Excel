VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "oVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public version As String
Public versionDate As Date
Public comment As String
Public Author As String
Public licence As String


Public Function Initialize() As Boolean

    Initialize = True
    
    Me.version = "1.00"
    Me.comment = "Last version on Github"
    Me.Author = "Renaud Coustellier"
    Me.licence = "GNU PUBLIC LICENCE"
    Me.versionDate = CDate(Format("15/05/2017", "DD/MM/YYY"))
    
End Function
