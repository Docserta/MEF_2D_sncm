VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Regle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_Regle
'Classe des carreaux horizontaux ou verticaux
'permattant de localiser les objets dans le plan
Private mvarNo As String
Private mvarLimit As Double

Public Property Let No(ByVal vData As String)
    mvarNo = vData
End Property

Public Property Get No() As String
    No = mvarNo
End Property

Public Property Let Limit(ByVal vData As Double)
    mvarLimit = vData
End Property

Public Property Get Limit() As Double
    Limit = mvarLimit
End Property


