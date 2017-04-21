VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_View"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_View

Private mvarNom As String
Private mvarEchelle As Double
Private mvarPosition As String
Private mvarRefView As DrawingView

Public Property Let Nom(ByVal vData As String)
    mvarNom = vData
End Property

Public Property Get Nom() As String
    Nom = mvarNom
End Property

Public Property Let echelle(ByVal vData As Double)
    mvarEchelle = vData
End Property

Public Property Get echelle() As Double
    echelle = mvarEchelle
End Property

Public Property Let Position(ByVal vData As String)
    mvarPosition = vData
End Property

Public Property Get Position() As String
    Position = mvarPosition
End Property

Public Property Let RefView(ByVal vData As DrawingView)
    mvarRefView = vData
End Property

Public Property Get RefView() As DrawingView
    RefView = mvarRefView
End Property
