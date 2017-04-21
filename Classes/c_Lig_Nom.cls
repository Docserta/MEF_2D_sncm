VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Lig_Nom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_Lig_Nom
'Classe des lignes de nomenclature Snecma CFMI std 121
'valable pour l'entète et pour les lignes de nom
'Les informations sont collectées dans un fichier texte
Private mvarNo As String        'Index
Private mvarNom As String       'Nom de l'oblet ligne dans Catia (ex Droite.50....)
Private mvarXd As Double         'Position en X du point de départ de la ligne
Private mvarYd As Double         'Position en Y du point de départ de la ligne
Private mvarXa As Double         'Position en X du point d'arrivee de la ligne
Private mvarYa As Double         'Position en Y du point d'arrivee de la ligne
Private mvarCoul As String        'Couleur de la ligne


Public Property Let No(ByVal vData As String)
    mvarNo = vData
End Property

Public Property Get No() As String
    No = mvarNo
End Property

Public Property Let Nom(ByVal vData As String)
    mvarNom = vData
End Property

Public Property Get Nom() As String
    Nom = mvarNom
End Property

Public Property Let Xd(ByVal vData As Double)
    mvarXd = vData
End Property

Public Property Get Xd() As Double
    Xd = mvarXd
End Property

Public Property Let Yd(ByVal vData As Double)
    mvarYd = vData
End Property

Public Property Get Yd() As Double
    Yd = mvarYd
End Property

Public Property Let Xa(ByVal vData As Double)
    mvarXa = vData
End Property

Public Property Get Xa() As Double
    Xa = mvarXa
End Property

Public Property Let Ya(ByVal vData As Double)
    mvarYa = vData
End Property

Public Property Get Ya() As Double
    Ya = mvarYa
End Property
