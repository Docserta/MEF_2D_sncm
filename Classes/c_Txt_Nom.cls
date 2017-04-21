VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Txt_Nom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_Txt_Nom
'Classe des texte de l'entète de nomenclature Snecma CFMI std 121
'Les informations sont collectées dans un fichier texte
Private mvarNo As String        'Index
Private mvarNom As String       'Nom de l'oblet texte dans Catia (ex TABLIG1_COL....)
Private mvarTexte As String     'Contenu du texte
Private mvarPolice As String    'Police du texte
Private mvarTaille As String    'Taille du text
Private mvarCoul As Long        'Couleur du texte
Private mvarX As Double         'Position en X du texte
Private mvarY As Double         'Position en Y du texte

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

Public Property Let Texte(ByVal vData As String)
    mvarTexte = vData
End Property

Public Property Get Texte() As String
    Texte = mvarTexte
End Property

Public Property Let Police(ByVal vData As String)
    mvarPolice = vData
End Property

Public Property Get Police() As String
    Police = mvarPolice
End Property

Public Property Let Coul(ByVal vData As Long)
    mvarCoul = vData
End Property

Public Property Get Coul() As Long
    Coul = mvarCoul
End Property

Public Property Let X(ByVal vData As Double)
    mvarX = vData
End Property

Public Property Get X() As Double
    X = mvarX
End Property

Public Property Let Y(ByVal vData As Double)
    mvarY = vData
End Property

Public Property Get Y() As Double
    Y = mvarY
End Property
