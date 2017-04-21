VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_ItNomSnecma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_ItNom
'Classe Item Nomenclature
Private mvarNo As String        'Index
Private mvarG03 As String       'Contenu de la colonne G03
Private mvarG02 As String       'Contenu de la colonne G02
Private mvarG01 As String       'Contenu de la colonne G01
Private mvarIdent As String     'Contenu de la colonne Identification
Private mvarCageCode As String  'Contenu de la colonne CageCode
Private mvarDet As String       'Contenu de la colonne Det/Sh/NO
Private mvarZone As String       'Contenu de la colonne ITEM/ZONE
Private mvarRep As String       'Contenu de la colonne ITEM/NO/Rep
Private mvarDesc As String       'Contenu de la colonne Description
Private mvarFourn As String      'Adresses des fournisseurs

Public Property Let No(ByVal vData As String)
    mvarNo = vData
End Property

Public Property Get No() As String
    No = mvarNo
End Property

Public Property Let G03(ByVal vData As String)
    mvarG03 = vData
End Property

Public Property Get G03() As String
    G03 = mvarG03
End Property

Public Property Let G02(ByVal vData As String)
    mvarG02 = vData
End Property

Public Property Get G02() As String
    G02 = mvarG02
End Property

Public Property Let G01(ByVal vData As String)
    mvarG01 = vData
End Property

Public Property Get G01() As String
    G01 = mvarG01
End Property

Public Property Let Ident(ByVal vData As String)
    mvarIdent = vData
End Property

Public Property Get Ident() As String
    Ident = mvarIdent
End Property

Public Property Let CageCode(ByVal vData As String)
    mvarCageCode = vData
End Property

Public Property Get CageCode() As String
    CageCode = mvarCageCode
End Property

Public Property Let Det(ByVal vData As String)
    mvarDet = vData
End Property

Public Property Get Det() As String
    Det = mvarDet
End Property

Public Property Let Zone(ByVal vData As String)
    mvarZone = vData
End Property

Public Property Get Zone() As String
    Zone = mvarZone
End Property

Public Property Let Rep(ByVal vData As String)
    mvarRep = vData
End Property

Public Property Get Rep() As String
    Rep = mvarRep
End Property

Public Property Let desc(ByVal vData As String)
    mvarDesc = vData
End Property

Public Property Get desc() As String
    desc = mvarDesc
End Property

Public Property Let Fourn(ByVal vData As String)
    mvarFourn = vData
End Property

Public Property Get Fourn() As String
    Fourn = mvarFourn
End Property

