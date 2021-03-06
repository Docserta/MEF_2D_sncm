VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Attributs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_Attributs
'Classe des attributs personali�s sur les part et products
'l'ordre correspond � la colonne de l'attribut dans le fichier de nomenclature Catia
'La classe ne devant pas contenir de doublon, elle est index�e
Private mcol As Collection

Public Function Add(Nom As String, _
            Optional Ordre As Integer = 100, _
            Optional Valeur As String = "") As c_Attribut
    
    Dim objNewMember As c_Attribut
    Set objNewMember = New c_Attribut
    
    objNewMember.Nom = Nom
    objNewMember.Ordre = Ordre
    objNewMember.Valeur = Valeur

    mcol.Add objNewMember, Nom

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As c_Attribut
  Set Item = mcol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mcol.Count
End Property

Public Property Get Items() As Collection
    Set Items = mcol
End Property

Public Sub Remove(vntIndexKey As Variant)
    mcol.Remove vntIndexKey
End Sub

Private Sub Class_initialize()
    Set mcol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcol = Nothing
End Sub




