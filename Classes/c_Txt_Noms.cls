VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Txt_Noms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_Txt_Noms
Private mcol As Collection

Public Function Add(No As String, _
                    Optional Nom As String = "", _
                    Optional Texte As String = "", _
                    Optional Police As String = "", _
                    Optional Coul As Long = 0, _
                    Optional X As Double = 0, _
                    Optional Y As Double = 0) As c_Txt_Nom
    
    Dim objNewMember As c_Txt_Nom
    Set objNewMember = New c_Txt_Nom
    
    objNewMember.No = No
    objNewMember.Nom = Nom
    objNewMember.Texte = Texte
    objNewMember.Police = Police
    objNewMember.Coul = Coul
    objNewMember.X = X
    objNewMember.Y = Y

    mcol.Add objNewMember, No

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Items() As Collection
    Set Items = mcol
End Property

Public Property Get Item(vntIndexKey As Variant) As c_Txt_Nom
  Set Item = mcol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mcol.Count
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







