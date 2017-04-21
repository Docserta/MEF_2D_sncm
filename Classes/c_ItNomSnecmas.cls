VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_ItNomSnecmas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_ItNoms
'Collection des item de nomenclature
Private mcol As Collection

Public Function Add(No As String, _
                    Optional G03 As String = "", _
                    Optional G02 As String = "", _
                    Optional G01 As String = "", _
                    Optional Ident As String = "", _
                    Optional CageCode As String = "", _
                    Optional Det As String = "", _
                    Optional Zone As String = "", _
                    Optional Rep As String = "", _
                    Optional desc As String = "", _
                    Optional Fourn As String = "") As c_ItNomSnecma
    
    Dim objNewMember As c_ItNomSnecma
    Set objNewMember = New c_ItNomSnecma
    
    objNewMember.No = No
    objNewMember.G03 = G03
    objNewMember.G02 = G02
    objNewMember.G01 = G01
    objNewMember.Ident = Ident
    objNewMember.CageCode = CageCode
    objNewMember.Det = Det
    objNewMember.Zone = Zone
    objNewMember.Rep = Rep
    objNewMember.desc = desc
    objNewMember.Fourn = Fourn

    mcol.Add objNewMember, No

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Items() As Collection
    Set Items = mcol
End Property

Public Property Get Item(vntIndexKey As Variant) As c_ItNomSnecma
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









