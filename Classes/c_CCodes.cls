VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_CCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c-CCodes
'Collection des cage Codes

Private mcol As Collection

Public Function Add(No As String, _
                    Optional Dom As String = "", _
                    Optional Nom As String = "", _
                    Optional EU As String = "", _
                    Optional US As String = "") As c_CCode
    
    Dim objNewMember As c_CCode
    Set objNewMember = New c_CCode
    
    objNewMember.No = No
    objNewMember.Dom = Dom
    objNewMember.Nom = Nom
    objNewMember.EU = EU
    objNewMember.US = US

    mcol.Add objNewMember, No

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Items() As Collection
    Set Items = mcol
End Property

Public Property Get Item(vntIndexKey As Variant) As c_CCode
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







