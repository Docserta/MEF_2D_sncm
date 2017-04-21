VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Labels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_Labels
Private mcol As Collection

Public Function Add(No As String, _
                    Optional Rep As String = "", _
                    Optional PL As String = "", _
                    Optional Position As String = "", _
                    Optional RefView As String = "") As c_Label
    
    Dim objNewMember As c_Label
    Set objNewMember = New c_Label
    
    objNewMember.No = No
    objNewMember.Rep = Rep
    objNewMember.PL = PL
    objNewMember.Position = Position
    objNewMember.RefView = RefView

    mcol.Add objNewMember

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Items() As Collection
    Set Items = mcol
End Property

Public Property Get Item(vntIndexKey As Variant) As c_Label
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





