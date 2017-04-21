VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Label"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_Label

Private mvarNo As String
Private mvarRep As String
Private mvarPL As String
Private mvarPosition As String
Private mvarRefView As String

Public Property Let No(ByVal vData As String)
    mvarNo = vData
End Property

Public Property Get No() As String
    No = mvarNo
End Property

Public Property Let Rep(ByVal vData As String)
    mvarRep = vData
End Property

Public Property Get Rep() As String
    Rep = mvarRep
End Property

Public Property Let PL(ByVal vData As String)
    mvarPL = vData
End Property

Public Property Get PL() As String
    PL = mvarPL
End Property

Public Property Let Position(ByVal vData As String)
    mvarPosition = vData
End Property

Public Property Get Position() As String
    Position = mvarPosition
End Property

Public Property Let RefView(ByVal vData As String)
    mvarRefView = vData
End Property

Public Property Get RefView() As String
    RefView = mvarRefView
End Property
