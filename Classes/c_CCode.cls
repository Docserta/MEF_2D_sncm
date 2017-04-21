VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_CCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_CCode
'Cage Codes

Private mvarNo As String
Private mvarDom As String
Private mvarNom As String
Private mvarEU As Integer
Private mvarUS As String

Public Property Let No(ByVal vData As String)
    mvarNo = vData
End Property

Public Property Get No() As String
    No = mvarNo
End Property

Public Property Let Dom(ByVal vData As String)
    mvarDom = vData
End Property

Public Property Get Dom() As String
    Dom = mvarDom
End Property

Public Property Let Nom(ByVal vData As String)
    mvarNom = vData
End Property

Public Property Get Nom() As String
    Nom = mvarNom
End Property

Public Property Let EU(ByVal vData As Integer)
    mvarEU = vData
End Property

Public Property Get EU() As Integer
    EU = mvarEU
End Property

Public Property Let US(ByVal vData As String)
    mvarUS = vData
End Property

Public Property Get US() As String
    US = mvarUS
End Property


