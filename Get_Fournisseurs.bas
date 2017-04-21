Attribute VB_Name = "Get_Fournisseurs"
Option Explicit

' *****************************************************************
' *
' * Ouvre le fichier Csv et construit la collection des cages codes
' * Création CFR le 10/02/2017
' * Dernière modification le :
' *****************************************************************

Public Function GetFourns() As c_CCodes

Dim oCCodes As c_CCodes
Dim oCCode As c_CCode
Dim Line_CSV As String
Dim Idx As Long
Dim fs, f
Dim oVal As Collection


    'Initialisation des variables
    Set oCCodes = New c_CCodes
    Set oCCode = New c_CCode
    Idx = 1
    Set fs = CreateObject("scripting.filesystemobject")
    Set f = fs.opentextfile(FicCageCode, 1, 1)

    Do While Not f.AtEndOfStream
        Set oVal = New Collection
        Line_CSV = f.ReadLine
        Set oVal = SplitCSV(Line_CSV, Idx)
        
        On Error Resume Next 'si oVal ne contient pas le nombre de valeurs

        oCCode.No = Idx
        oCCode.Dom = oVal.Item(1)
        oCCode.Nom = oVal.Item(2)
        oCCode.EU = oVal.Item(3)
        oCCode.US = oVal.Item(4)

        On Error GoTo 0
        
        Idx = Idx + 1
        oCCodes.Add oCCode.No, oCCode.Dom, oCCode.Nom, oCCode.EU, oCCode.US
    Loop
    f.Close

    Set GetFourns = oCCodes
    
    'Libération des oblets
    Set f = Nothing
    Set fs = Nothing
    Set oCCode = Nothing
    Set oCCodes = Nothing

End Function



