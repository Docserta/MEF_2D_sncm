Attribute VB_Name = "Fonctions"
Option Explicit

Public Function check_Env(Env As String) As Boolean
'Check si l'environnement est conforme aux prérequis de lancement des macros
'Env = "Part", "Product", "CatDrawing"
On Error Resume Next
Dim mPart As PartDocument
Dim mprod As ProductDocument
Dim mDraw As DrawingDocument
check_Env = False

Select Case UCase(Env)
    Case "PART" 'Test si un CatPart est actif
        Set mPart = CATIA.ActiveDocument
        If Err.Number <> 0 Then
            MsgBox "Activez un CATPart avant de lancer cette macro !", vbCritical, "Document actif incorrect"
            Err.Clear
            End
        Else
            check_Env = True
        End If
    Case "PRODUCT" 'Test si un CatProduct est actif
        Set mprod = CATIA.ActiveDocument
        If Err.Number <> 0 Then
            MsgBox "Activez un Catproduct avant de lancer cette macro !", vbCritical, "Document actif incorrect"
            Err.Clear
            End
        Else
            check_Env = True
        End If
    Case "DRAWING" 'Test si un CatDrawing est actif
        Set mDraw = CATIA.ActiveDocument
        If Err.Number <> 0 Then
            MsgBox "Activez un CatDrawing avant de lancer cette macro !", vbCritical, "Document actif incorrect"
            Err.Clear
            End
        Else
            check_Env = True
        End If
    End Select
    
    On Error GoTo 0
End Function

Public Function FormatScale(dScale As Double) As String
'renvoi l'echelle au format x/x
    Select Case dScale
        Case 0.1
            FormatScale = "1/10"
        Case 0.125
            FormatScale = "1/8"
        Case 0.2
            FormatScale = "1/5"
        Case 0.25
            FormatScale = "1/4"
        Case 0.5
            FormatScale = "1/2"
        Case 1
            FormatScale = "1/1"
        Case 2
            FormatScale = "2/1"
        Case 4
            FormatScale = "4/1"
        Case 5
            FormatScale = "5/1"
        Case 8
            FormatScale = "8/1"
        Case 10
            FormatScale = "10/1"
        Case 20
            FormatScale = "20/1"
        Case 50
            FormatScale = "50/1"
        Case 100
            FormatScale = "100/1"
    End Select
    
End Function

Public Function FormatSource(str As String) As String
'formate le contenu du champs source
'remplace "Inconu" ou "Unknown" par une chaine vide.
'remplace les codes champs sources par une string
FormatSource = str
Select Case str
    Case "Inconnu", "Unknown"
        FormatSource = ""
    Case "Bought", catProductBought
        FormatSource = "Acheté"
    Case "Made", catProductMade
        FormatSource = "Fabriqué"
End Select
End Function

Public Sub InitLanguage()
'Configure les champs en fonction de la langue
    Langue = Language
    If Langue = "EN" Then
        nQt = "Quantity"
        nRef = "Part Number"
        nRev = "Revision"
        nDef = "Definition"
        nNom = "Nomenclature"
        nDesc = "Product Description"
        nSrce = "Source"
    Else
        nQt = "Quantité"
        nRef = "Référence"
        nRev = "Révision"
        nDef = "Définition"
        nNom = "Nomenclature"
        nDesc = "Description du produit"
        nSrce = "Source"
    End If
End Sub

Public Function Language() As String
'Détecte la langue de l'interface Catia
'Ouvre un part vierge et test le nom du "Main Body"
Dim oFolder, ofs
Dim EmptyPartFolder, EmptyPartFile
Dim oEmptyPart  As PartDocument

On Error Resume Next
Set ofs = CreateObject("Scripting.FileSystemObject")
Set oFolder = ofs.GetFolder(CATIA.Parent.Path)
Set EmptyPartFolder = ofs.GetFolder(oFolder.ParentFolder.ParentFolder.Path & "\startup\templates") ' dossier relatif des modèles vides
Set EmptyPartFile = ofs.GetFile(EmptyPartFolder.Path & "\empty.CATPart")

If Err.Number = 0 Then
    On Error GoTo 0
    Set oEmptyPart = CATIA.Documents.Open(EmptyPartFile.Path)
    If oEmptyPart.Part.MainBody.Name = "PartBody" Then
        Language = "EN"
    Else
        Language = "FR"
    End If
End If

    oEmptyPart.Close
 Set oEmptyPart = Nothing
 Set EmptyPartFile = Nothing
 Set EmptyPartFolder = Nothing
 Set oFolder = Nothing
 Set ofs = Nothing
 On Error GoTo 0

End Function

'Public Function lecture_CSV() As c_CCodes
''Ouvre le fichier Csv et construit la collection des cages codes
'Dim oCCode As c_CCode
'Dim oCCodes As c_CCodes
'Dim Line_CSV As String
'Dim Idx As Long
'Dim fs, f
'
'    Set oCCode = New c_CCode
'    Set oCCodes = New c_CCodes
'    Idx = 1
'
'    Set fs = CreateObject("scripting.filesystemobject")
'    Set f = fs.opentextfile(FicCageCode, 1, 1)
'
'    Do While Not f.AtEndOfStream
'        Line_CSV = f.ReadLine
'        Set oCCode = SplitCSV(Line_CSV, Idx)
'        Idx = Idx + 1
'
'        oCCodes.Add oCCode.No, oCCode.Dom, oCCode.Nom, oCCode.EU, oCCode.US
'    Loop
'    f.Close
'
'    Set lecture_CSV = oCCodes
'    Set oCCode = Nothing
'    Set oCCodes = Nothing
'End Function

Public Function NoDebLstPieces(objWBk) As Long
'recherche la 1ere ligne des attributs des pièces
' la ligne commence par "Liste des pièces"
    Dim NomSeparateur As String
    Dim NoLigne As Long
    NoLigne = 1
    NomSeparateur = "Liste des pièces"
    While Left(objWBk.ActiveSheet.cells(NoLigne, 1).Value, Len(NomSeparateur)) <> NomSeparateur
        NoLigne = NoLigne + 1
    Wend
    NoDebLstPieces = NoLigne

End Function

Public Function NoDebRecap(objWBk As Variant) As Long
'recherche la 1ere ligne du récapitulatif des pièces
' la ligne commence par "Nomenclature de" ou "Recapitulation of:"
    Dim NomSeparateur As String
    Dim NoLigne As Integer
    NoLigne = 1
    If Langue = "EN" Then
        NomSeparateur = "Recapitulation of:"
    ElseIf Langue = "FR" Then
        NomSeparateur = "Récapitulatif sur"
    End If
    While Left(objWBk.ActiveSheet.cells(NoLigne, 1).Value, Len(NomSeparateur)) <> NomSeparateur
        NoLigne = NoLigne + 1
    Wend
    NoDebRecap = NoLigne
End Function

Public Function NoDerniereLigne(objWBk As Variant) As Long
'recherche la dernière ligne du fichier excel
'On part du principe que 2 lignes vide indiquent la fin du fichier
Dim NoLigne As Integer, NbLigVide As Integer
    NoLigne = 1
    NbLigVide = 0
    While NbLigVide < 2
        If objWBk.ActiveSheet.cells(NoLigne, 1).Value = "" Then
            NbLigVide = NbLigVide + 1
        Else
            NbLigVide = 0
        End If
    NoLigne = NoLigne + 1
    Wend
    NoDerniereLigne = NoLigne - 2
End Function

Public Function TestEstSSE(Ligne As String) As String
'test si la ligne correspond a une entète de sous ensemble
' la ligne commence par "Nomenclature de" ou "Bill of Material"
Dim NomSeparateur As String
Dim tmpNomSSE As String

    If Langue = "EN" Then
        NomSeparateur = "Bill of Material: "
    ElseIf Langue = "FR" Then
        NomSeparateur = "Nomenclature de "
    End If
    On Error Resume Next 'Test si la chaine est vide ou inférieur a len(nomséparateur)
    tmpNomSSE = Right(Ligne, Len(Ligne) - Len(NomSeparateur))
    If Err.Number <> 0 Then
         TestEstSSE = "False"
    Else
        If Left(Ligne, Len(NomSeparateur)) = NomSeparateur Then
            TestEstSSE = tmpNomSSE
        Else
            TestEstSSE = "False"
        End If
    End If
End Function

Public Function SplitCSV(str As String, Idx As Long) As Collection
'Extrait les valeurs de la chaine str séparées par le séparateur du fichier CSV (SepCSV)
Dim oVal As Collection

    Set oVal = New Collection

    Do While InStr(1, str, SepCSV, vbTextCompare) > 0
        oVal.Add Left(str, InStr(1, str, SepCSV, vbTextCompare) - 1)
        str = Right(str, Len(str) - InStr(1, str, SepCSV, vbTextCompare))
    Loop
    oVal.Add str
    Set SplitCSV = oVal
    
    'Libération des objets
    Set oVal = Nothing
End Function

'Private Function SplitCSV(str As String, Idx As Long) As c_CCode
''Extrait les valeurs de la chaine str séparées par le séparateur du fichier CSV (SepCSV)
'Dim oMember As c_CCode
'Dim oVal As Collection
'
'    Set oMember = New c_CCode
'    Set oVal = New Collection
'
'    Do While InStr(1, str, SepCSV, vbTextCompare) > 0
'        oVal.Add Left(str, InStr(1, str, SepCSV, vbTextCompare) - 1)
'        str = Right(str, Len(str) - InStr(1, str, SepCSV, vbTextCompare))
'    Loop
'    oVal.Add str
'    On Error Resume Next 'si oVal ne contient pas le nombre de valeurs
'
'    oMember.No = Idx
'    oMember.Dom = oVal.Item(1)
'    oMember.Nom = oVal.Item(2)
'    oMember.EU = oVal.Item(3)
'    oMember.US = oVal.Item(4)
'
'    On Error GoTo 0
'Set SplitCSV = oMember
'Set oVal = Nothing
'Set oMember = Nothing
'End Function

Public Sub LogUtilMacro(ByVal mPath As String, ByVal mFic As String, ByVal mMacro As String, ByVal mModule As String, ByVal mVersion As String)
'Log l'utilisation de la macro
'Ecrit une ligne dans un fichier de log sur le serveur
'mPath = localisation du fichier de log ("\\serveur\partage")
'mFic = Nom du fichier de log ("logUtilMacro.txt")
'mMacro = nom de la macro ("NomGSE")
'mVersion = Version de la macro ("version 9.1.4")
'mModule = Nom du module ("_Info_Outillage")

Dim mDate As String
Dim mUser As String
Dim nFicLog As String
Dim nLigLog As String
Const ForWriting = 2, ForAppending = 8

    mDate = Date & " " & Time()
    mUser = ReturnUserName()
    nFicLog = mPath & "\" & mFic

    nLigLog = mDate & ";" & mUser & ";" & mMacro & ";" & mModule & ";" & mVersion

    Dim fs, f
    Set fs = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set f = fs.GetFile(nFicLog)
    If Err.Number <> 0 Then
        Set f = fs.opentextfile(nFicLog, ForWriting, 1)
    Else
        Set f = fs.opentextfile(nFicLog, ForAppending, 1)
    End If
    
    f.Writeline nLigLog
    f.Close
    On Error GoTo 0
    
End Sub

Private Function ReturnUserName() As String 'extrait d'un code de Paul, Dave Peterson Exelabo
'Renvoi le user name de l'utilisateur de la station
'fonctionne avec la fonction GetUserName dans l'entète de déclaration
    Dim Buffer As String * 256
    Dim BuffLen As Long
    BuffLen = 256
    If GetUserName(Buffer, BuffLen) Then _
    ReturnUserName = Left(Buffer, BuffLen - 1)
End Function

Public Function TestParamExist(mparams As Parameters, nParam As String) As String
'test si le paramètre passé en argument existe dans le part.
'si oui renvoi sa valeur,
'sinon la crée et lui affecte une chaine vide
Dim oParam As StrParam
On Error Resume Next
    Set oParam = mparams.Item(nParam)
If (Err.Number <> 0) Then
    ' Le paramètre n'existe pas, on le crée
    Err.Clear
    Set oParam = mparams.CreateString(nParam, "")
    oParam.Value = ""
End If
TestParamExist = oParam.Value
End Function
