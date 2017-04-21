Attribute VB_Name = "Declarations_publiques"
Option Explicit

'Fonction de récupération du username
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'Version de la macro
Public Const VMacro As String = "Version 0.3 du 20/04/17"
Public Const nMacro As String = "MEF_2D_Snecma"
Public Const nPath As String = "\\srvxsiordo\xLogs\01_CatiaMacros"
Public Const nFicLog As String = "logUtilMacro.txt"

Public Const Pi As Double = 3.1415926535

Public Langue As String

Public LgCarreau As Double
Public HtCarreau As Double

'Fichier de la nomenclature Catia
Public pubFicNomCatia As String
'Fichier des Cages Codes
Public Const FicCageCode As String = "C:\CFR\Dropbox\Macros\Nomenclature-Xto-SNECMA\STD121\Fournisseurs_Standard_Safran_Aircraft_Engines.csv"
'Séparateur des fichiers CSV
Public Const SepCSV As String = ";"

'Nombre de paramètres de base dans les nomenclatures (Qte, Reference, Révision, Definition, Nomenclature, source et product description)
'Les paramètre particuliers de l'environnement débute apres
Public Const NbColPrmStd As Integer = 7
'Nombre de paramètres non modifiables (Qte, Reference)
Public Const NbColPrmNomModif As Integer = 2

'Nom des champs de nomenclature (français/Anglais)
Public nQt As String
Public nRef As String
Public nRev As String
Public nDef As String
Public nNom As String
Public nDesc As String
Public nSrce As String

Public Type CoordXY
    X As Double
    Y As Double
End Type

'Variable de la barre de progression
Public pNbEt As Long
Public pEtape As Long
Public pItem As Long
Public pItems As Long

Public DimCadre As CoordXY 'Taille des carreau de localisation dans le plan
Public Std_Form As String 'Standard des formats de plans (Snecma, CFMI, Sylvercrest, Powerjet)

Public Function InitRegleH(Frmt As String) As c_Regles
'Initialise la règle horizontale de localisation des élements dans les plans
'No = Numéro du carreau
'Limit = limite supérieure du carreau en mm
'Frmt = format de plan (Snecma A0, CFMI E, Sylvercrest A0, Popwerjet A0)
Dim tRegle As c_Regle
Dim tRegles As c_Regles
Dim Marge As Double
Dim PremCarreau As Double
Dim AutreCarreau As Double
Dim LimPrecCarreau As Double
Dim NumCarreau As String
Dim i As Integer

    'Initialisation des valeurs
    Set tRegle = New c_Regle
    Set tRegles = New c_Regles
    Select Case Frmt
        Case "Snecma"
            Marge = 10
            PremCarreau = 130
            AutreCarreau = 130
        Case "CFMI"
            Marge = 13
            PremCarreau = 135.6
            AutreCarreau = 148.6
        Case " Sylvercrest"
            Marge = 10
            PremCarreau = 130
            AutreCarreau = 130
        Case "Powerjet"
            Marge = 10
            PremCarreau = 130
            AutreCarreau = 130
        Case Else
            MsgBox "erreur dans le choix du format des format de plans (Snecma, CFMI, Sylvercrest, Powerjet)", vbCritical, "Erreur de format"
            End
    End Select
    
    'Construction de la collection
    'La marge
    tRegle.No = 1
    tRegle.Limit = Marge
    tRegles.Add tRegle.No, tRegle.Limit
    'Le premier carreau
    tRegle.No = 2
    tRegle.Limit = Marge + PremCarreau
    tRegles.Add tRegle.No, tRegle.Limit
    LimPrecCarreau = tRegle.Limit
    
    'les carreau suivants
    For i = 3 To 10 'les carreaux vont de 1 à 9 maxi
        tRegle.No = i
        tRegle.Limit = LimPrecCarreau + AutreCarreau
        tRegles.Add tRegle.No, tRegle.Limit
        LimPrecCarreau = tRegle.Limit
    Next i
    Set InitRegleH = tRegles
    
    'Libération des classe
    Set tRegle = Nothing
    Set tRegles = Nothing
End Function

Public Function InitRegleV(Frmt As String) As c_Regles
'Initialise la règle horizontale de localisation des élements dans les plans
'En prenant en compte les lettres "interdites" (G,I,O,P,S,X,Y,Z)
'No = Numéro du carreau
'Limit = limite supérieure du carreau en mm
'Frmt = format de plan (Snecma A0, CFMI E, Sylvercrest A0, Popwerjet A0)
Dim tRegle As c_Regle
Dim tRegles As c_Regles
Dim Marge As Double
Dim PremCarreau As Double
Dim AutreCarreau As Double
Dim LimPrecCarreau As Double
Dim NumCarreau As String
Dim i As Integer
Dim lettres(9) As String

    'Initialisation des valeurs
    Set tRegle = New c_Regle
    Set tRegles = New c_Regles
    
    'Lettres communes
    For i = 1 To 9
        lettres(i) = Choose(i, "A", "B", "C", "D", "E", "F", "H", "J", "K")
    Next
    Select Case Frmt
        Case "Snecma"
            Marge = 10
            PremCarreau = 120
            AutreCarreau = 120
            lettres(7) = "G"
            lettres(8) = "H"
            lettres(9) = "J"
        Case "CFMI"
            Marge = 13
            PremCarreau = 92
            AutreCarreau = 105.1
        Case " Sylvercrest"
            Marge = 10
            PremCarreau = 120
            AutreCarreau = 120
        Case "Powerjet"
            Marge = 10
            PremCarreau = 120
            AutreCarreau = 120
        Case Else
            MsgBox "erreur dans le choix du format des format de plans (Snecma, CFMI, Sylvercrest, Powerjet)", vbCritical, "Erreur de format"
            End
    End Select
    
    'Construction de la collection
    'La marge
    tRegle.No = lettres(1)
    tRegle.Limit = Marge
    tRegles.Add tRegle.No, tRegle.Limit
    'Le premier carreau
    tRegle.No = lettres(2)
    tRegle.Limit = Marge + PremCarreau
    tRegles.Add tRegle.No, tRegle.Limit
    LimPrecCarreau = tRegle.Limit
    
    'les carreau suivants
    For i = 3 To 9 'les carreaux vont de 1 à 9 maxi
        tRegle.No = lettres(i)
        tRegle.Limit = LimPrecCarreau + AutreCarreau
        tRegles.Add tRegle.No, tRegle.Limit
        LimPrecCarreau = tRegle.Limit
    Next i
    Set InitRegleV = tRegles
    
    'Libération des classe
    Set tRegle = Nothing
    Set tRegles = Nothing
End Function



