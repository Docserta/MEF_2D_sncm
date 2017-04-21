Attribute VB_Name = "ExportNom"
Option Explicit

' *****************************************************************
' * Export vers un fichier Excel de la nomenclature du Drawing actif
' * recup�re la nomenclature brute g�n�r�e par catia sur le product de r�f�rence
' * Collecte les labels et calcule leur position dans le plan
' * Collecte la liste des fournisseur et recup�re les adresse dans un export SageX3
' * Cr�ation CFR le 23/03/17
' *
' *****************************************************************

Sub catmain()

Dim mdoc As Document
Dim oLigNomCatias As c_LNomCatias
Dim oLabels As c_Labels
Dim ofourns As c_CCodes
Dim oNomSnecmas As c_ItNomSnecmas
Dim oreglesH As c_Regles
Dim oreglesV As c_Regles

'Log de l'utilisation de la macro
LogUtilMacro nPath, nFicLog, nMacro, "ExportNom", VMacro

    'Check environnement
    'Test si un CatDrawing est actif
    If check_Env("Drawing") Then Set mdoc = CATIA.ActiveDocument
    'Configure les champs en fonction de la langue
    InitLanguage
    
    'Chargement du formulaire
    Load Frm_ExpNom
    Frm_ExpNom.Show
    'Bouton "annuler" choisi, on decharge le formulaire et on quite
    If Frm_ExpNom.CB_Annule Then
        Unload Frm_ExpNom
        Exit Sub
    End If
    'Stockage des infos du formulaire
    If Frm_ExpNom.Rbt_Sncm = True Then
        Std_Form = "Snecma"
    ElseIf Frm_ExpNom.Rbt_CFMI = True Then
        Std_Form = "CFMI"
    ElseIf Frm_ExpNom.Rbt_Sylver = True Then
        Std_Form = "Sylvercrest"
    ElseIf Frm_ExpNom.Rbt_Power = True Then
        Std_Form = "Powerjet"
    End If
    Unload Frm_ExpNom

    'Initialisation des variables
    Set oLigNomCatias = New c_LNomCatias
    Set oLabels = New c_Labels
    Set oNomSnecmas = New c_ItNomSnecmas
    Set oreglesH = InitRegleH(Std_Form)
    Set oreglesV = InitRegleV(Std_Form)
    
    'Collecte des Label
    Set oLabels = GetLabels(mdoc, oreglesH, oreglesV)

    'Collecte des adresses fournisseur
    Set ofourns = GetFourns
    
    'Construction de la classe des lignes de nomenclature catia
    Set oLigNomCatias = GetNomCatia(pubFicNomCatia)

    'Construction de la classe des lignes de nomenclature Snecma
    Set oNomSnecmas = GetNomSnecma(oLigNomCatias, oLabels)
    
    'Ajour des adresses fournisseurs
     AjoutFourn oNomSnecmas, ofourns

    'Export vers eXcel pour v�rification
    exportXlNomSnecma oNomSnecmas

End Sub

Public Function GetNomCatia(cibleNomCatia As String) As c_LNomCatias
'Formate le fichier excel de la nomenclature g�n�r� par Catia
'Regoupe les ensembles
'Regroupe les d�tails

Dim objexcel
Dim objWBkNomCatia
Dim LigActive As Long
Dim ColActive As Integer
Dim NoLigFinNom As Long
Dim NoLigFinEns As Long
Dim NoLigDebDet As Long
Dim NoLigDebSSE As Long
Dim NomSSE As String
Dim cLigNomEn As c_LNomCatia
Dim cLigNomENs As c_LNomCatias
Dim cAttribut As c_Attribut
Dim cAttributs As c_Attributs
Dim cAttributEnv As c_Attribut
Dim cAttributEnvs As c_Attributs 'Collections des attributs de l'environnement (ex GSE)
Dim i As Long
Dim pEtape As Long, pNbEt As Long, pItem As Long, pItems As Long
Dim mbarre As c_ProgressBarre
Dim pos As Integer

'Initialisation des classes
    Set cAttributEnv = New c_Attribut
    Set cAttributEnvs = New c_Attributs
    Set cLigNomEn = New c_LNomCatia
    Set cLigNomENs = New c_LNomCatias
    
'Chargement de la barre de progression
    Set mbarre = New c_ProgressBarre
    mbarre.Titre = "Extraction des param�tres"
    mbarre.Progression = 1
    mbarre.Affiche
    pNbEt = 6: pEtape = 2: pItem = 1: pItems = 1

'Initialisation des objets Excel
    'Ouverture de la nom g�n�r�e par Catia
    Set objexcel = CreateObject("EXCEL.APPLICATION")
    Set objWBkNomCatia = objexcel.Workbooks.Open(CStr(cibleNomCatia))
    objexcel.Visible = True

    'Recherche la position des lignes dans le fichier Excel
    NoLigFinNom = NoDerniereLigne(objWBkNomCatia)
    NoLigFinEns = NoDebRecap(objWBkNomCatia)
    NoLigDebDet = NoLigFinEns + 4
    NoLigDebSSE = 3
    
    'collecte de la liste des propri�t�s sp�cifiques � l'environnement (nom des attributs)
    LigActive = 4
    ColActive = NbColPrmStd + 1 'Premi�re colonne apr�s les attributs standards
    While objWBkNomCatia.ActiveSheet.cells(LigActive, ColActive).Value <> ""
        cAttributEnv.Nom = objWBkNomCatia.ActiveSheet.cells(LigActive, ColActive).Value
        cAttributEnv.Ordre = ColActive
        cAttributEnvs.Add cAttributEnv.Nom, cAttributEnv.Ordre
        ColActive = ColActive + 1
    Wend
           
    'Collecte des sous ensembles
    LigActive = 5

    'Creation de la liste des SSe (reference)
    pEtape = 3: pItem = 1: pItems = NoLigFinEns - LigActive
    While LigActive < NoLigFinEns
        mbarre.CalculProgression pEtape, pNbEt, pItem, pItems, "Cr�ation de la liste des sous ensembles"
        NomSSE = TestEstSSE(objWBkNomCatia.ActiveSheet.cells(LigActive, 1).Value)
        If NomSSE <> "False" Then
            cLigNomEn.ref = NomSSE
            'Ajout du sous ensemble a la collection
            cLigNomENs.Add cLigNomEn.ref
        End If
        LigActive = LigActive + 1
        pItem = pItem + 1
    Wend
    
'collecte des propri�t�s de chaque SSE
    LigActive = 5
    pEtape = 4: pItem = 1: pItems = NoLigFinEns - LigActive
    While LigActive < NoLigFinEns
        mbarre.CalculProgression pEtape, pNbEt, pItem, pItems, "collecte des propri�t�s de chaque sous ensemble"
        For Each cLigNomEn In cLigNomENs.Items
            If objWBkNomCatia.ActiveSheet.cells(LigActive, 2).Value = cLigNomEn.ref Then
                Set cAttributs = New c_Attributs
                Set cAttribut = New c_Attribut
                'Quantit�
                pos = 1 'cAttributEnvs.Item(nQt).Ordre
                cLigNomEn.Qte = objWBkNomCatia.ActiveSheet.cells(LigActive, pos).Value
                ' R�vision
                pos = 3 ' cAttributEnvs.Item(nRev).Ordre
                cLigNomEn.Rev = objWBkNomCatia.ActiveSheet.cells(LigActive, pos).Value
                ' Definition
                pos = 4 ' cAttributEnvs.Item(nDef).Ordre
                cLigNomEn.Def = objWBkNomCatia.ActiveSheet.cells(LigActive, pos).Value
                ' Nomenclature
                pos = 5 'cAttributEnvs.Item(nNom).Ordre
                cLigNomEn.Nom = objWBkNomCatia.ActiveSheet.cells(LigActive, pos).Value
                ' source
                pos = 6 'cAttributEnvs.Item(nSrce).Ordre
                cLigNomEn.Source = FormatSource(objWBkNomCatia.ActiveSheet.cells(LigActive, pos).Value)
                ' product description
                pos = 7 'cAttributEnvs.Item(nDesc).Ordre
                cLigNomEn.desc = objWBkNomCatia.ActiveSheet.cells(LigActive, pos).Value
                cLigNomEn.Comp = "E"
                'collecte des attributs li�s a l'environnement
                For Each cAttributEnv In cAttributEnvs.Items
                    cAttribut.Nom = cAttributEnv.Nom
                    cAttribut.Ordre = cAttributEnv.Ordre
                    cAttribut.Valeur = objWBkNomCatia.ActiveSheet.cells(LigActive, cAttributEnv.Ordre).Value
                    cAttributs.Add cAttribut.Nom, cAttribut.Ordre, cAttribut.Valeur
                Next
                'Ajout de la collection des attributs a la ligne de nomenclature
                cLigNomEn.Attributs = cAttributs
                'vidage de la collection des attributs
                Set cAttribut = Nothing
            End If
        Next
        LigActive = LigActive + 1
        pItem = pItem + 1
    Wend
    Set cLigNomEn = Nothing
    
'Collecte des d�tails
    LigActive = NoLigDebDet + 1
    Set cLigNomEn = New c_LNomCatia
    pEtape = 5: pItem = 1: pItems = NoLigFinNom - LigActive
    While LigActive < NoLigFinNom
        mbarre.CalculProgression pEtape, pNbEt, pItem, pItems, "Collecte des propri�t�s des d�tails"
        Set cAttributs = New c_Attributs
        Set cAttribut = New c_Attribut
        'Collecte de la valeur des attributs Standards
        'Part Number
        cLigNomEn.ref = objWBkNomCatia.ActiveSheet.cells(LigActive, 2).Value
        ' Quantit�
        cLigNomEn.Qte = objWBkNomCatia.ActiveSheet.cells(LigActive, 1).Value
        ' R�vision
        cLigNomEn.Rev = objWBkNomCatia.ActiveSheet.cells(LigActive, 3).Value
        ' Definition
        cLigNomEn.Def = objWBkNomCatia.ActiveSheet.cells(LigActive, 4).Value
        ' Nomenclature
        cLigNomEn.Nom = objWBkNomCatia.ActiveSheet.cells(LigActive, 5).Value
        ' source
        cLigNomEn.Source = FormatSource(objWBkNomCatia.ActiveSheet.cells(LigActive, 6).Value)
        ' product description
        cLigNomEn.desc = objWBkNomCatia.ActiveSheet.cells(LigActive, 7).Value
        cLigNomEn.Comp = "D"
        'collecte de la valeur des attributs sp�cifiques a l'environnement
        For Each cAttributEnv In cAttributEnvs.Items
            cAttribut.Nom = cAttributEnv.Nom
            cAttribut.Ordre = cAttributEnv.Ordre
            cAttribut.Valeur = objWBkNomCatia.ActiveSheet.cells(LigActive, cAttributEnv.Ordre).Value
            cAttributs.Add cAttribut.Nom, cAttribut.Ordre, cAttribut.Valeur
        Next
        'Ajout de la collection des attributs a la ligne de nomenclature
        cLigNomEn.Attributs = cAttributs
        'vidage de la collection des attributs
        Set cAttribut = Nothing
        LigActive = LigActive + 1
        cLigNomENs.Add cLigNomEn.ref, cLigNomEn.Comp, cLigNomEn.Qte, cLigNomEn.Rev, cLigNomEn.Def, cLigNomEn.Nom, cLigNomEn.Source, cLigNomEn.desc, cLigNomEn.Attributs
        pItem = pItem + 1
    Wend

Set GetNomCatia = cLigNomENs

'Lib�ration des classes
'Fermeture du fichier excel de nomenclature catia
    objWBkNomCatia.Close
Set cLigNomEn = Nothing
Set cLigNomENs = Nothing
Set cAttributEnv = Nothing
Set cAttributEnvs = Nothing
Set mbarre = Nothing

End Function

Private Sub AjoutFourn(ByRef oNomSnecmas, ByRef ofourns)
'Ajoute les adresses des fournisseur a la collection
Dim oFourn As c_CCode
Dim oNomSnecma As c_ItNomSnecma

    'Initialisation des classes
    Set oNomSnecma = New c_ItNomSnecma
    Set oFourn = New c_CCode

    'Pour chaque ligne de nomenclature on recherche le Fournisseur associ�
    For Each oNomSnecma In oNomSnecmas.Items
        
        Set oFourn = SearchFirstFourn(oNomSnecma.desc, ofourns)
        If oFourn Is Nothing Then
        Else
            oNomSnecma.Fourn = Replace(oFourn.US, "$", Chr(10), 1, , vbTextCompare)
        End If
    Next
    
    'Lib�ration des objets
    Set oNomSnecma = Nothing
    Set oFourn = Nothing

End Sub

Public Function GetNomSnecma(oLigNomCatias, oLabels) As c_ItNomSnecmas
'R�cup�ration des attributs de la nomenclature Catia
'Association des labels aux item de la nomenclature
Dim olignom As c_LNomCatia
Dim oLabel As c_Label
Dim oNomSnecma As c_ItNomSnecma
Dim oNomSnecmas As c_ItNomSnecmas
Dim i As Long
Dim mbarre As c_ProgressBarre

    'Initialisation des classes
    Set oNomSnecma = New c_ItNomSnecma
    Set oNomSnecmas = New c_ItNomSnecmas
    Set olignom = New c_LNomCatia
    
    'Chargement de la barre de progression
    Set mbarre = New c_ProgressBarre
    mbarre.Affiche
    pNbEt = 6: pEtape = 5: pItem = 1: pItems = 1
    
    'Pour chaque ligne de nomenclature on recherche le label associ�
    For Each olignom In oLigNomCatias.Items
        pItems = oLigNomCatias.Count
        mbarre.CalculProgression pEtape, pNbEt, pItem, pItems, "G�n�ration de la nomenclature SNECMA"
        i = i + 1
        pItem = i
        Set oLabel = SearchFirstLbl(olignom.ref, oLabels)
        oNomSnecma.G03 = olignom.Qte
        oNomSnecma.Ident = olignom.Nom
        oNomSnecma.Rep = olignom.ref
        oNomSnecma.Fourn = olignom.desc
        'Le nom du fournisseur est enfouis dans la description on l'extraira plus tard
        oNomSnecma.Fourn = olignom.desc
        If oLabel Is Nothing Then
            oNomSnecma.desc = olignom.desc
        Else
            oNomSnecma.Det = listOfPlanche(olignom.ref, oLabels)
            oNomSnecma.Zone = oLabel.Position & "-" & oLabel.PL
            'Concat�nation des d�signation FR & EN
            oNomSnecma.desc = olignom.Attributs.Item("x_designation").Valeur
            oNomSnecma.desc = oNomSnecma.desc & Chr(10)
            oNomSnecma.desc = oNomSnecma.desc & olignom.Attributs.Item("x_designation anglais").Valeur
        End If
        oNomSnecmas.Add CStr(i), oNomSnecma.G03, oNomSnecma.G02, oNomSnecma.G01, oNomSnecma.Ident, oNomSnecma.CageCode, oNomSnecma.Det, oNomSnecma.Zone, oNomSnecma.Rep, oNomSnecma.desc, oNomSnecma.Fourn
        'RAZ de la ligne de nomsnecma
        Set oNomSnecma = New c_ItNomSnecma
    Next
    
    Set GetNomSnecma = oNomSnecmas

    'Liberation des classes
    Set oNomSnecma = Nothing
    Set oNomSnecmas = Nothing
    Set olignom = Nothing

End Function

Private Sub exportXlNomSnecma(oNomSnecmas)
'Export le contenu de la collection vers un fichier excel pour verification
Dim oXl
Dim oWBK
Dim lig As Long
Dim oNomSnecma As c_ItNomSnecma
    lig = 1
    Set oXl = CreateObject("EXCEL.APPLICATION")
    Set oWBK = oXl.Workbooks.Add
    oXl.Visible = True
    oWBK.ActiveSheet.cells(lig, "A").Value = "No ligne"
    oWBK.ActiveSheet.cells(lig, "B").Value = "G01"
    oWBK.ActiveSheet.cells(lig, "C").Value = "G02"
    oWBK.ActiveSheet.cells(lig, "D").Value = "G03"
    oWBK.ActiveSheet.cells(lig, "E").Value = "Ident"
    oWBK.ActiveSheet.cells(lig, "F").Value = "CageCode"
    oWBK.ActiveSheet.cells(lig, "G").Value = "Det"
    oWBK.ActiveSheet.cells(lig, "H").Value = "Zone"
    oWBK.ActiveSheet.cells(lig, "I").Value = "Rep"
    oWBK.ActiveSheet.cells(lig, "J").Value = "Description"
    oWBK.ActiveSheet.cells(lig, "K").Value = "Fournisseur"
    
    For lig = 1 To oNomSnecmas.Count
        Set oNomSnecma = oNomSnecmas.Item(lig)
        oWBK.ActiveSheet.cells(lig + 1, "A").Value = oNomSnecma.No
        oWBK.ActiveSheet.cells(lig + 1, "B").Value = oNomSnecma.G01
        oWBK.ActiveSheet.cells(lig + 1, "C").Value = oNomSnecma.G02
        oWBK.ActiveSheet.cells(lig + 1, "D").Value = oNomSnecma.G03
        oWBK.ActiveSheet.cells(lig + 1, "E").Value = oNomSnecma.Ident
        oWBK.ActiveSheet.cells(lig + 1, "F").Value = oNomSnecma.CageCode
        oWBK.ActiveSheet.cells(lig + 1, "G").Value = "'" & oNomSnecma.Det
        oWBK.ActiveSheet.cells(lig + 1, "H").Value = oNomSnecma.Zone
        oWBK.ActiveSheet.cells(lig + 1, "I").Value = oNomSnecma.Rep
        oWBK.ActiveSheet.cells(lig + 1, "J").Value = oNomSnecma.desc
        oWBK.ActiveSheet.cells(lig + 1, "K").Value = oNomSnecma.Fourn
    Next

'lib�ration des classes
Set oNomSnecma = Nothing
Set oXl = Nothing
Set oWBK = Nothing

End Sub

Private Function SearchFirstLbl(ref As String, oLabels) As c_Label
'Recherche la premi�re occurence du label dans la collection
Dim otemplab As c_Label
    Set otemplab = New c_Label
    For Each otemplab In oLabels.Items
        If otemplab.Rep = ref Then
            Set SearchFirstLbl = otemplab
            Exit For
        End If
    Next
End Function

Private Function SearchFirstFourn(desc As String, ofourns) As c_CCode
'Recherche la premi�re occurence du fournisseur dans la collection
Dim oTempfour As c_CCode
    
    Set oTempfour = New c_CCode
    For Each oTempfour In ofourns.Items
        If oTempfour.Nom = desc Then
            Set SearchFirstFourn = oTempfour
            Exit For
        End If
    Next
    
End Function

Private Function listOfPlanche(ref As String, oLabels) As String
'collecte la liste des planches dans lesquelles apparait le rep�re
'ne renvois que les N� de planche a partir de la seconde
Dim otemplab As c_Label
Dim NbOcc As Integer
    listOfPlanche = ""
    Set otemplab = New c_Label
    For Each otemplab In oLabels.Items
        If otemplab.Rep = ref Then
            If NbOcc = 1 Then
                listOfPlanche = otemplab.PL
            ElseIf NbOcc > 1 Then
                listOfPlanche = listOfPlanche & "-" & otemplab.PL
            End If
            NbOcc = NbOcc + 1
        End If
    Next
End Function

