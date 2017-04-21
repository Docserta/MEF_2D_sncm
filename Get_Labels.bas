Attribute VB_Name = "Get_Labels"
Option Explicit

' *****************************************************************
' * Construction de la collection des Labels
' * récupère les lablels sur toutes les planches et dans toutes les vues
' * calcul la position des label dans le quadrillage du plan
' * Création CFR le 27/03/17
' * modification le :
' *
' *
' *****************************************************************

Public Function GetLabels(mdoc, oregleH, oregleV) As c_Labels

Dim mSheets As DrawingSheets
Dim mSheet As DrawingSheet

Dim mviews As DrawingViews
Dim mview As DrawingView

Dim mTexts As DrawingTexts
Dim mText As DrawingText
Dim mLabel As DrawingText

Dim mLeads As DrawingLeaders
Dim mLead As DrawingLeader
Dim DimSheet As CoordXY 'Taille du calque
Dim PosView As CoordXY 'position de la vue dans le calque
Dim mbarre As c_ProgressBarre
Dim NbLbl As Long, cpt As Long
Dim NoPl As Integer 'numero de la planche dans l'odre de la collection des drawingSheets
Dim collabs As c_Labels
Dim LabEc As c_Label

'initialisation des variables
    Set mSheets = mdoc.Sheets
    Set collabs = New c_Labels

'Chargement de la barre de progression
    Set mbarre = New c_ProgressBarre
    mbarre.Titre = "Traitement des Labels"
    mbarre.Progression = 1
    mbarre.Affiche
    pNbEt = 6: pEtape = 1: pItem = 1: pItems = 1

    'Calcul du nombre de label pour barre de progression
    On Error Resume Next
    For Each mSheet In mSheets
        Set mviews = mSheet.Views
        If mviews.Count > 0 Then
            For Each mview In mviews
                Set mTexts = mview.Texts
                If mTexts.Count > 0 Then
                    For Each mText In mTexts
                        If InStr(1, CStr(mText.Name), "Numéro de pièce.") > 0 Then '"Numéro de pièce." ou "Balloon."
                            NbLbl = NbLbl + 1
                            mbarre.Balayage = NbLbl
                        End If
                    Next
                End If
            Next
        End If
    Next
    On Error GoTo 0
   
    For Each mSheet In mSheets 'Pour chaque planche
        NoPl = NoPl + 1
        'DimCadre = InitCadre(mSheet.PaperSize)
        Set mviews = mSheet.Views
        'Récupération de la longueur et de la hauteur du calque
        DimSheet.X = mSheet.GetPaperWidth
        DimSheet.Y = mSheet.GetPaperHeight
        If mviews.Count > 0 Then
            For Each mview In mviews 'Pour chaque vue
                Set mTexts = mview.Texts
                If mTexts.Count > 0 Then
                    For Each mText In mTexts 'Pour chaque texte
                        If InStr(1, CStr(mText.Name), "Numéro de pièce.") > 0 Then '"Numéro de pièce." ou "Balloon."
                            cpt = cpt + 1
                            Set LabEc = New c_Label
                            Set mLabel = mText
                            On Error Resume Next
                            Set mLeads = mLabel.Leaders
                            If Err.Number = 0 Then
                                mbarre.ProgressEtape (100 / NbLbl) * cpt, "Label : " & mText.Text
                                
                                LabEc.No = collabs.Count + 1
                                'LabEc.Rep = LabelRef(collabs, mTexts, mText, mview.Scale) 'si le repère est déjà présent dans la collection on ajoute "REF"
                                '#test suppression zero non significatifs
                                If IsNumeric(mText.Text) Then
                                    LabEc.Rep = CInt(mText.Text)
                                End If
                                'LabEc.Rep = mText.Text
                                LabEc.PL = NoPl
                                'LabEc.Position = PosLbl(mLabel, mview, DimCadre, DimSheet)
                                LabEc.Position = PosLbl(mLabel, mview, oregleH, oregleV, DimSheet)
                                LabEc.RefView = mview.Name
                                'Vérifie les fleches du label
                                For Each mLead In mLeads
                                    HeadLead mLead
                                Next
                                collabs.Add LabEc.No, LabEc.Rep, LabEc.PL, LabEc.Position, LabEc.RefView
                            Else
                                Err.Clear
                                On Error GoTo 0
                            End If
                        ElseIf InStr(1, CStr(mText.Name), "Numéro de pièce.") > 0 Then '"Numéro de pièce." ou "Balloon." Then
                        
                        End If
                    Next
                End If
             Next
        End If
    Next
    
'Export vers eXcel pour vérification
exportXlLbl collabs
Set GetLabels = collabs

'Libération des classes
Set collabs = Nothing
Set LabEc = Nothing
Set mbarre = Nothing

End Function

Private Sub HeadLead(mLead)
'Change la flèche du Label si elle n'est pas conforme au standard
'la flèche doit être de type catFilledCircle () si elle ne pointe pas sur une arrète
    If mLead.HeadSymbol <> 14 Then
        mLead.HeadSymbol = catFilledCircle
    End If
End Sub

Private Function LabelRef(collabs, mTexts, mText As DrawingText, ech As Double) As String
'si le repère est déjà présent dans la collection on ajoute "REF"
Dim olab As c_Label
Dim TxtRef As DrawingText
Dim DecalRef As CoordXY

    DecalRef.X = AnchorLbl(mText).X / ech
    DecalRef.Y = AnchorLbl(mText).Y / ech
    Set olab = New c_Label
    On Error Resume Next
    If collabs.Count = 0 Then
        LabelRef = mText.Text
    Else
        For Each olab In collabs.Items
            If IsNumeric(mText.Text) Then
                If olab.Rep = mText.Text Then
                    LabelRef = mText.Text & "REF"
                    'Set TxtRef = mTexts.Add("REF", mText.X + DecalRef.X, mText.Y + DecalRef.Y)
                    'TxtRef.AnchorPosition = catMiddleLeft + 20
                    'TxtRef.SetFontSize 0, 0, 4
                    'TxtRef.SetFontName 0, 0, "Times New Roman"
                    'AssociatText mText, TxtRef
                    Exit For
                Else
                    LabelRef = mText.Text
                End If
            Else
                LabelRef = mText.Text
            End If
        Next
    End If

End Function

Private Function LocateLbl(lbl As DrawingText, ech As Double, Angle As Double) As CoordXY
'Renvoi la position du label dans la vue multiplié par le facteur d'echelle de la vue
'Calcul le changement de repère pour les vues inclinées
Dim rAngle As Double
Dim AnchorBulle As CoordXY 'coordonées corrigés en fonction du point d'accrochage du texte

    AnchorBulle = AnchorLbl(lbl)

    If Sgn(Angle) < 0 Then
        Angle = Pi * 2 + Angle
    End If
    rAngle = Round(Angle, 10)
    If rAngle = 0 Then
        LocateLbl.X = lbl.X * ech
        LocateLbl.Y = lbl.Y * ech
    ElseIf rAngle > 0 And rAngle < Round(Pi / 2, 10) Then
        LocateLbl.X = (cos(rAngle) * lbl.X - Sin(rAngle) * lbl.Y) * ech
        LocateLbl.Y = (cos(rAngle) * lbl.X + Sin(rAngle) * lbl.Y) * ech
    ElseIf rAngle = Round(Pi / 2, 10) Then
        LocateLbl.X = -lbl.Y * ech
        LocateLbl.Y = lbl.X * ech
    ElseIf rAngle > Round(Pi / 2, 10) And rAngle < Round(Pi, 10) Then
        LocateLbl.X = -((cos(rAngle) * lbl.X - Sin(rAngle) * lbl.Y) * ech)
        LocateLbl.Y = (cos(rAngle) * lbl.X + Sin(rAngle) * lbl.Y) * ech
    ElseIf rAngle = Round(Pi, 10) Then
        LocateLbl.X = -lbl.X * ech
        LocateLbl.Y = -lbl.Y * ech
    ElseIf rAngle > Round(Pi, 10) And rAngle < Round(Pi + (Pi / 2), 10) Then
        LocateLbl.X = -((cos(rAngle) * lbl.X - Sin(rAngle) * lbl.Y) * ech)
        LocateLbl.Y = -((cos(rAngle) * lbl.X + Sin(rAngle) * lbl.Y) * ech)
    ElseIf rAngle = Round(Pi + (Pi / 2), 10) Then
        LocateLbl.X = lbl.Y * ech
        LocateLbl.Y = -lbl.X * ech
    ElseIf rAngle > Round(Pi + (Pi / 2), 10) And rAngle < Round(2 * Pi, 10) Then
        LocateLbl.X = (cos(rAngle) * lbl.X - Sin(rAngle) * lbl.Y) * ech
        LocateLbl.Y = -((cos(rAngle) * lbl.X + Sin(rAngle) * lbl.Y) * ech)
    Else
        MsgBox "Hola !"
    End If
End Function

Private Function AnchorLbl(lbl As DrawingText) As CoordXY
'corrige les coordonnées du point d'accrochage du label en fonctio du point d'ancrage
Dim AnchorPt As cattextanchorposition
Dim Rayon As Double

    AnchorPt = lbl.AnchorPosition - 20
    Rayon = RayBulle(lbl)
    ' Correction du X
    Select Case AnchorPt
        Case catTopLeft, catMiddleLeft, catBottomLeft
            AnchorLbl.X = AnchorLbl.X + 2 * Rayon
        Case catTopCenter, catMiddleCenter, catBottomCenter
            AnchorLbl.X = AnchorLbl.X + Rayon
        Case catTopRight, catMiddleRight, catBottomRight
            AnchorLbl.X = AnchorLbl.X
    End Select
     ' Correction du Y
    Select Case AnchorPt
        Case catTopLeft, catTopCenter, catTopRight
            AnchorLbl.Y = AnchorLbl.Y - Rayon
        Case catMiddleLeft, catMiddleCenter, catMiddleRight
        
        Case catBottomLeft, catBottomCenter, catBottomRight
            AnchorLbl.Y = AnchorLbl.Y + Rayon
    End Select
    
End Function

Private Function RayBulle(txt As DrawingText) As Double
'Calcule le rayon de la bulle d'un label en fonction du nombre de carractères
Dim nbCar As Integer

    nbCar = Len(txt.Text)
    Select Case nbCar
        Case 2
            RayBulle = (txt.TextProperties.FONTSIZE * nbCar * 1.2857) / 2 '1.372
        Case 3
            RayBulle = (txt.TextProperties.FONTSIZE * nbCar * 1.0666) / 2 ' 1.248
    End Select
End Function

Private Function PosLbl(lbl As DrawingText, Vue As DrawingView, oreglesH, oreglesV, DimCalque As CoordXY) As String
'Renvois les coordonnées A1-G9 d'un label en fonction de sa position dans la vue et de lapositin de la vue dans le calque
'En prenant en compte les lettres "interdites" (G,I,O,P,S,X,Y,Z)
'Lbl = Label
'Vue = drawingView
'DimCalque = dimension du calque

Dim CoordLbl As CoordXY 'Position du label dans la vue
Dim CoordVue As CoordXY 'Position de la vue dans le plan
Dim CoordPlan As CoordXY 'Position relative du label dans le plan
Dim NoCarreauH As String, NoCarreauV As String
Dim tRegle As c_Regle
    
    'initialisation des classes
    Set tRegle = New c_Regle

    'calcul de la position du label en fonction de l'échelle de la vue
    CoordLbl = LocateLbl(lbl, Vue.Scale, Vue.Angle)
    CoordVue.X = Vue.xAxisData
    CoordVue.Y = Vue.yAxisData
    
    CoordPlan.X = ArrondiVersInfinis(DimCalque.X - (CoordVue.X + CoordLbl.X))
    CoordPlan.Y = ArrondiVersInfinis(CoordVue.Y + CoordLbl.Y)
    'Recherche du carreau hrizontal
    For Each tRegle In oreglesH.Items
        If CoordPlan.X > tRegle.Limit Then
            NoCarreauH = tRegle.No
        End If
    Next
    'Recherche du carreau vertical
    For Each tRegle In oreglesV.Items
        If CoordPlan.Y > tRegle.Limit Then
            NoCarreauV = tRegle.No
        End If
    Next
    
    PosLbl = NoCarreauV & NoCarreauH
    
    'Libération des classes
    Set tRegle = Nothing

End Function


Private Sub exportXlLbl(collabs)
'Export le contenu de la collection vers un fichier excel pour verification
Dim oXl
Dim oWBK
Dim lig As Long
'Dim olab As c_Label
Dim olab As c_Label

    lig = 1
    Set oXl = CreateObject("EXCEL.APPLICATION")
    Set oWBK = oXl.Workbooks.Add
    oXl.Visible = True
    oWBK.ActiveSheet.cells(lig, "A").Value = "No"
    oWBK.ActiveSheet.cells(lig, "B").Value = "Rep"
    oWBK.ActiveSheet.cells(lig, "C").Value = "Position"
    oWBK.ActiveSheet.cells(lig, "D").Value = "Planche"
    oWBK.ActiveSheet.cells(lig, "E").Value = "Vue"

    For lig = 1 To collabs.Count
        Set olab = collabs.Item(lig)
        oWBK.ActiveSheet.cells(lig + 1, "A").Value = olab.No
        oWBK.ActiveSheet.cells(lig + 1, "B").Value = olab.Rep
        oWBK.ActiveSheet.cells(lig + 1, "C").Value = olab.Position
        oWBK.ActiveSheet.cells(lig + 1, "D").Value = olab.PL
        oWBK.ActiveSheet.cells(lig + 1, "E").Value = olab.RefView
    Next
    
'libération des classes
Set olab = Nothing
Set oXl = Nothing
Set oWBK = Nothing

End Sub

Private Sub AssociatText(TxtParent As DrawingText, TxtEnfant As DrawingText)
'Lie un texte à un autre
    TxtEnfant.AssociativeElement = TxtParent
End Sub

Public Function ArrondiVersInfinis(ByVal Nombre, Optional ByVal Decimales = 0)
' Fonction Arrondi vers les infinis : au chiffre supérieur
  ArrondiVersInfinis = (Fix(Nombre * 10 ^ Decimales) _
    + IIf(Nombre = Fix(Nombre * 10 ^ Decimales), 0, Sgn(Nombre))) / 10 ^ Decimales

End Function



