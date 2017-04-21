VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_ExpNom 
   Caption         =   "Export Nomenclature SNECMA"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   OleObjectBlob   =   "Frm_ExpNom.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Frm_ExpNom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Btn_Navigateur_Click()
'Recupere le fichier de la nomenclature Catia
'Dim NomComplet As String

    'Ouverture du fichier de paramètres
    pubFicNomCatia = CATIA.FileSelectionBox("Selectionner le fichier de la nomenclature Catia", "*.xls", CatFileSelectionModeOpen)
    If pubFicNomCatia = "" Then Exit Sub 'on vérifie que qque chose a bien été selectionné

    Me.TBX_nNomCatia = pubFicNomCatia

End Sub

Private Sub BtnAnnul_Click()
    Me.CB_Annule = True
    Me.Hide
End Sub

Private Sub BtnOK_Click()
    Me.CB_Annule = False
    Me.Hide
End Sub


Private Sub UserForm_Initialize()
Me.Rbt_Sncm = True
End Sub
