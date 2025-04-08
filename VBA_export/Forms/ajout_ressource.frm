VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ajout_ressource 
   Caption         =   "Ajouter une ressource"
   ClientHeight    =   1704
   ClientLeft      =   108
   ClientTop       =   444
   ClientWidth     =   4476
   OleObjectBlob   =   "ajout_ressource.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ajout_ressource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Private Sub annuler_ressource_button_Click()
    Unload Me
End Sub

Private Sub envoyer_ressource_button_Click()
    Tâches.up = False
    If nom_ressource_tb.value <> "" Then
        Call retrieve_ressources 'recuperation tableau de ressources
            
        Dim r As Ressource
        Set r = New Ressource
        r.set_attributes nom_ressource_tb.value 'remplissage infos
        r.Display
        MsgBox r.str
        
        Call actualiser
        
        'clearing form
        Dim ctrl As Control
        For Each ctrl In Me.Controls
            If TypeName(ctrl) = "TextBox" Then ctrl.value = ""
        Next ctrl
    Else
        MsgBox "Indiquez un nom pour la ressource!"
    End If
    Tâches.up = True
End Sub

Private Sub UserForm_Click()

End Sub
