VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ajout_tache 
   Caption         =   "Ajouter une t�che"
   ClientHeight    =   4716
   ClientLeft      =   108
   ClientTop       =   444
   ClientWidth     =   3696
   OleObjectBlob   =   "ajout_tache.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ajout_tache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub annuler_tache_button_Click()
    Unload ajout_tache
End Sub


Private Sub envoyer_tache_button_Click()
    T�ches.up = False
    If intitule_tache_tb.value <> "" And duree_tache_tb <> "" And ressources_tb <> "" Then
        
        If IsNumeric(duree_tache_tb) Then
        
            If Not wrong_list_entry(ressources_tb, "J10:J1000") Then
                
                'If Not wrong_list_entry(predecesseurs_tb, "B10:B1000") Then
            
                    Call retrieve_tasks 'recuperation tableau de taches
                    
                    'ReDim Preserve taches(0 To UBound(taches) + 1) 'augmentation de la dimension (+1)
                    
                    Dim t As Tache
                    Set t = New Tache
                    If duree_opti_tache_tb.value <> "" Then
                        If IsNumeric(duree_opti_tache_tb) Then
                            t.set_attributes intitule_tache_tb.value, duree_tache_tb.value * 8, ressources_tb.value, predecesseurs_tb.value, duree_opti_tache_tb.value * 8 'remplissage infos
                        Else
                            MsgBox "La dur�e optimale d'une t�che doit �tre num�rique (jours)."
                        End If
                    Else
                        t.set_attributes intitule_tache_tb.value, duree_tache_tb.value * 8, ressources_tb.value, predecesseurs_tb.value
                    End If
                    t.Display
                    MsgBox t.str
                    
                    Call actualiser
                    
                    'clearing form
                    Dim ctrl As Control
                    For Each ctrl In Me.Controls
                        If TypeName(ctrl) = "TextBox" Then ctrl.value = ""
                    Next ctrl
                    
                'Else
                '    MsgBox "Erreur lors de la saisie des pr�decesseurs."
                '    predecesseurs_tb.Text = ""
                'End If
            Else
                MsgBox "Erreur lors de la saisie des ressources."
                ressources_tb.Text = ""
            End If
        Else
            MsgBox "La dur�e de t�che doit �tre num�rique (jours)."
            duree_tache_tb.Text = ""
        End If
    Else
        MsgBox "Les champs marqu�s par une �toile* sont obligatoires."
    End If
    T�ches.up = True
End Sub

Private Sub UserForm_Activate()

intitule_tache_tb.ControlTipText = "Description de la t�che en quelques mots."
label_intitule.ControlTipText = "Description de la t�che en quelques mots."

duree_tache_tb.ControlTipText = "Estimation de la dur�e de la t�che. (jours)"
duree_label.ControlTipText = "Estimation de la dur�e de la t�che. (jours)"

duree_opti_tache_tb.ControlTipText = "Estimation optimiste de la dur�e. Dans les meilleures conditions, combien de" & Chr(13) & Chr(10) & "temps faut-il pour la r�aliser? (jours)"
duree_opti_tache_label.ControlTipText = "Estimation optimiste de la dur�e. Dans les meilleures conditions, combien de" & Chr(13) & Chr(10) & "temps faut-il pour la r�aliser? (jours)"

pred_label.ControlTipText = "Liste de la/les t�che(s) devant �tre effectu�e(s) en amont (descendance directe). Exemple : 1,5,6"
predecesseurs_tb.ControlTipText = "Liste de la/les t�che(s) devant �tre effectu�e(s) en amont (descendance directe). Exemple : 1,5,6"

ressource_tache_label.ControlTipText = "Liste des ressources qui r�alisent la t�che. Exemple : D,G"
ressources_tb.ControlTipText = "Liste des ressources qui r�alisent la t�che. Exemple : D,G"

End Sub

Private Function wrong_list_entry(l As String, r As String)
    
    Dim i As Integer, res As Boolean
    res = False
    If l <> "" Then
        For i = 1 To Len(l)
            If Not i Mod 2 = 0 Then 'indice impair, on attend une ressource
                Dim rg As Range
                Set rg = Range(r).Find(Mid(l, i, 1))
                If rg Is Nothing Then
                    res = True
                End If
            Else 'indice pair, on attend une virgule
                If Mid(l, i, 1) <> "," Then
                    res = True
                End If
            End If
        Next i
    End If
    wrong_list_entry = res
End Function




