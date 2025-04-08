Attribute VB_Name = "taches_ressources"
Option Explicit
'suppression d'une tâche
Sub d_task()
    Call retrieve_tasks
    Dim answer As Integer
    
    Dim v As Integer, h As Integer
    h = TSK_horizontal_margin
    v = TSK_vertical_margin
    Tâches.up = False
    
    'si curseur bien positionné sur la liste des tâches
    If ActiveCell.Row >= v And ActiveCell.Row <= v + taches.Count And ActiveCell.Column >= h And ActiveCell.Column <= h + 5 Then
                   
        answer = MsgBox("Supprimer """ + CStr(Cells(ActiveCell.Row, h + 1).value) + """ ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmer la suppression")
               
        If answer = vbYes Then
               
            Dim id_supp As Integer 'id de la tâche supprimée
            id_supp = Cells(ActiveCell.Row, h).value
            'MsgBox taches.Count
            'MsgBox id_supp
            taches.Remove id_supp 'on retire la cible du tableau de tâches
            
            'maj prédecesseurs et afficher à nouveau les tâches restantes
            Dim i As Integer, j As Integer, preds() As String
            'MsgBox taches.Count
            For i = 1 To taches.Count
                Dim new_str As String: new_str = ""
                'If i <> id_supp Then
                'si dans les preds on a la tâche supp on supprime
                    preds = Split(taches(i).get_preds, ",")
                    For j = 0 To UBound(preds)
                        If preds(j) <> "" Then
                            If id_supp <> CInt(preds(j)) Then
                                If id_supp < CInt(preds(j)) Then
                                    preds(j) = CStr(CInt(preds(j) - 1))
                                End If
                                If j = 0 Then
                                    new_str = new_str + preds(j)
                                Else
                                    new_str = new_str + "," + preds(j)
                                End If
                            End If
                        End If
                    Next j
                    taches(i).set_preds new_str
                'End If
                taches(i).set_ID (i) 'mise à jours indice
                'MsgBox taches(i).get_Intitule
                taches(i).Display
                'MsgBox taches.Count
                'MsgBox i
                If i = taches.Count - 1 Then
                    'effacer tâches de l'affichage
                    'up = True
                    Range(Cells(taches.Count + v, h), Cells(taches.Count + v, h + 5)).Interior.Color = RGB(255, 242, 204)
                    Range(Cells(taches.Count + v, h), Cells(taches.Count + v, h + 5)).Borders.LineStyle = xlLineStyleNone
                    Range(Cells(taches.Count + v, h), Cells(taches.Count + v, h + 5)) = ""
                End If
            Next i
            
            'Décalage du tableau
            'Dim indice As Integer, i1 As Integer, i2 As Integer
            'indice = id_supp
            'i1 = ActiveCell.Row
            'i2 = i1 + 1
            'Dim s As Worksheet
            'Set s = ThisWorkbook.Worksheets("TÂCHES")
            'While i1 < taches.Count + v
                    
                '3 à 7 colonne
                's.Cells(i1, 3) = s.Cells(i2, 3)
                's.Cells(i1, 4) = s.Cells(i2, 4)
                's.Cells(i1, 5) = s.Cells(i2, 5)
                's.Cells(i1, 6) = s.Cells(i2, 6)
                's.Cells(i1, 7) = s.Cells(i2, 7)
                    
                'i1 = i1 + 1
                'i2 = i2 + 1
            'Wend
            'Suppression visuelle de la dernière case
            'Range(Cells(taches.Count + v - 1, h), Cells(taches.Count + v - 1, h + 5)).Interior.Color = RGB(255, 242, 204)
            'Range(Cells(taches.Count + v - 1, h), Cells(taches.Count + v - 1, h + 5)).Borders.LineStyle = xlLineStyleNone
            'Range(Cells(taches.Count + v - 1, h), Cells(taches.Count + v - 1, h + 5)) = ""

            
            'If id_supp = taches.Count Then
            
            'End If
            
        Else
            'up = True
        End If
    End If
    'Call update_preds(id_supp)
    Call actualiser 'maj ressources
    Tâches.up = True
End Sub

'suppression d'une ressource
Sub delete_ressource()
    Call retrieve_ressources
    Call retrieve_tasks
    Dim answer As Integer, v As Integer, h As Integer, rsc_letter As String
    v = RSC_vertical_margin
    h = RSC_horizontal_margin
    Tâches.up = False

    
    'si cursieur bien positionné sur la liste de ressources
    If ActiveCell.Row >= v And ActiveCell.Row <= v + ressources.Count And ActiveCell.Column >= h And ActiveCell.Column <= h + 2 Then
        
        answer = MsgBox("Supprimer """ + Cells(ActiveCell.Row, h + 1).value + """ ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmer la suppression")
        
        If answer = vbYes Then
            rsc_letter = Cells(ActiveCell.Row, h).value
            ressources.Remove Asc(Cells(ActiveCell.Row, h).value) - 64 'on retire la cible du tableau de tâches
            
            'gérer les tâches
            Dim j As Integer
            For j = 1 To taches.Count
                Dim ress() As String, new_str As String: new_str = ""
                ress = Split(taches(j).get_ress, ",")
                Dim k As Integer
                For k = 0 To UBound(ress)
                    If ress(k) <> rsc_letter Then
                        If k = 0 Then
                            new_str = new_str + ress(k)
                        Else
                            new_str = new_str + "," + ress(k)
                        End If
                    End If
                Next k
                taches(j).set_ress new_str
                If new_str = "" Then
                    MsgBox "Attention, plus aucune ressource pour la tâche " + CStr(taches(j).get_ID) + " : " + taches(j).get_Intitule
                End If
                taches(j).Display
            Next j
            
            
            'afficher à nouveau
            Dim i As Integer
            For i = 1 To ressources.Count
                ressources(i).set_ID (i)
                ressources(i).Display
                
                If i = ressources.Count - 1 Then
                    'effacer ressrouces
                    'up = True
                    Range(Cells(ressources.Count + v, h), Cells(ressources.Count + v, h + 2)).Interior.Color = RGB(255, 242, 204)
                    Range(Cells(ressources.Count + v, h), Cells(ressources.Count + v, h + 2)).Borders.LineStyle = xlLineStyleNone
                    Range(Cells(ressources.Count + v, h), Cells(ressources.Count + v, h + 2)) = ""
                End If
                
            Next i
            
            'Décalage du tableau
            'Dim indice As Integer, i1 As Integer, i2 As Integer
            'Dim id_supp As Integer
            'indice = id_supp
            'i1 = ActiveCell.Row
            'i2 = i1 + 1
            'Dim s As Worksheet
            'Set s = ThisWorkbook.Worksheets("TÂCHES")
            'While i1 < taches.Count + v
        
            '    s.Cells(i1, 11) = s.Cells(i2, 11)
            '    i1 = i1 + 1
            '    i2 = i2 + 1
            '
            'Wend
            
            'Suppression visuelle de la dernière case
            'Range(Cells(ressources.Count + v - 1, h), Cells(ressources.Count + v - 1, h + 2)).Interior.Color = RGB(255, 242, 204)
            'Range(Cells(ressources.Count + v - 1, h), Cells(ressources.Count + v - 1, h + 2)).Borders.LineStyle = xlLineStyleNone
            'Range(Cells(ressources.Count + v - 1, h), Cells(ressources.Count + v - 1, h + 2)) = ""
            
            
        Else
            'up = True
        End If
    End If
    Call actualiser
    Tâches.up = True
End Sub


'mise à jour de la colonne tâche pour les ressources
Sub actualiser() 'Optional ByVal test As Integer)
    
    Tâches.up = False
    
    ' 1. Refaire les tableaux de ressources et tâches via retrieve
    Call retrieve_tasks
    Call retrieve_ressources
    
    ' 2. Pour chaque tâche, on regarde ses ressources et on les stocke dans un tableau associé à la ressource? dans la variable de la ressource?
    Dim i As Integer, j As Integer, l As Integer, k As Integer, s As Worksheet, Split1() As String
    Set s = ThisWorkbook.Worksheets("TÂCHES")
    
    For i = 1 To taches.Count
        Split1 = Split(taches(i).get_ress, ",")
        For j = LBound(Split1) To UBound(Split1)
        
            For l = 1 To ressources.Count
            
                If Chr(ressources(l).get_ID + 64) = Split1(j) Then
                    If ressources(l).get_t = "" Then ' Pour ne pas mettre de virgule inutile
                        ressources(l).set_t (ressources(l).get_t & taches(i).get_ID)
                    Else
                        ressources(l).set_t (ressources(l).get_t & "," & taches(i).get_ID)
                    End If
                End If
                ' 3. Quand on a tout parcouru, on réécrit dans la colonnes tâches du tableau ressource.
                s.Cells(RSC_vertical_margin + ressources(l).get_ID - 1, RSC_horizontal_margin + 2) = ressources(l).get_t
            Next l
        Next j
    Next i
    ' Il pourrait être intéressant d'utiliser cette méthode pour vérifier si le split réfère bien à une ressource existante (si E est écrit dans les ressources des tâches mais que la ress E n'existe pas).
    Tâches.up = True
End Sub

