Attribute VB_Name = "buffers"

'calcul de la durée du buffer pour chaque chaîne, préparation au positionnement dans le GANTT
'"cc" : chaîne critique, "sc" : tableau de chaînes secondaires
Sub generate_buffers(cc As Collection, sc As Collection, Optional alert As Integer = 0)
    
    Dim i As Integer, j As Integer, t As Tache, sum As Integer
    Dim indice_max As Integer: indice_max = cc.Count
    For i = 1 To cc.Count
        sum = sum + (cc(i).get_duree_nominale - cc(i).get_duree) 'somme des durées nominale - opti
        'sum = sum + cc(i).get_duree
        If cc(i).get_fin > cc(indice_max).get_fin Then
            indice_max = i
        End If
    Next i
    
    'création d'une tâche qui joue le rôle de la chaîne critique
    Set t = New Tache
    'son prédecesseur c'est la tâche qui finit le plus tard (dernière tâche pas nécessairement la dernière du projet)
    't.set_attributes "Buffer chaîne critique", sum / 2, "1", CStr(cc(indice_max).get_ID)
    t.set_attributes "Buffer chaîne critique", sum, "1", CStr(cc(indice_max).get_ID)
    t.set_type 4
    taches.Add t
    ThisWorkbook.Worksheets("LOGS").Cells(15, 16).value = sum ' / 2 'on enregistre la longueur du buffer
    
    
    Dim chains As Collection
    Set chains = New Collection
    
    'For i = 1 To sc.Count
    '    For j = 1 To sc(i).Count
    '        MsgBox sc(i)(j).get_Intitule
    '    Next j
    'Next i
    
    'pour les chaînes secondaires, le but ça va être de d'abord les séparer (on sait comment ordo procède pour les ajouter au tableau)
    For i = 1 To sc.Count
        If sc(i).Count > 0 Then
            Dim chain As Collection
            Dim res As Boolean: res = False
            
            Set chain = New Collection
            chain.Add sc(i)(1)
            'MsgBox sc(i)(1).get_Intitule
            chains.Add chain
            
            'Dim indice_stop As Integer: indice_stop = 1
            'sc(i).Remove 1
            For j = 2 To sc(i).Count
                Dim preds() As String, k As Integer
                preds() = Split(sc(i)(j - 1).get_preds, ",") 'on récup les preds de l'élément d'avant
                'MsgBox "oh mon pote : " + sc(i)(j).get_preds
                For k = 0 To UBound(preds)
                    'MsgBox preds(k)
                    If sc(i)(j).get_ID = CInt(preds(k)) Then 'si j'en fais partie alors faut pas passer à la chaine suivante
                        'MsgBox "at least 1ce"
                        res = True
                    End If
                Next k
                If res = False Then
                    Set chain = New Collection
                    chain.Add sc(i)(j)
                    chains.Add chain
                End If
            Next j
            For j = 1 To chains.Count
                'MsgBox chains(j)(i).get_ID
                Call remove_task_by_id(chains(j)(i).get_ID, sc(i))
            Next j
        End If
    Next i
    
    'juste au dessus on a set les premières tâches de chaque chaîne
    
        For i = 1 To sc.Count
        If sc(i).Count > 0 Then
            While (sc(i).Count > 0)
            
                'MsgBox "oh " + sc(i)(1).get_Intitule
                'If sc(i)(j).get_preds <> "" Then
                Dim ante As Collection, p As Integer
                res = False
                For p = 1 To chains.Count
                    'Set t = sc(i)(1)
                    'MsgBox chains(p)(chains(p).Count).get_Intitule
                    If res = False Then
                        Set t = sc(i)(1)
                        Set ante = antecedants(t, chains(p))
                        If ante.Count > 0 Then 'dans ce cas la tache appartient à la chaine p
                            res = True
                            chains(p).Add sc(i)(1)
                            sc(i).Remove 1
                        End If
                    End If
                Next p
            Wend
        End If
    Next i
    
    For i = 1 To chains.Count
    
        'on veut écrire la chaine sur la fiche de logs
        Dim log As String: log = ""
    
        Dim s As Integer: s = 0
        Dim w As Integer
        For j = 1 To chains(i).Count
            w = chains(i).Count - j + 1
            'somme des marges pour buffer
            s = s + (chains(i)(j).get_duree_nominale - chains(i)(j).get_duree)
            If j = chains(i).Count Then
                log = log + CStr(chains(i)(w).get_ID)
            Else
                log = log + CStr(chains(i)(w).get_ID) + ","
            End If
        Next j
        ThisWorkbook.Worksheets("LOGS").Cells(15 + i, 15).value = log
        ThisWorkbook.Worksheets("LOGS").Cells(15 + i, 16).value = s '/ 2
        
        
        'création d'une tâche qui joue le rôle de buffer
        Set t = New Tache
        't.set_attributes "Buffer chaîne " + CStr(chains(i)(1).get_ID), s / 2, CStr(i + 1), CStr(chains(i)(1).get_ID) 'prédecesseur = dernière tâche de la chaîne
        t.set_attributes "Buffer chaîne " + CStr(chains(i)(1).get_ID), s, CStr(i + 1), CStr(chains(i)(1).get_ID)  'prédecesseur = dernière tâche de la chaîne
        t.set_type 4
        taches.Add t
        
        Dim ctr As Integer: ctr = 0
        
        If alert = 0 Then
            'on veut casser le lien de prédecesseur entre la tache de chaine critique et la première tâche de la chaine
            For j = 1 To cc.Count
                preds = Split(taches(cc(j).get_ID).get_preds, ",")
                Dim countr As String: countr = "" 'retenir combien de caractère de la chaine sont des id de tâche
                For k = 0 To UBound(preds)
                    countr = countr + preds(k)
                    If chains(i)(1).get_ID = CInt(preds(k)) Then '
                        'MsgBox taches(cc(j).get_ID).get_preds
                        'MsgBox k + ctr
                        If k + ctr = UBound(preds) + ctr Then
                            
                            'MsgBox "end " + CStr(taches(cc(j).get_ID).get_ID)
                            
                            
                            If k = 0 Then
                                'MsgBox "moi?"
                                taches(cc(j).get_ID).set_preds (vbNullString)
                            End If
                            'MsgBox "ici"
                            'MsgBox taches(cc(j).get_ID).get_preds
                            'taches(cc(j).get_ID).set_preds (Replace(taches(cc(j).get_ID).get_preds, UCase("," + CStr(chains(i)(1).get_ID)), "", , 1))
                            'MsgBox CStr(CInt(Len(taches(cc(j).get_ID).get_preds)) - CInt(Len(preds(k))) - 1)
                            Dim test As Long
                            test = Len(taches(cc(j).get_ID).get_preds) - Len(preds(k)) - 1
                            'MsgBox test
                            taches(cc(j).get_ID).set_preds Left(taches(cc(j).get_ID).get_preds, test)
                            'MsgBox taches(cc(j).get_ID).get_preds
                            
                            ctr = ctr + 1
                        Else
                            'MsgBox "hihi oups"
                            'MsgBox preds(k)
                            Dim right_part As String, left_part As String
                            left_part = Left(taches(cc(j).get_ID).get_preds, Len(countr) - Len(preds(k)) + k)
                            'MsgBox left_part
                            'MsgBox right_part
                            right_part = Right(taches(cc(j).get_ID).get_preds, Len(taches(cc(j).get_ID).get_preds) - Len(left_part) - Len(preds(k)) - 1)
                            taches(cc(j).get_ID).set_preds (left_part + right_part)
                            
                            ctr = ctr + 1
                        End If
                    End If
                Next k
                'MsgBox taches(cc(j).get_ID).get_preds
            Next j
        End If
        
        'il faut que le buffer soit prédécesseur de la tâche critique à la source
        Dim a As Integer
        Set t = chains(i)(1)
        'MsgBox antecedants(t, cc)(1).get_ID
        a = antecedants(t, cc)(1).get_ID
        'MsgBox "persuader " + CStr(taches.Count)
        taches(a).set_preds (taches(a).get_preds + "," + CStr(taches.Count))
        
        
    Next i
    
End Sub


'calcule le pourcentage de buffer consommé, le note en log pour mettre à jour le graph
'"pos_actuelle" : la date concernée, "col" : la colonne associée au buffer en logs
Sub consume_buffers(pos_actuelle As Integer, col As Integer)
    Dim d As Integer, chains As Collection, s As Worksheet, l As Worksheet
    d = colonne_date_actuelle
    Set chains = retrieve_chains()
    Set s = ThisWorkbook.Worksheets("GANTT")
    Set l = ThisWorkbook.Worksheets("LOGS_FV_CHART")
    
    'ThisWorkbook.Worksheets("DASHBOARD").Cells(GANTT_vertical_margin - 3, 24).value = s.Cells(GANTT_vertical_margin - 2, d).value
    ThisWorkbook.Worksheets("DASHBOARD").Cells(3, 24).value = s.Cells(1, 16).value
    Dim i As Integer
    For i = 1 To chains.Count
        'Dim avancement As Integer: avancement = 0
        'Dim duree_totale_chaine As Integer: duree_totale_chaine = 0
        
        l.Cells(16, 4 * i + 1).value = 0
        l.Cells(16, 4 * i + 2).value = 0
        l.Cells(16, 4 * i + 3).value = 1
        
        Dim debut_chaine As Integer
        debut_chaine = chains(i)(1).get_debut
        
        'Dim pos_actuelle As Integer
        'pos_actuelle = (d - GANTT_horizontal_margin) * 2 + 2
        
        Dim fin_chaine As Integer
        fin_chaine = chains(i)(chains(i).Count).get_fin
        
        Dim quantite_effectuee As Integer
        quantite_effectuee = 0
        '17 5
        'définir les x
        'l.Cells(17 + (d - GANTT_horizontal_margin) / 4, 5).value = (pos_actuelle - debut_chaine) / (fin_chaine - debut_chaine)
        
        'on va définir les x et y du point
        'pour ça on récupère la durée du buffer
        Dim duree_buffer As Integer
        duree_buffer = ThisWorkbook.Worksheets("LOGS").Cells(14 + i, 16).value
        'on parcoure toutes les tâches de la chaîne
        Dim sh As Worksheet
        Set sh = ThisWorkbook.Worksheets("LOGS_AV")
        Dim updated As Boolean: updated = False 'indique qu'on a enregistré le point
        Dim j As Integer
        For j = 1 To chains(i).Count
            
            'on cherche son indice sur le tableau ordonnancé des tâches
            Dim k As Integer: k = 2
            While sh.Cells(k, 1).value <> CStr(chains(i)(j).get_ID)
                k = k + 1
            Wend
            
            If sh.Cells(k, col).value = 1 And updated = False Then
                quantite_effectuee = quantite_effectuee + chains(i)(j).get_duree
            End If
            
            Dim w As Integer: w = 1
            'quelle ligne doit-on remplir?
            While l.Cells(16 + w, 4 * i + 1).value <> ""
                w = w + 1
            Wend
            
            If (sh.Cells(k, col).value < 1 And updated = False) Or (sh.Cells(k, col).value = 1 And j = chains(i).Count And l.Cells(1, 4 * i + 1).value <> 1) Then 'si son avancement est <100% (marche parce qu'on parcoure le tableau dans le bon sens)
                'dernière tâche en cours
                updated = True
                
                If pos_actuelle = l.Cells(15 + w, 4 * i + 4).value Then
                    w = w - 1
                End If
                
                'ajout de la date à droite
                l.Cells(16 + w, 4 * i + 4).value = pos_actuelle
                
                'définir les x :
                If j < chains(i).Count Then
                    quantite_effectuee = quantite_effectuee + chains(i)(j).get_duree * sh.Cells(k, col).value
                    l.Cells(16 + w, 4 * i + 1).value = (quantite_effectuee / (fin_chaine - debut_chaine)) * 100
                Else
                    l.Cells(16 + w, 4 * i + 1).value = (quantite_effectuee / (fin_chaine - debut_chaine)) * 100
                    l.Cells(1, 4 * i + 1).value = 1 'la chaîne a été completée
                End If
                
                'il faut récupèrer la quantité consommée de buffer
                Dim buffer_consom As Integer: buffer_consom = 0
                
                'calcul de l'écart entre avancement théorique et avancement réel
                Dim avancement_theo As Double
                'MsgBox chains(i)(j).get_duree
                'MsgBox chains(i)(j).get_Intitule
                avancement_theo = (pos_actuelle - chains(i)(j).get_debut) / chains(i)(j).get_duree '"on devrait en être là"
                'MsgBox pos_actuelle
                Dim difference As Double
                difference = avancement_theo - sh.Cells(k, col).value 'avancement théorique - avancement réel, positif si retard
                'MsgBox difference
                'si on était en avance et que du coup on a rien foutu mais on reste en avance, ne pas retirer de consommation
                If Not (difference < 0 And l.Cells(16 + w, 4 * i + 1).value = l.Cells(15 + w, 4 * i + 1).value) Then
                    'cet écart d'avancement correspond à cb de jours?
                    'mise à jour de la consommation
                    buffer_consom = difference * chains(i)(j).get_duree 'on a pas fait tel %, ça correspond a tant de jours consommés
                End If
                'si on a consommé aucun buffer ou qu'on prend bcp d'avance
                If buffer_consom < 0 Then
                    buffer_consom = 0
                End If
                
                'définir le y
                l.Cells(16 + w, 4 * i + 2).value = (buffer_consom / duree_buffer * 2) * 100
                Debug.Print "Duree buffer " & duree_buffer & " Buffer conso " & buffer_consom
                
                'mise à l'échelle pour affichage graph
                l.Cells(16 + w, 4 * i + 3).value = l.Cells(16 + w, 4 * i + 1).value / 10 + 1
                
            End If
            
        Next j
        
    Next i
    
    'maj des cellules calendrier
    Dim left_limit As Integer
    left_limit = (pos_actuelle + 2) / 2 + GANTT_horizontal_margin - 1
                
    'redéfinir la zone sélectionnable
    'Set s = ThisWorkbook.Worksheets("GANTT")
    ThisWorkbook.Worksheets("GANTT").Select
    If left_limit = 0 Then
        MsgBox " Une erreur est survenue."
        Exit Sub
    End If
    ActiveSheet.Range(Cells(GANTT_vertical_margin - 2, GANTT_horizontal_margin), Cells(GANTT_vertical_margin - 2, left_limit - 4)).Interior.Color = RGB(200, 200, 200)
    
End Sub
