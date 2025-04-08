Attribute VB_Name = "ordonnancement"
Sub ordo(Optional d As Integer = 0)
    
    planning_alert = False
    
    If d <> 1 Then
        Dim p As Integer, sh As Worksheet
        Set sh = ThisWorkbook.Worksheets("LOGS")
        While sh.Cells(22 + p, 9).value <> ""
            p = p + 1
        Wend
        sh.Range("O15:P200").Clear
        sh.Range(sh.Cells(22, 9), sh.Cells(24 + p, 11)).Clear
        Dim re_plan As Boolean: re_plan = False
        Call retrieve_tasks
    End If
    Dim t_triees As Collection, s As Worksheet, ta As Integer
    Set s = ThisWorkbook.Worksheets("GANTT")
    Set t_triees = New Collection 'tableaux de t�ches tri�es
    
    '--------- G�n�ration de la cha�ne critique ---------'
    
        'par quelle t�che commencer? forc�ment une "premi�re t�che" -> aucun pr�d�cesseur
    Dim j As Integer, indice_max As Integer, max As Integer, previous_max As Integer, ante As Object
    previous_max = 0
    indice_max = 1
    For j = 1 To taches.Count
        If taches(j).get_preds = "" Then
            max = taches(j).get_duree
            Dim z As Tache
            Set z = taches(j)
            While antecedants(z, taches).Count >= 1 'Tant qu'il reste des ant�c�dents
                
                Set ante = antecedants(z, taches) 'r�cup�rer les ant�c�dants de la derni�re tache tri�es
                max = max + ante(next_tache(ante)).get_duree()
                If max > 10000 Then
                    MsgBox " Veuillez v�rifier les pr�cesseurs saisis."
                    Exit Sub
                End If
                Set z = ante(next_tache(ante))
                
            Wend
            If previous_max > 0 Then
                If previous_max < max Then
                    indice_max = j
                    previous_max = max
                End If
            Else
                indice_max = j
                previous_max = max
            End If
            'MsgBox max
            'MsgBox indice_max
        End If
    Next j
    'MsgBox indice_max
    't_triees.Add taches(next_tache(taches, True)) 'si indication "true" : on cherche une t�che sans pr�decesseur
    'taches.Remove next_tache(taches, True)
    
    t_triees.Add taches(indice_max)
    taches.Remove indice_max
    
    t_triees(1).set_debut (0)
    t_triees(1).set_fin (t_triees(1).get_duree)
    t_triees(1).set_type (1)
        
    While antecedants(t_triees(t_triees.Count), taches).Count >= 1 'Tant qu'il reste des t�ches � ajouter en cha�ne critique
    
        'Dim ante As Object
        Set ante = antecedants(t_triees(t_triees.Count), taches) 'r�cup�rer les ant�c�dants de la derni�re tache tri�e
        t_triees.Add ante(next_tache(ante)) 'on r�cup�re la plus grande dur�e t�ches parmi ces ant�c�dants
        Call remove_task_by_id(t_triees(t_triees.Count).get_ID, taches) 'on supprime des t�ches non tri�es celle qui vient de l'�tre
    
        t_triees(t_triees.Count).set_debut (t_triees(t_triees.Count - 1).get_fin)
        t_triees(t_triees.Count).set_fin (CInt(t_triees(t_triees.Count).get_debut) + CInt(t_triees(t_triees.Count).get_duree))
        t_triees(t_triees.Count).set_type (1) 'type cha�ne critique
        
    Wend
    
    If d = 1 Then
        t_triees(t_triees.Count).set_type 4
    End If
    
    Dim chaine_critique As Collection
    Set chaine_critique = New Collection
    For j = 1 To t_triees.Count
        chaine_critique.Add t_triees(j) 'enregistrer chaine critique
    Next j
    'Set chaine_critique = t_triees 'on enregistre la cha�ne dans un tableau
    
    '--------- Fin g�n�ration de la cha�ne critique ---------'

    '--------- Positionnement des t�ches li�es � la cha�ne critique ---------'
    
        'les t�ches restantes dans "taches" ne sont pas encore tri�es
        'on les positionne autour de la cha�ne critique en �vitant le chevauchement des ressources
    
    Dim i As Integer, k As Integer, end_loop As Integer, secondary_chains As Collection 'secondary_chains is meant to collect all 2ndary chains
    end_loop = t_triees.Count
    Set secondary_chains = New Collection
    
    Dim previous_size As Integer
    previous_size = chaine_critique.Count
    
    'MsgBox previous_size
    
    For i = 2 To end_loop
        k = end_loop + 2 - i 'on parcoure dans le sens inverse pour pas �tre affect� par l'insertion
        
        'MsgBox t_triees(k).get_Intitule
        
        Dim s_chain As Collection
        Set s_chain = New Collection 'one 2dary chain
        
        If d = 1 Then
            Call recursive_positioning(k, t_triees, s_chain, chaine_critique, 1) 'lancement r�cursivit� avec surveillance des alertes
        Else
            Call recursive_positioning(k, t_triees, s_chain, chaine_critique)
        End If
        
        'saving 2ndary chain only if it has tasks
        If s_chain.Count > 0 Then
            'MsgBox "aaaaahhh"
            secondary_chains.Add s_chain
        End If
    Next i
    
    'MsgBox chaine_critique.Count
    
    While previous_size <> chaine_critique.Count
        
        If re_plan = False And d = 1 Then
            re_plan = True
        End If
        
        For i = 2 To chaine_critique.Count
            
            k = chaine_critique.Count + 2 - i
            
            Set s_chain = New Collection 'one 2dary chain
            
            Call recursive_positioning(i, t_triees, s_chain, chaine_critique)
            
            'saving 2ndary chain only if it has tasks
            If s_chain.Count > 0 Then
                'MsgBox "aaaaahhh"
                secondary_chains.Add s_chain
            End If
            
        Next i
        previous_size = chaine_critique.Count
    Wend
    
    'ATTENTION : � am�liorer, les t�ches qui ont p� �tre ajout�e en cha�ne critique puis qui en sont sortie existent encore dans les s_chain
    'faut �liminer des s_chains toute t�che qui existe aussi dans la cha�ne critique
    
    '--------- Fin positionnement des t�ches li�es � la cha�ne critique ---------'
    
    '--------- Positionnement des derni�res t�ches non-tri�es ---------'
    
        '� ce stade, il peut encore y avoir des t�ches non tri�es (pr�sentent dans le tableau "taches")
        'elles ont potentiellement un/des pr�decesseurs : on connait leur limite � gauche
        'si elles n'ont pas d'antecedants, on peut les placer synchro avec la derni�re t�che (� condition de non chevauchement ressources)
    'If d = 1 Then 'on le fait que si buffers d�j� g�n�r�s
        Dim free_chains As Collection 'stockage des cha�nes "libres" qui vont �tre g�n�r�es
        Dim counter As Integer: counter = 0
        Set free_chains = New Collection
        
        'tant qu'il reste des t�ches non tri�es
        While taches.Count > 0
            Dim Target As Integer
            Target = 1
            For i = 2 To taches.Count
                'trying to focus first on the tasks that have no antecedants
                If antecedants(taches(i), taches).Count = 0 Then
                   Target = i
                End If
            Next i
                'nous avons s�lectionn� une t�che sans ant�c�dants, nous allons la positionner
            
            'Dim fake_task As Tache
            'Set fake_task = New Tache
            'fake_task.set_attributes "", "8", "Z", "" 'remplissage infos
            'fake_task.set_debut (t_triees(last_task_indice(t_triees)).get_fin)
            'fake_task.set_fin (fake_task.get_debut + CInt(fake_task.get_duree))
            
            
            't_triees.Add fake_task
            
            Call set_intermediate_task(t_triees, t_triees.Count - counter, taches(Target), last_task_indice(t_triees), max_preds_end(taches(Target), t_triees))
            counter = counter + 1
            taches(Target).set_type (3) 'free type
            t_triees.Add taches(Target)
            taches.Remove Target
            
            't_triees.Remove t_triees.Count
            
            
            
                'nous avons plac� une t�che de d�but pour de potentielles r�cursivit�s
            
            'lancer la recursivit� si predecesseurs
            If t_triees(t_triees.Count).get_preds <> "" Then
                Dim f_chain As Collection
                Set f_chain = New Collection
                Call recursive_positioning(t_triees.Count, t_triees, f_chain, chaine_critique)
            
                If f_chain.Count > 0 Then
                    free_chains.Add f_chain
                End If
            End If
            
        Wend
    'End If
    
    '--------- Fin positionnement des derni�res t�ches non-tri�es ---------'
    Dim oh As Integer: oh = 0
    Dim rev As Integer
    For i = 1 To t_triees.Count
        rev = t_triees.Count + 1 - i
        ThisWorkbook.Worksheets("LOGS").Cells(i + 21, 9) = t_triees(i).get_ID
        ThisWorkbook.Worksheets("LOGS").Cells(i + 21, 10) = t_triees(i).get_debut
        ThisWorkbook.Worksheets("LOGS").Cells(i + 21, 11) = t_triees(i).get_fin
        If t_triees(rev).get_type = 4 Then
            ThisWorkbook.Worksheets("LOGS").Cells(15 + oh, 17) = t_triees(rev).get_debut
            oh = oh + 1
        End If
    Next i
    
    If d = 1 Then
        If re_plan = False Then
            Call affichage_GANTT(t_triees)
        Else
            Call affichage_GANTT(t_triees)
            'Call retrieve_tasks
            'Call remove_chains_first_task(secondary_chains)
            'Call generate_buffers(chaine_critique, secondary_chains)
            'Call ordo(1)
        End If
    End If
    If d <> 1 Then
        Dim c As String
        c = ""
        For i = 1 To chaine_critique.Count
            c = c + CStr(chaine_critique(i).get_ID)
            If i < chaine_critique.Count Then
                c = c + ","
            End If
        Next i
        ThisWorkbook.Worksheets("LOGS").Cells(15, 15).value = c
        
        Call retrieve_tasks
        Call generate_buffers(chaine_critique, secondary_chains)
        Call ordo(1)
        
        'ThisWorkbook.Worksheets("LOGS").Cells(15, 16).value = 20
        '    Call affichage_GANTT(t_triees)
    End If

    
    
    
    'If d = 1 And planning_alert = False Then
    '    Call affichage_GANTT(t_triees)
    'End If
    'If d <> 1 And planning_alert = False Then
    '    Dim c As String
    '    c = ""
    '    For i = 1 To chaine_critique.Count
    '        c = c + CStr(chaine_critique(i).get_ID)
    '        If i < chaine_critique.Count Then
    '            c = c + ","
    '        End If
    '    Next i
    '    ThisWorkbook.Worksheets("LOGS").Cells(15, 15).value = c
    '
    '    If d = 2 Then
    '        MsgBox "bz"
    '        Call retrieve_tasks
    '        Call generate_buffers(chaine_critique, secondary_chains, 1)
    '        Call ordo(1)
    '    Else
    '        Call retrieve_tasks
    '        Call generate_buffers(chaine_critique, secondary_chains)
    '        Call ordo(1)
    '    End If
    'End If
    '
    'If planning_alert = True Then
    '    Call ordo(2)
    'End If

End Sub


Public Function next_tache(t As Collection, Optional premiere_tache As Boolean = False) As Integer

    Dim i As Integer, max As Integer, indice As Integer
    max = 0
    indice = 0
    If premiere_tache = True Then
        For i = 1 To t.Count
            If t(i).get_preds = "" Then ' C'est une t�che initiale
                If t(i).get_duree() > max Then
                    indice = i
                    max = t(i).get_duree()
                End If
            End If
        Next i
    Else
        For i = 1 To t.Count
            If t(i).get_duree() > max Then
                indice = i
                max = t(i).get_duree()
            End If
        Next i
    End If
    next_tache = indice
End Function


Function recursive_positioning(k As Integer, t As Collection, s_chain As Collection, critical_chain As Collection, Optional alerte_on As Integer = 0) As Integer ', Optional ByRef counter As Integer = 0)
    
    
    If CStr(t(k).get_preds) <> "" Then
        Dim preds_id() As String, j As Integer, i As Integer, counter As Integer, left_limit As Integer, critical As Boolean
        preds_id = Split(t(k).get_preds, ",")
        
        counter = 0 'cb de t�ches j'ai plac�
        
        For j = 0 To UBound(preds_id)

            If preds_id(j) <> "" Then

                If task_in_tab_by_id(CInt(preds_id(j)), taches) = True Then 'la tache n'est pas encore tri�e
                    
                    left_limit = max_preds_end(taches(get_task_index_by_id(CInt(preds_id(j)), taches)), t)
                    'If left_limit > 0 Then
                    '    MsgBox CStr(left_limit) + "pour " + taches(get_task_index_by_id(CInt(preds_id(j)), taches)).get_Intitule
                    'End If
                    critical = set_intermediate_task(t, k + counter, taches(get_task_index_by_id(CInt(preds_id(j)), taches)), k + counter, left_limit) ', counter)
                    
                    'ajout de la t�che dans t_triees
                    Call insertion_by_indice(taches(get_task_index_by_id(CInt(preds_id(j)), taches)), t, k) 'target task becomes t(k)
                    counter = counter + 1
                    
                    If critical = False Then
                    
                        s_chain.Add t(k) 'task is registered in a secondary chain
                        If t(k).get_type <> 4 Then
                            t(k).set_type (2) 'intermediate type
                        End If
                    Else
                        
                        If alerte_on = 1 Then
                            'd�clencher alerte
                            planning_alert = True
                        End If
                        'si critique, tout doit �tre pass� en cha�ne critique
                        critical_chain.Add t(k)
                        t(k).set_type (1)
                        
                        'For i = 0 To counter
                        '    'si la tache n'est pas d�j� en cha�ne critique
                        '    If task_in_tab_by_id(t(k + i).get_ID, critical_chain) = False Then
                        '        critical_chain.Add t(k + i) 'on l'ajoute
                        '        If t(k + i).get_type <> 4 Then
                        '            t(k + i).set_type (1) 'type critique
                        '        End If
                        '    End If
                        'Next
                    End If
                    Call remove_task_by_id(t(k).get_ID, taches) 'retirer du tableau de t�ches
    
                End If
            End If
        Next j
        
        Dim previous_adds As Integer
        previous_adds = 0
        
        If critical = False Then
        
            If UBound(preds_id) >= 0 Then 'on a encore des pr�d�cesseurs faut poursuivre la branche
                For i = 0 To counter - 1
                    'appelle de la m�thode pour le prochain pallier
                    previous_adds = previous_adds + recursive_positioning(k + i + previous_adds, t, s_chain, critical_chain, alerte_on)
                Next i
            End If
        
            recursive_positioning = counter + previous_adds
        
        End If
        
    End If

End Function


'positionner la t�che au mieux en prenant en compte le chevauchement potentiel de ressources (indice_initial = indice de l'antecedant)
Function set_intermediate_task(t As Collection, indice As Integer, cible As Tache, indice_initial As Integer, left_limit As Integer, Optional first_i As Integer = 0) ', counter As Integer)
    
    Dim i As Integer, rsrcs As Collection, debut As Integer, fin As Integer, match As Boolean, critical As Boolean
    Set rsrcs = New Collection
    match = False
    critical = False
    
    'cr�ation d'un tableau de ressources pour la cible
    For i = 1 To Len(cible.get_ress)
        If Not i Mod 2 = 0 Then
            rsrcs.Add Mid(cible.get_ress, i, 1)
        End If
    Next i
    
    'calcul de la date de d�but th�orique pour la cible
    debut = t(indice).get_debut - cible.get_duree + 1 'marge de 1 pour �viter les d�tections aux limites
    fin = t(indice).get_debut - 1
    
    Dim k As Integer
    For k = 1 To t.Count
        Dim j As Integer, w As Integer, delay As Integer, duration As Integer
        i = t.Count - k + 1
        j = 1
        'condition de chevauchement
            
        If (t(i).get_fin >= debut And t(i).get_fin <= fin) Or (t(i).get_debut <= fin And t(i).get_debut >= debut) Or (t(i).get_debut <= debut And t(i).get_fin >= fin) Then
            While j <= rsrcs.Count And match = False
                'condition de ressource identique
                If InStr(1, t(i).get_ress, rsrcs(j)) <> 0 Then
                            
                    match = True
                            
                    If first_i = 0 Then 'premi�re fois qu'on boucle = tache imm�diatement � gauche
                        first_i = i
                    End If
                            
                    If t(i).get_debut - cible.get_duree > left_limit Then 'si on peut positionner notre t�che � gauche de celle en conflit
                        critical = set_intermediate_task(t, i, cible, indice_initial, left_limit, first_i) 'au bout du compte la t�che devra-t-elle �tre plac�e � droite?
                    Else
                        critical = True
                        duration = cible.get_duree
                        
                        'la diff�rence entre i et first_i nous donne le nombre d'it�rations
                        Dim nb_iterations As Integer
                        nb_iterations = first_i - i
                        'ainsi la premi�re t�ches � devoir �tre d�caler est first_i - nb_iterations + 1
                        'la t�che cible doit donc se glisser entre celle-ci et first_i - nb_iterations
                        
                        'pour ins�rer la t�che en cha�ne critique, il faut que le d�calagage soit :
                        'la dur�e de la t�che moins la diff entre la fin de firt_i-nb_iterations et le d�but calcul� de la t�che
                       
                        'delay = cible.get_duree - (t(first_i - nb_iterations).get_fin - debut)
                        delay = cible.get_duree
                        'delay = t(first_i).get_fin - t(indice_initial).get_debut + cible.get_duree
                                
                        '"d�caler toutes les t�ches � droite"
                        'MsgBox nb_iterations
                        
                        For w = first_i - nb_iterations + 1 To t.Count
                            'MsgBox t(w).get_Intitule
                            t(w).set_fin (t(w).get_fin + delay)
                            t(w).set_debut (t(w).get_debut + delay)
                        Next w
                        'cible.set_debut (t(first_i).get_fin)
                        'cible.set_fin (cible.get_debut + duration)
                        cible.set_debut (t(first_i - nb_iterations).get_fin)
                        cible.set_fin (cible.get_debut + duration)
                    End If
                            
                End If
                j = j + 1
            Wend
        End If
    Next k

    'si aucun chevauchement de ressource, on positionne juste au plus tard
    If match = False Then
        cible.set_fin (t(indice).get_debut)
        cible.set_debut (t(indice).get_debut - cible.get_duree)

        If cible.get_debut < left_limit Then
            'ATTENTION : a am�liorer, si la date de fin recalcul�e d�passe la date de fin actuelle : il faut mettre la t�che en chaine critique (simplement retourner TRUE)
            'il faut que les t�ches de cha�ne critique qui ont un d�but > � cette t�che doivent �tre sortie de cha�ne critique si elles y sont (faudra surement l'avoir en param)
            cible.set_fin (cible.get_fin + left_limit - cible.get_debut)
            cible.set_debut (cible.get_debut + left_limit - cible.get_debut)
        End If
        
        
    End If
    
   set_intermediate_task = critical
    
End Function


'utilis�e pour calculer une "limite gauche" lors du positionnement d'une t�che interm�diaire
'on cherche parmi ses pr�decesseurs d�j� plac�s celui qui a la plus grande date de fin
Function max_preds_end(task As Tache, t As Collection) As Integer


    Dim res As Integer, i As Integer
    res = 0
    
    Dim preds() As String
    preds = Split(task.get_preds, ",")
    
    For i = 0 To UBound(preds)
        If preds(i) <> "" Then
            If task_in_tab_by_id(CInt(preds(i)), t) = True Then
                If t(get_task_index_by_id(CInt(preds(i)), t)).get_fin > res Then
                    res = t(get_task_index_by_id(CInt(preds(i)), t)).get_fin
                End If
            End If
        End If
    Next i

    max_preds_end = res

End Function

