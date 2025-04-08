Attribute VB_Name = "ordonnancement"
' scheduling
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
    Set t_triees = New Collection 'Sorted task tables

    '--------- Critical chain generation ---------'

    'What task should I start with? Necessarily a "first task" - > any predecessor
    Dim j As Integer, indice_max As Integer, max As Integer, previous_max As Integer, ante As Object
    previous_max = 0
    indice_max = 1
    For j = 1 To taches.Count
        If taches(j).get_preds = "" Then
            max = taches(j).get_duree
            Dim z As Tache
            Set z = taches(j)
            While antecedants(z, taches).Count >= 1 'As long as there are still some teeth

                Set ante = antecedants(z, taches) 'Recover the antecdants of the last task sorted
                max = max + ante(next_tache(ante)).get_duree()
                If max > 10000 Then
                    MsgBox " Please check the previous entries."
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

    While antecedants(t_triees(t_triees.Count), taches).Count >= 1 'As long as there are still tasks to be added in a critical chain
        'Dim ante As Object
        Set ante = antecedants(t_triees(t_triees.Count), taches) 'Recover the previous dants of the last tri task
        t_triees.Add ante(next_tache(ante)) 'the greatest hardship is recovered among these predecessors
        Call remove_task_by_id(t_triees(t_triees.Count).get_ID, taches) 'the unsorted tasks are removed from the unsorted tasks

        t_triees(t_triees.Count).set_debut (t_triees(t_triees.Count - 1).get_fin)
        t_triees(t_triees.Count).set_fin (CInt(t_triees(t_triees.Count).get_debut) + CInt(t_triees(t_triees.Count).get_duree))
        t_triees(t_triees.Count).set_type (1) 'type critical chain

    Wend

    If d = 1 Then
        t_triees(t_triees.Count).set_type 4
    End If

    Dim chaine_critique As Collection
    Set chaine_critique = New Collection
    For j = 1 To t_triees.Count
        chaine_critique.Add t_triees(j) 'Save critical chain
    Next j
    'Set chaine_critique = t_triees 'the chain is recorded in a table

    '--------- End of the critical chain ---------'

    '--------- Positioning of tasks linked to the critical chain ---------'

        'The remaining tasks in "tasks" are not yet sorted
        'they are positioned around the critical chain while avoiding overlapping resources

    Dim i As Integer, k As Integer, end_loop As Integer, secondary_chains As Collection 'secondary_chains is meant to collect all 2ndary chains
    end_loop = t_triees.Count
    Set secondary_chains = New Collection

    Dim previous_size As Integer
    previous_size = chaine_critique.Count

    'MsgBox previous_size

    For i = 2 To end_loop
        k = end_loop + 2 - i 'we travel in the opposite direction to avoid being affected by the insertion

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

    'ATTENTION : To improve, the tasks that were added to and out of critical chains still exist in the s_chain
    'any stain that also exists in the critical chain must be eliminated from the s_chains

    '--------- Fine positioning of the tasks linked to the critical chain ---------'

    '--------- Positioning of the last unsorted tasks ---------'

        'At this stage, there may still be unsorted tasks (present in the table "tasks")
        'they potentially have one or more predecessors : we know their left limit
        'If they have no history, they can be placed in sync with the last task (Non-overlapping resource condition)
    'If d = 1 Then 'We do it only if buffers d�j� g�n�r�s
        Dim free_chains As Collection 'storage of the "free" chains that go �tre g�n�r�es
        Dim counter As Integer: counter = 0
        Set free_chains = New Collection

        'as long as there are still unsorted tasks
        While taches.Count > 0
            Dim Target As Integer
            Target = 1
            For i = 2 To taches.Count
                'trying to focus first on the tasks that have no antecedants
                If antecedants(taches(i), taches).Count = 0 Then
                   Target = i
                End If
            Next i
                'We have selected a task without prior control, we will position it

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



                'We have placed a starting task for potential recursivities

            'run recursion if preceded
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

    '--------- End positioning of the last unsorted tasks ---------'
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
            If t(i).get_preds = "" Then ' This is an initial task
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

                If task_in_tab_by_id(CInt(preds_id(j)), taches) = True Then 'The task is not yet sorted

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
                        'If critical, everything must be passed in a critical chain
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
                    Call remove_task_by_id(t(k).get_ID, taches) 'Remove from task table

                End If
            End If
        Next j

        Dim previous_adds As Integer
        previous_adds = 0

        If critical = False Then

            If UBound(preds_id) >= 0 Then 'we still have predecessors we have to continue the branch
                For i = 0 To counter - 1
                    'calls for a method for the next step
                    previous_adds = previous_adds + recursive_positioning(k + i + previous_adds, t, s_chain, critical_chain, alerte_on)
                Next i
            End If

            recursive_positioning = counter + previous_adds

        End If

    End If

End Function


'position the task in the best possible way by taking into account the potential overlap of resources (indice_initial = antecedent index)
Function set_intermediate_task(t As Collection, indice As Integer, cible As Tache, indice_initial As Integer, left_limit As Integer, Optional first_i As Integer = 0) ', counter As Integer)

    Dim i As Integer, rsrcs As Collection, debut As Integer, fin As Integer, match As Boolean, critical As Boolean
    Set rsrcs = New Collection
    match = False
    critical = False

    'Creating a resource table for the target
    For i = 1 To Len(cible.get_ress)
        If Not i Mod 2 = 0 Then
            rsrcs.Add Mid(cible.get_ress, i, 1)
        End If
    Next i

    'Calculation of the theoretical start date for the target
    debut = t(indice).get_debut - cible.get_duree + 1 'margin of 1 to avoid boundary detections
    fin = t(indice).get_debut - 1

    Dim k As Integer
    For k = 1 To t.Count
        Dim j As Integer, w As Integer, delay As Integer, duration As Integer
        i = t.Count - k + 1
        j = 1
        'Overlap condition

        If (t(i).get_fin >= debut And t(i).get_fin <= fin) Or (t(i).get_debut <= fin And t(i).get_debut >= debut) Or (t(i).get_debut <= debut And t(i).get_fin >= fin) Then
            While j <= rsrcs.Count And match = False
                'Same resource condition
                If InStr(1, t(i).get_ress, rsrcs(j)) <> 0 Then

                    match = True

                    If first_i = 0 Then 'first time looping = immediately left task
                        first_i = i
                    End If

                    If t(i).get_debut - cible.get_duree > left_limit Then 'if we can position our left spot from the one in conflict
                        critical = set_intermediate_task(t, i, cible, indice_initial, left_limit, first_i) 'In the end, will the task have to be placed on the right?
                    Else
                        critical = True
                        duration = cible.get_duree

                        'The difference between I and first_i gives us the number of itrations
                        Dim nb_iterations As Integer
                        nb_iterations = first_i - i
                        'so the first task to be postponed is first_i - nb_iterations + 1
                        'the target task must therefore slip between it and first_i - nb_iterations

                        'To put the task in a critical chain, the shift must be:
                        'the duration of the task minus the difference between the end of the firt_i-nb_iterations and the start of the task calculation

                        'delay = cible.get_duree - (t(first_i - nb_iterations).get_fin - debut)
                        delay = cible.get_duree
                        'delay = t(first_i).get_fin - t(indice_initial).get_debut + cible.get_duree

                        '"Adjust all the tasks right"
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

    'if there is no overlapping resource, we position just at the latest
    If match = False Then
        cible.set_fin (t(indice).get_debut)
        cible.set_debut (t(indice).get_debut - cible.get_duree)

        If cible.get_debut < left_limit Then
            'ATTENTION : to be improved, if the recalculation end date e d exceeds the current end date: We have to put the task in a critical chain (simplement retourner TRUE)
            'the critical chain tasks that have a beginning > This task must be removed from the critical chain if it is (will surely have to have it as a param)
            cible.set_fin (cible.get_fin + left_limit - cible.get_debut)
            cible.set_debut (cible.get_debut + left_limit - cible.get_debut)
        End If


    End If

   set_intermediate_task = critical

End Function


'Used to calculate a "left limit" when positioning an intermediate task
'We look among his predecessors already placed for the one with the longest end date
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
