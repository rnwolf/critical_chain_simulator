Attribute VB_Name = "avancement"
' advancement
Sub ordonnancement_CCPM()
    ' scheduling

    Call retrieve_tasks
    Call record_avancement
    Call couleur_avancement
    Dim t As Collection, s As Worksheet, k As Worksheet, sh As Worksheet, cc As Worksheet, i As Integer, j As Integer, l As Integer, m As Integer, variable As Integer
    Dim case_debut As Integer, case_fin As Integer, avancement As Double, date_actuelle As Integer, vertical_margin As Integer, marge As Integer
    Dim check As Boolean, u As Integer, max As Integer
    Dim temps_theorique As Integer
    Dim case_debut_theorique As Integer 'num�ro de colonne
    Dim conso_buffer As Integer
    Dim pourcentage_conso_buffer As Integer
    Dim duree_buffer As Integer
    Dim indice_ligne As Integer
    Dim splito() As String
    Dim min_debut As Integer
    Dim decalage As Integer
    Dim sauv As Integer
    Dim chaine As Integer
    Dim debut_buffer As Integer, fin_buffer As Integer
    Dim fini As Integer
    Dim marge_fin As Integer
    Set s = ThisWorkbook.Worksheets("GANTT")
    Set k = ThisWorkbook.Worksheets("LOGS")
    Set sh = ThisWorkbook.Worksheets("LOGS_FV_CHART")
    Set cc = ThisWorkbook.Worksheets("LOGS_CCPM")
    Set t = taches 'start index 1
    marge = 6 'columns
    vertical_margin = 6 ' lines
    marge_fin = ThisWorkbook.Worksheets("LOGS").Cells(2, 1).value + 3 'marge + 4 avant

    j = 0
    date_actuelle = colonne_date_actuelle
    max = 0
    Range(cc.Cells(2, 3), cc.Cells(250, 4)).ClearContents 'Clean offset values
    If date_actuelle = 0 Then
        Exit Sub
    End If

    For i = 1 To t.Count ' set les d�buts

        l = 1
        While k.Cells(21 + l, 9) <> t(i).get_ID
            l = l + 1
        Wend
        t(i).set_debut (k.Cells(21 + l, 10) / 2 + marge)

    Next i ' set les d�buts

    Dim nb_chaines As Integer
    nb_chaines = 0

    While k.Cells(j + 15, 15) <> "" ' Calculate the number of strings
        nb_chaines = nb_chaines + 1
        j = j + 1
    Wend

    j = 0
    While k.Cells(j + 15, 15) <> ""

        '--------------------------------- Critical Chain -----------------------------------
        If j = 0 Then ' We are in the critical chain
            Debug.Print "Cha�ne critique"
            l = 0
            While sh.Cells(17 + l, 6) <> "" 'Take the last % of consumption buffer
                l = l + 1
            Wend

            pourcentage_conso_buffer = sh.Cells(17 + l - 1, 6) ' It is a percentage
            Debug.Print "Buffer Percentage Consumed " & pourcentage_conso_buffer

            If pourcentage_conso_buffer > 0 Then ' we consumed buffer

                duree_buffer = CInt(k.Cells(15, 16) / 4)
                conso_buffer = pourcentage_conso_buffer / 100 * duree_buffer
                Debug.Print "Buffer Consumes " & conso_buffer
                cc.Cells(t.Count + 2, 4) = conso_buffer

                For m = 1 To t.Count ' Calculation of shifts
                    indice_ligne = trouver_ligne_indice(t(m).get_ID)
                    avancement = s.Cells(indice_ligne, 3)

                    decalage = conso_buffer
                    sauv = cc.Cells(1 + t(m).get_ID, 3)
                    cc.Cells(1 + t(m).get_ID, 3) = sauv + decalage

                Next m

                '------------- Timing of other buffers  -----------

                For m = 1 To nb_chaines
                    If m <> 1 Then
                    indice_ligne = trouver_ligne_indice(t.Count + m)
                    'If IsNumeric(avancement) = True Then
                    'avancement = s.Cells(indice_ligne, 3)
                    'decalage = avancement * k.Cells(15 + m - 1, 16) / 4 + conso_buffer
                    decalage = conso_buffer
                    sauv = cc.Cells(1 + t.Count + m, 3)
                    cc.Cells(1 + t.Count + m, 3) = sauv + decalage
                    End If
                Next m

            End If ' we consumed buffer

            '----------------------------- Secondary chain -----------------------------------

        Else ' in a secondary chain

            Debug.Print "Non-critical channel number " & j
            l = 0

            While sh.Cells(17 + l, 4 * (j + 1) + 2) <> "" 'take the last % of buffer consumption, before test
                l = l + 1
            Wend

            pourcentage_conso_buffer = sh.Cells(17 + l - 1, 4 * (j + 1) + 2) ' It is a percentage
            Debug.Print "Buffer Percentage Consumed " & pourcentage_conso_buffer
            If pourcentage_conso_buffer > 0 Then ' we consumed buffer

                'Buffer settings
                duree_buffer = CInt(k.Cells(15 + j, 16) / 4)
                conso_buffer = pourcentage_conso_buffer / 100 * duree_buffer
                cc.Cells(t.Count + 2 + j, 4) = conso_buffer
                debut_buffer = k.Cells(15 + j, 17) 'We get the start date back
                fin_buffer = debut_buffer + CInt(k.Cells(15 + j, 16) / 4)
                Debug.Print "D�but buffer " & debut_buffer & " et fin buffer " & fin_buffer

                splito = Split(k.Cells(15 + j, 15), ",")

                        '------------------- Overconsumption buffer -------------------


                If pourcentage_conso_buffer > 100 Then ' We overconsumed the buffer, we postpone all the tasks that date from after the end of the buffer

                    For m = 1 To t.Count ' We postpone all the tasks after the end of the buffer of the chain

                        indice_ligne = trouver_ligne_indice(t(m).get_ID) 'We find our line
                        avancement = s.Cells(indice_ligne, 3)

                        If fin_buffer < t(m).get_debut Then ' task that starts after the buffer ends
                            decalage = conso_buffer
                            sauv = cc.Cells(1 + t(m).get_ID, 3)
                            cc.Cells(1 + t(m).get_ID, 3) = sauv + decalage

                        End If ' task started after the end of buffer
                    Next m

                    '----- Offsetting the other buffers if needed. ---------
                    For m = 1 To nb_chaines
                        indice_ligne = trouver_ligne_indice(t.Count + m) 'We find our line
                        'avancement = s.Cells(indice_ligne, 3)

                        If m <> j + 1 Then ' not to calibrate the chain even of the buffer
                            If fin_buffer < k.Cells(15 + m - 1, 17) Then ' Buffer that starts after the buffer ends
                                decalage = conso_buffer
                                sauv = cc.Cells(1 + t.Count + m, 3)
                                cc.Cells(1 + t.Count + m, 3) = sauv + decalage
                                Debug.Print "On decale la chaine " & m - 1 & " de " & decalage & "indice " & 1 + t.Count + m
                            End If ' buffer starts after the end of buffer of our task
                        End If
                    Next m
                            ' ------------- Non-Integral Consumption -------------

                Else ' we have consumed but not entirely, we shift the tasks of the chain

                    Debug.Print "Lbound split = " & LBound(splito) & "Ubound split = " & UBound(splito)
                    For m = LBound(splito) To UBound(splito) 'for the chain tasks, not the buffer in any case so osef to sort

                        indice_ligne = trouver_ligne_indice(CInt(splito(m))) 'The line of the task can be found in the GANTT tab
                        avancement = s.Cells(indice_ligne, 3)

                        decalage = conso_buffer
                        sauv = cc.Cells(1 + CInt(splito(m)), 3)
                        cc.Cells(1 + CInt(splito(m)), 3) = sauv + decalage

                    Next m

                End If ' Buffer consumption. If less than 0 or equal, nothing happens.

            End If 'secondary channel, the indep will be outside the

        End If 'in the critical chain
        j = j + 1
    Wend

    '------------- Task tracing -------------

    For m = 1 To t.Count

        indice_ligne = trouver_ligne_indice(t(m).get_ID) 'We find our line
        avancement = s.Cells(indice_ligne, 3)
        fini = cc.Cells(1 + t(m).get_ID, 5)
        If fini = 0 Then
            chaine = dans_quel_chaine(t(m).get_ID)
            Debug.Print "Task " & t(m).get_ID & " in each " & chaine

            case_debut = t(m).get_debut + avancement * t(m).get_duree / 2 + cc.Cells(1 + t(m).get_ID, 3)
            case_fin = case_debut + (1 - avancement) * t(m).get_duree / 2 - 1

            If avancement = 1 Then ' Last course of this task
                cc.Cells(1 + t(m).get_ID, 5) = 1
            End If

            Debug.Print "task " & t(m).get_ID & "case_deb " & case_debut & " fin " & case_fin & " oui " & t(m).get_duree / 2 - 1 & "avancement " & 1 - avancement & " donc " & (1 - avancement) * t(m).get_duree / 2 - 1
            If case_fin < 0 Then
                MsgBox "Please check the value of the advances entered."
                Exit Sub
            End If
            'Clean
            Range(s.Cells(indice_ligne + 1, marge), s.Cells(indice_ligne + 1, marge_fin)).ClearContents
            Range(s.Cells(indice_ligne + 1, marge), s.Cells(indice_ligne + 1, marge_fin)).Interior.ColorIndex = 2
            Range(s.Cells(indice_ligne, CInt(t(m).get_debut)), s.Cells(indice_ligne, CInt(t(m).get_debut + t(m).get_duree / 2))).Interior.Pattern = xlPatternSolid

            'Tracing
            If avancement <> 1 Then
                s.Cells(indice_ligne + 1, case_debut) = t(m).get_ID ' Num�roter la t�che

                If chaine = 0 Then ' Color Management
                    Range(s.Cells(indice_ligne + 1, case_debut), s.Cells(indice_ligne + 1, case_fin)).Interior.ColorIndex = 22
                    If avancement <> 0 Then
                        Range(s.Cells(indice_ligne, CInt(t(m).get_debut)), s.Cells(indice_ligne, CInt(t(m).get_debut + t(m).get_duree / 2 - 1))).Interior.ColorIndex = 3
                        If avancement >= 1 Then
                            Range(s.Cells(indice_ligne, CInt(t(m).get_debut)), s.Cells(indice_ligne, CInt(t(m).get_debut + t(m).get_duree / 2 - 1))).Interior.Pattern = xlPatternLightUp
                        Else
                        Range(s.Cells(indice_ligne, CInt(t(m).get_debut)), s.Cells(indice_ligne, CInt(t(m).get_debut + avancement * t(m).get_duree / 2 - 1))).Interior.Pattern = xlPatternLightUp
                        End If
                        s.Cells(indice_ligne, case_debut).Font.ColorIndex = 2
                    End If
                ElseIf chaine = -1 Then 'The task is not in a chain
                    If avancement <> 0 Then
                        Range(s.Cells(indice_ligne, CInt(t(m).get_debut)), s.Cells(indice_ligne, CInt(t(m).get_debut + t(m).get_duree / 2 - 1))).Interior.ColorIndex = 5
                        If avancement >= 1 Then
                            Range(s.Cells(indice_ligne, CInt(t(m).get_debut)), s.Cells(indice_ligne, CInt(t(m).get_debut + t(m).get_duree / 2 - 1))).Interior.Pattern = xlPatternLightUp
                        Else
                        Range(s.Cells(indice_ligne, CInt(t(m).get_debut)), s.Cells(indice_ligne, CInt(t(m).get_debut + avancement * t(m).get_duree / 2 - 1))).Interior.Pattern = xlPatternLightUp
                        End If
                    End If
                    Range(s.Cells(indice_ligne + 1, case_debut), s.Cells(indice_ligne + 1, case_fin)).Interior.ColorIndex = 34

                Else ' the task is in a secondary chain
                    If avancement <> 0 Then
                        Range(s.Cells(indice_ligne, CInt(t(m).get_debut)), s.Cells(indice_ligne, CInt(t(m).get_debut + t(m).get_duree / 2 - 1))).Interior.ColorIndex = 4
                        If avancement >= 1 Then
                            Range(s.Cells(indice_ligne, CInt(t(m).get_debut)), s.Cells(indice_ligne, CInt(t(m).get_debut + t(m).get_duree / 2 - 1))).Interior.Pattern = xlPatternLightUp
                        Else
                        Range(s.Cells(indice_ligne, CInt(t(m).get_debut)), s.Cells(indice_ligne, CInt(t(m).get_debut + avancement * t(m).get_duree / 2 - 1))).Interior.Pattern = xlPatternLightUp
                        End If

                    End If
                    Range(s.Cells(indice_ligne + 1, case_debut), s.Cells(indice_ligne + 1, case_fin)).Interior.ColorIndex = 35
                End If ' which chain (color management)
            Else 'avancement �gale � 1
                Range(s.Cells(indice_ligne, CInt(t(m).get_debut)), s.Cells(indice_ligne, CInt(t(m).get_debut + t(m).get_duree / 2 - 1))).Interior.Pattern = xlPatternLightUp
                Range(s.Cells(indice_ligne + 1, marge), s.Cells(indice_ligne + 1, marge_fin)).ClearContents
                Range(s.Cells(indice_ligne + 1, marge), s.Cells(indice_ligne + 1, marge_fin)).Interior.ColorIndex = 2
                s.Cells(indice_ligne, 3).Interior.ColorIndex = 15

            End If 'Avancement diff�rent de 1
        Else
            Range(s.Cells(indice_ligne, CInt(t(m).get_debut)), s.Cells(indice_ligne, CInt(t(m).get_debut + t(m).get_duree / 2 - 1))).Interior.Pattern = xlPatternLightUp
            Range(s.Cells(indice_ligne + 1, marge), s.Cells(indice_ligne + 1, marge_fin)).ClearContents
            Range(s.Cells(indice_ligne + 1, marge), s.Cells(indice_ligne + 1, marge_fin)).Interior.ColorIndex = 2
            s.Cells(indice_ligne, 3).Interior.ColorIndex = 15
        End If ' Fini
    Next m

    '------------- Buffer tracking + their consumption -----------
    For m = 1 To nb_chaines
        indice_ligne = trouver_ligne_indice(t.Count + m)
        debut_buffer = k.Cells(15 + m - 1, 17) / 2 + 6 + cc.Cells(t.Count + m + 1, 3) 'On r�cup�re la date de d�but et on ajoute le d�calage (d� � autres cha�nes), /2 +6 pr conversion heures en colonne
        'debut_buffer = debut_buffer / 2 + 6
        conso_buffer = cc.Cells(t.Count + m + 1, 4)
        duree_buffer = CInt(k.Cells(15 + m - 1, 16) / 4)
        If debut_buffer = 0 Then 'Protection
            MsgBox "Problem encountered me, please update the classic GANTT."
            Exit Sub
        End If 'Protection

        'Debug.Print "D�but buffer cha�ne " & m - 1 & " � " & debut_buffer & " indice de ligne " & indice_ligne & " et conso " & conso_buffer & " et dur�e " & duree_buffer
        'Clean
        Range(s.Cells(indice_ligne + 1, marge), s.Cells(indice_ligne + 1, marge_fin)).ClearContents
        Range(s.Cells(indice_ligne + 1, marge), s.Cells(indice_ligne + 1, marge_fin)).Interior.ColorIndex = 2
        'Tra�age
        If conso_buffer <> 0 Then
            Range(s.Cells(indice_ligne + 1, debut_buffer), s.Cells(indice_ligne + 1, debut_buffer + conso_buffer - 1)).Interior.ColorIndex = 1 ' y'a un -1 � mettre mais jsp pq
        Else
            s.Cells(indice_ligne + 1, debut_buffer).Interior.ColorIndex = 15
            'Range(s.Cells(indice_ligne + 1, debut_buffer), s.Cells(indice_ligne + 1, debut_buffer + duree_buffer - 1)).Interior.ColorIndex = 15 'rajout� r�cemment
        End If
        If m = 1 Then ' num�roter
            s.Cells(indice_ligne + 1, debut_buffer).Font.ColorIndex = 2
            s.Cells(indice_ligne + 1, debut_buffer) = "Buffer cha�ne critique" ' Num�roter le buffer

        Else
            s.Cells(indice_ligne + 1, debut_buffer) = "Buffer " & m - 1 ' Num�roter le buffer
            s.Cells(indice_ligne + 1, debut_buffer).Font.ColorIndex = 2
        End If ' num�roter
        If conso_buffer < duree_buffer Then 'Si on a pas tout consomm�
            Range(s.Cells(indice_ligne + 1, debut_buffer + conso_buffer), s.Cells(indice_ligne + 1, debut_buffer + duree_buffer - 1)).Interior.ColorIndex = 15 'pareil pr le -1
        End If ' if we don't have tt consumption
    Next m

End Sub


'Write to logs and call the buffer consumption
Sub record_avancement()

    Dim i As Integer, s As Worksheet, pos_actuelle As Integer
    Set s = ThisWorkbook.Worksheets("LOGS_AV")
    pos_actuelle = (colonne_date_actuelle - GANTT_horizontal_margin) * 2 ' +2
    i = 2
    If colonne_date_actuelle > 30000 Then 'Protection erreur
        Exit Sub
    End If

    While s.Cells(1, i).value <> pos_actuelle And s.Cells(1, i).value <> ""
        i = i + 1
    Wend

    s.Cells(1, i) = pos_actuelle
    Dim j As Integer, g As Worksheet
    j = GANTT_vertical_margin
    Set g = ThisWorkbook.Worksheets("GANTT")
    While g.Cells(j, 3).value <> ""
        s.Cells(j / 2 - 1, i).value = g.Cells(j, 3).value
        j = j + 2
    Wend

    Call consume_buffers(pos_actuelle, i)

End Sub
