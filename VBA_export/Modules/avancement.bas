Attribute VB_Name = "avancement"
Sub ordonnancement_CCPM()

    Call retrieve_tasks
    Call record_avancement
    Call couleur_avancement
    Dim t As Collection, s As Worksheet, k As Worksheet, sh As Worksheet, cc As Worksheet, i As Integer, j As Integer, l As Integer, m As Integer, variable As Integer
    Dim case_debut As Integer, case_fin As Integer, avancement As Double, date_actuelle As Integer, vertical_margin As Integer, marge As Integer
    Dim check As Boolean, u As Integer, max As Integer
    Dim temps_theorique As Integer
    Dim case_debut_theorique As Integer 'numéro de colonne
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
    Set t = taches 'commence à indice 1
    marge = 6 'colonnes
    vertical_margin = 6 ' lignes
    marge_fin = ThisWorkbook.Worksheets("LOGS").Cells(2, 1).value + 3 'marge + 4 avant
    
    j = 0
    date_actuelle = colonne_date_actuelle
    max = 0
    Range(cc.Cells(2, 3), cc.Cells(250, 4)).ClearContents 'Clean des valeurs de décalage
    If date_actuelle = 0 Then
        Exit Sub
    End If

    For i = 1 To t.Count ' set les débuts
        
        l = 1
        While k.Cells(21 + l, 9) <> t(i).get_ID
            l = l + 1
        Wend
        t(i).set_debut (k.Cells(21 + l, 10) / 2 + marge)
        
    Next i ' set les débuts
    
    Dim nb_chaines As Integer
    nb_chaines = 0
    
    While k.Cells(j + 15, 15) <> "" ' calculer le nombre de chaînes qu'il y a
        nb_chaines = nb_chaines + 1
        j = j + 1
    Wend

    j = 0
    While k.Cells(j + 15, 15) <> ""
        
        '--------------------------------- Chaîne critique -----------------------------------
        If j = 0 Then ' On est dans la chaîne critique
            Debug.Print "Chaîne critique"
            l = 0
            While sh.Cells(17 + l, 6) <> "" 'prendre le dernier % de conso buffer
                l = l + 1
            Wend
                
            pourcentage_conso_buffer = sh.Cells(17 + l - 1, 6) ' C'est un pourcentage
            Debug.Print "Pourcentage buffer consommé " & pourcentage_conso_buffer
                
            If pourcentage_conso_buffer > 0 Then ' on a consommé du buffer
            
                duree_buffer = CInt(k.Cells(15, 16) / 4)
                conso_buffer = pourcentage_conso_buffer / 100 * duree_buffer
                Debug.Print "buffer consommé " & conso_buffer
                cc.Cells(t.Count + 2, 4) = conso_buffer
                    
                For m = 1 To t.Count ' Calcul des décalages
                    indice_ligne = trouver_ligne_indice(t(m).get_ID)
                    avancement = s.Cells(indice_ligne, 3)
                    
                    decalage = conso_buffer
                    sauv = cc.Cells(1 + t(m).get_ID, 3)
                    cc.Cells(1 + t(m).get_ID, 3) = sauv + decalage
                        
                Next m
                
                '------------- Décalage des autres buffers  -----------
                
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

            End If ' on a consommé du buffer
            
            '----------------------------- Chaîne secondaire -----------------------------------
            
        Else ' dans une chaîne secondaire
            
            Debug.Print "Chaîne non critique numéro " & j
            l = 0
    
            While sh.Cells(17 + l, 4 * (j + 1) + 2) <> "" 'prendre le dernier % de conso buffer, avant j ct test
                l = l + 1
            Wend
                
            pourcentage_conso_buffer = sh.Cells(17 + l - 1, 4 * (j + 1) + 2) ' C'est un pourcentage
            Debug.Print "Pourcentage buffer consommé " & pourcentage_conso_buffer
            If pourcentage_conso_buffer > 0 Then ' on a consommé du buffer
                
                'Paramètrage buffer
                duree_buffer = CInt(k.Cells(15 + j, 16) / 4)
                conso_buffer = pourcentage_conso_buffer / 100 * duree_buffer
                cc.Cells(t.Count + 2 + j, 4) = conso_buffer
                debut_buffer = k.Cells(15 + j, 17) 'On récupère la date de début
                fin_buffer = debut_buffer + CInt(k.Cells(15 + j, 16) / 4)
                Debug.Print "Début buffer " & debut_buffer & " et fin buffer " & fin_buffer
                
                splito = Split(k.Cells(15 + j, 15), ",")
                
                        '------------------- Surconsommation buffer -------------------
                
                
                If pourcentage_conso_buffer > 100 Then ' On a surconsommé le buffer, on décale toutes les tâches qui date d'après la fin du buffer
                        
                    For m = 1 To t.Count ' On décale toutes les tâches après la fin du buffer de la chaîne
                            
                        indice_ligne = trouver_ligne_indice(t(m).get_ID) 'On trouve sa ligne
                        avancement = s.Cells(indice_ligne, 3)

                        If fin_buffer < t(m).get_debut Then ' tâche qui commence après la fin du buffer
                            decalage = conso_buffer
                            sauv = cc.Cells(1 + t(m).get_ID, 3)
                            cc.Cells(1 + t(m).get_ID, 3) = sauv + decalage
                                    
                        End If ' tâche commence après fin buffer
                    Next m
                      
                    '----- Décalage des autres buffers si needed ---------
                    For m = 1 To nb_chaines
                        indice_ligne = trouver_ligne_indice(t.Count + m) 'On trouve sa ligne
                        'avancement = s.Cells(indice_ligne, 3)
                        
                        If m <> j + 1 Then ' pas décaler la chaîne même du buffer
                            If fin_buffer < k.Cells(15 + m - 1, 17) Then ' buffer qui commence après la fin du buffer
                                decalage = conso_buffer
                                sauv = cc.Cells(1 + t.Count + m, 3)
                                cc.Cells(1 + t.Count + m, 3) = sauv + decalage
                                Debug.Print "On decale la chaine " & m - 1 & " de " & decalage & "indice " & 1 + t.Count + m
                            End If ' buffer commence après fin buffer de notre tâche
                        End If
                    Next m
                            ' ------------- Consommation non intégrale -------------
                            
                Else ' on a consommé mais pas entièrement, on décale les tâches de la chaîne
                
                    Debug.Print "Lbound splito = " & LBound(splito) & "Ubound splito = " & UBound(splito)
                    For m = LBound(splito) To UBound(splito) 'pour les tâches de la chaine, pas le buffer dans tous les cas donc osef de trier
                            
                        indice_ligne = trouver_ligne_indice(CInt(splito(m))) 'On trouve la ligne de la tâche dans l'onglet GANTT
                        avancement = s.Cells(indice_ligne, 3)
                        
                        decalage = conso_buffer
                        sauv = cc.Cells(1 + CInt(splito(m)), 3)
                        cc.Cells(1 + CInt(splito(m)), 3) = sauv + decalage
                        
                    Next m
                        
                End If ' conso buff. Si inférieur ou égale à 0, rien ne se passe.
  
            End If 'chaîne secondaire, les indep seront en dehors du while
                        
        End If 'dans la chaîne critique
        j = j + 1
    Wend
    
    '------------- Traçage des tâches -------------
    
    For m = 1 To t.Count
    
        indice_ligne = trouver_ligne_indice(t(m).get_ID) 'On trouve sa ligne
        avancement = s.Cells(indice_ligne, 3)
        fini = cc.Cells(1 + t(m).get_ID, 5)
        If fini = 0 Then
            chaine = dans_quel_chaine(t(m).get_ID)
            Debug.Print "Tâche " & t(m).get_ID & " dans chaîne " & chaine
            
            case_debut = t(m).get_debut + avancement * t(m).get_duree / 2 + cc.Cells(1 + t(m).get_ID, 3)
            case_fin = case_debut + (1 - avancement) * t(m).get_duree / 2 - 1
            
            If avancement = 1 Then ' Dernier parcours de cette tâche
                cc.Cells(1 + t(m).get_ID, 5) = 1
            End If
            
            Debug.Print "tache " & t(m).get_ID & "case_deb " & case_debut & " fin " & case_fin & " oui " & t(m).get_duree / 2 - 1 & "avancement " & 1 - avancement & " donc " & (1 - avancement) * t(m).get_duree / 2 - 1
            If case_fin < 0 Then
                MsgBox "Veuillez vérifier la valeur des avancements saisies svp."
                Exit Sub
            End If
            'Clean
            Range(s.Cells(indice_ligne + 1, marge), s.Cells(indice_ligne + 1, marge_fin)).ClearContents
            Range(s.Cells(indice_ligne + 1, marge), s.Cells(indice_ligne + 1, marge_fin)).Interior.ColorIndex = 2
            Range(s.Cells(indice_ligne, CInt(t(m).get_debut)), s.Cells(indice_ligne, CInt(t(m).get_debut + t(m).get_duree / 2))).Interior.Pattern = xlPatternSolid
            
            'Traçage
            If avancement <> 1 Then
                s.Cells(indice_ligne + 1, case_debut) = t(m).get_ID ' Numéroter la tâche
            
                If chaine = 0 Then ' gestion des couleurs
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
                ElseIf chaine = -1 Then 'la tâche n'est pas dans une chaîne
                    If avancement <> 0 Then
                        Range(s.Cells(indice_ligne, CInt(t(m).get_debut)), s.Cells(indice_ligne, CInt(t(m).get_debut + t(m).get_duree / 2 - 1))).Interior.ColorIndex = 5
                        If avancement >= 1 Then
                            Range(s.Cells(indice_ligne, CInt(t(m).get_debut)), s.Cells(indice_ligne, CInt(t(m).get_debut + t(m).get_duree / 2 - 1))).Interior.Pattern = xlPatternLightUp
                        Else
                        Range(s.Cells(indice_ligne, CInt(t(m).get_debut)), s.Cells(indice_ligne, CInt(t(m).get_debut + avancement * t(m).get_duree / 2 - 1))).Interior.Pattern = xlPatternLightUp
                        End If
                    End If
                    Range(s.Cells(indice_ligne + 1, case_debut), s.Cells(indice_ligne + 1, case_fin)).Interior.ColorIndex = 34
                    
                Else ' la tâche est dans une chaîne secondaire
                    If avancement <> 0 Then
                        Range(s.Cells(indice_ligne, CInt(t(m).get_debut)), s.Cells(indice_ligne, CInt(t(m).get_debut + t(m).get_duree / 2 - 1))).Interior.ColorIndex = 4
                        If avancement >= 1 Then
                            Range(s.Cells(indice_ligne, CInt(t(m).get_debut)), s.Cells(indice_ligne, CInt(t(m).get_debut + t(m).get_duree / 2 - 1))).Interior.Pattern = xlPatternLightUp
                        Else
                        Range(s.Cells(indice_ligne, CInt(t(m).get_debut)), s.Cells(indice_ligne, CInt(t(m).get_debut + avancement * t(m).get_duree / 2 - 1))).Interior.Pattern = xlPatternLightUp
                        End If
                        
                    End If
                    Range(s.Cells(indice_ligne + 1, case_debut), s.Cells(indice_ligne + 1, case_fin)).Interior.ColorIndex = 35
                End If ' quelle chaine (gestion des couleurs)
            Else 'avancement égale à 1
                Range(s.Cells(indice_ligne, CInt(t(m).get_debut)), s.Cells(indice_ligne, CInt(t(m).get_debut + t(m).get_duree / 2 - 1))).Interior.Pattern = xlPatternLightUp
                Range(s.Cells(indice_ligne + 1, marge), s.Cells(indice_ligne + 1, marge_fin)).ClearContents
                Range(s.Cells(indice_ligne + 1, marge), s.Cells(indice_ligne + 1, marge_fin)).Interior.ColorIndex = 2
                s.Cells(indice_ligne, 3).Interior.ColorIndex = 15
            
            End If 'Avancement différent de 1
        Else
            Range(s.Cells(indice_ligne, CInt(t(m).get_debut)), s.Cells(indice_ligne, CInt(t(m).get_debut + t(m).get_duree / 2 - 1))).Interior.Pattern = xlPatternLightUp
            Range(s.Cells(indice_ligne + 1, marge), s.Cells(indice_ligne + 1, marge_fin)).ClearContents
            Range(s.Cells(indice_ligne + 1, marge), s.Cells(indice_ligne + 1, marge_fin)).Interior.ColorIndex = 2
            s.Cells(indice_ligne, 3).Interior.ColorIndex = 15
        End If ' Fini
    Next m
    
    '------------- Traçage des buffers + leur conso -----------
    For m = 1 To nb_chaines
        indice_ligne = trouver_ligne_indice(t.Count + m)
        debut_buffer = k.Cells(15 + m - 1, 17) / 2 + 6 + cc.Cells(t.Count + m + 1, 3) 'On récupère la date de début et on ajoute le décalage (dû à autres chaînes), /2 +6 pr conversion heures en colonne
        'debut_buffer = debut_buffer / 2 + 6
        conso_buffer = cc.Cells(t.Count + m + 1, 4)
        duree_buffer = CInt(k.Cells(15 + m - 1, 16) / 4)
        If debut_buffer = 0 Then 'Protection
            MsgBox "Problème rencontré, veuillez réactualiser le GANTT classique svp."
            Exit Sub
        End If 'Protection
                
        'Debug.Print "Début buffer chaîne " & m - 1 & " à " & debut_buffer & " indice de ligne " & indice_ligne & " et conso " & conso_buffer & " et durée " & duree_buffer
        'Clean
        Range(s.Cells(indice_ligne + 1, marge), s.Cells(indice_ligne + 1, marge_fin)).ClearContents
        Range(s.Cells(indice_ligne + 1, marge), s.Cells(indice_ligne + 1, marge_fin)).Interior.ColorIndex = 2
        'Traçage
        If conso_buffer <> 0 Then
            Range(s.Cells(indice_ligne + 1, debut_buffer), s.Cells(indice_ligne + 1, debut_buffer + conso_buffer - 1)).Interior.ColorIndex = 1 ' y'a un -1 à mettre mais jsp pq
        Else
            s.Cells(indice_ligne + 1, debut_buffer).Interior.ColorIndex = 15
            'Range(s.Cells(indice_ligne + 1, debut_buffer), s.Cells(indice_ligne + 1, debut_buffer + duree_buffer - 1)).Interior.ColorIndex = 15 'rajouté récemment
        End If
        If m = 1 Then ' numéroter
            s.Cells(indice_ligne + 1, debut_buffer).Font.ColorIndex = 2
            s.Cells(indice_ligne + 1, debut_buffer) = "Buffer chaîne critique" ' Numéroter le buffer
            
        Else
            s.Cells(indice_ligne + 1, debut_buffer) = "Buffer " & m - 1 ' Numéroter le buffer
            s.Cells(indice_ligne + 1, debut_buffer).Font.ColorIndex = 2
        End If ' numéroter
        If conso_buffer < duree_buffer Then 'Si on a pas tout consommé
            Range(s.Cells(indice_ligne + 1, debut_buffer + conso_buffer), s.Cells(indice_ligne + 1, debut_buffer + duree_buffer - 1)).Interior.ColorIndex = 15 'pareil pr le -1
        End If ' si on a pas tt consommé
    Next m
    
End Sub


'écriture en logs et appel de la conso buffer
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
