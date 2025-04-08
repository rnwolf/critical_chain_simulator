Attribute VB_Name = "resets"
'effacer les données d'avancement
Sub reinit_avancement()

    Dim answer As Integer
    
    answer = MsgBox("Cette action va supprimer toutes vos données d'avancement. Poursuivre?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmer la suppression")
        
    If answer = vbYes Then
        ThisWorkbook.Worksheets("LOGS_FV_CHART").Cells.Clear 'on supprime les données d'avancement
        'ThisWorkbook.Worksheets("LOGS_AV").Cells.Clear 'on supprime les données d'avancement
        
        'remise à jour de la partie à droite dans gantt les %
        Dim s As Worksheet
        Set s = ThisWorkbook.Worksheets("GANTT")
        Dim i As Integer: i = 6
        While s.Cells(i, 3).Interior.Color = vbWhite
            s.Cells(i, 3).value = 0
            i = i + 2
        Wend
        
        Dim sh As Worksheet
        Set s = ThisWorkbook.Worksheets("LOGS")
        Set sh = ThisWorkbook.Worksheets("LOGS_AV")
        
        Dim l As Integer
        l = s.Cells(15, 17).value + s.Cells(15, 16).value
        
        For i = 2 To l / 4
            sh.Columns(i).ClearContents
        Next i
        
        's.Range(s.Cells(22, 9), s.Cells(22 + i - 1, 9)).Copy sh.Range(sh.Cells(2, 1), s.Cells(2 + i - 1, 1))
        Call reinitialiser_GANTT_reel
    End If
    
End Sub


'remise à 0 de l'excel : suppression du projet
Sub reset_project()

    Dim answer As Integer
    
    answer = MsgBox("Cette action vas supprimer toutes vos informations. Poursuivre?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmer la suppression")
        
    If answer = vbYes Then

        Dim s As Worksheet
        Set s = ThisWorkbook.Worksheets("TÂCHES")
        
        Call retrieve_tasks
        Call retrieve_ressources
        
        Dim l As Integer: l = taches.Count + ressources.Count
        s.Range(s.Cells(10, 2), s.Cells(10 + l, 12)).Clear
        s.Range(s.Cells(10, 2), s.Cells(10 + l, 12)).Interior.Color = RGB(255, 242, 204)
        s.Cells(2, 1).value = ""
        
        
        Set s = ThisWorkbook.Worksheets("GANTT")
        
        Dim shp As Shape

        For Each shp In s.Shapes
           If shp.Type = 1 Then 'arrow type
                shp.Delete
            End If
        Next shp
        
        Dim e As Integer
        e = ThisWorkbook.Worksheets("LOGS").Cells(2, 1).value * 2
        s.Range(s.Cells(6, 1), s.Cells(6 + l * 2, e)).UnMerge
        s.Range(s.Cells(6, 1), s.Cells(6 + l * 2, e)).Clear
        s.Range(s.Cells(6, 1), s.Cells(6 + l * 2, e)).Interior.Color = RGB(255, 242, 204)
        
        Set s = ThisWorkbook.Worksheets("DASHBOARD")
        s.Range(s.Cells(6, 1), s.Cells(6 + l * 2, e)).UnMerge
        s.Range(s.Cells(6, 1), s.Cells(6 + l * 2, e)).Clear
        s.Range(s.Cells(6, 1), s.Cells(6 + l * 2, e)).Interior.Color = RGB(255, 242, 204)
        
        ThisWorkbook.Worksheets("LOGS_FV_CHART").Cells.Clear
        ThisWorkbook.Worksheets("LOGS_AV").Cells.Clear
        
        'supprimer les graphs
        With ThisWorkbook.Worksheets("DASHBOARD")
            While .ChartObjects.Count > 1
                .ChartObjects(.ChartObjects.Count).Delete
            Wend
            .ChartObjects(1).Chart.ChartTitle.Text = "Buffer chaîne critique"
        End With
        
        MsgBox "La réinitialisation a supprimé la date de début. Veuillez indiquer la date de lancement du projet en cellule A2!"
        Call reinitialiser_GANTT_reel
    End If
    
End Sub
Sub reinitialiser_GANTT_reel()
    
    Call retrieve_tasks
    Dim k As Worksheet, s As Worksheet, sh As Worksheet
    Dim i As Integer, j As Integer, t As Collection, marge As Integer, m As Integer
    Set s = ThisWorkbook.Worksheets("GANTT")
    Set k = ThisWorkbook.Worksheets("LOGS")
    Set sh = ThisWorkbook.Worksheets("LOGS_CCPM")
    Set t = taches
    marge = 6
    'For j = 0 To UBound(taches)
    '    t.Add taches(j)
    'Next j
    j = 0
    Dim nb_chaines As Integer
    nb_chaines = 0
    
    While k.Cells(j + 15, 15) <> "" ' calculer le nombre de chaînes qu'il y a
        nb_chaines = nb_chaines + 1
        j = j + 1
    Wend
    m = ThisWorkbook.Worksheets("LOGS").Cells(2, 1).value + 3 'marge + 4 avant
    j = 7
    For i = 0 To t.Count + nb_chaines - 1
        Range(s.Cells(j - 1, marge), s.Cells(j - 1, m - 1)).Interior.Pattern = xlPatternSolid
        Range(s.Cells(j, marge), s.Cells(j, m - 1)).Interior.ColorIndex = 2
        Range(s.Cells(j, m + 1), s.Cells(j, m + 50)).Interior.Color = RGB(255, 242, 204)
        Range(s.Cells(j, marge), s.Cells(j, m)).ClearContents
        
        j = j + 2
    Next i
    
    'Range(s.Cells(6, 3), s.Cells(t.Count * 2 + 5, 3)) = 0 ' Permet de remettre à 0 tous les avancements saisis. Non demandé par le client.
    'Range(s.Cells(6, 3), s.Cells(t.Count * 2 + 5, 3)).Interior.ColorIndex = xlColorIndexNone ' Remise de la couleur à 0 (gris si completion).
    Range(k.Cells(26, 4), k.Cells(t.Count + 25, marge)).ClearContents
    Range(sh.Cells(2, 3), sh.Cells(250, 5)).ClearContents
End Sub
