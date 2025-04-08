Attribute VB_Name = "affichages"

'-----------GANTT------------'

Public Function affichage_GANTT(t_triees As Collection)

    Dim s As Worksheet, i As Integer, j As Integer, ligne As Integer, diff As Integer, case_debut As Integer, case_fin As Integer
    Set s = ThisWorkbook.Worksheets("GANTT")
    Dim sheets As Collection
    Set sheets = New Collection
    sheets.Add s
    sheets.Add ThisWorkbook.Worksheets("DASHBOARD")
    Dim sh As Worksheet
    
    'Clean de l'ancien GANTT et des graphs
    Call GANTT_clear
    'Range(s.Cells(GANTT_vertical_margin, GANTT_horizontal_margin), s.Cells(t_triees.Count * 2 + 5, last_task(t_triees) / 2 + 8)).Clear
    For Each sh In sheets
        Range(sh.Cells(GANTT_vertical_margin, 1), sh.Cells(t_triees.Count * 2 + 5, 3)).Clear
    Next
    
    Dim shp As Shape

    For Each shp In s.Shapes
       If shp.Type = 1 Then 'arrow type
            shp.Delete
        End If
    Next shp
    
    Call creer_calendrier(t_triees)
    
    'affichage taches à gauche
    For Each sh In sheets
        For j = 1 To 3
            For i = 1 To t_triees.Count * 2 Step 2
                Dim r As Range
                Set r = Range(sh.Cells(i + GANTT_vertical_margin - 1, j), sh.Cells(i + GANTT_vertical_margin, j))
                r.Borders.LineStyle = xlContinuous
                r.Merge
                r.HorizontalAlignment = xlCenter
                r.VerticalAlignment = xlCenter
                If j = 3 Then
                    sh.Cells(i + GANTT_vertical_margin - 1, j).NumberFormat = "0.00%"
                    sh.Cells(i + GANTT_vertical_margin - 1, j) = 0
                End If
            Next i
        Next j
    Next
    
    
    ' Calcul du nombre de chaînes
    Dim nb_chaines As Integer
    Dim ka As Worksheet
    Set ka = ThisWorkbook.Worksheets("LOGS")
    nb_chaines = 0
    j = 0
    
    While ka.Cells(j + 15, 15) <> ""
        nb_chaines = nb_chaines + 1
        j = j + 1
    Wend
    
    'Enregistrer la position du buffer chaîne critique
    Dim row_buffer As Integer, column_buffer As Integer
    'row_buffer = trouver_ligne_indice(t_triees.Count - nb_chaines + 1)
    column_buffer = ka.Cells(15, 17) / 2 + 6
    Dim condition As Boolean
    condition = False
    i = 1
    While condition = False
        If t_triees(i).get_ID = t_triees.Count - nb_chaines + 1 Then
            condition = True
        Else
            i = i + 1
        End If
    Wend
    row_buffer = 6 + (i - 1) * 2
    
    
    'Affichage du nouveau GANTT
    For i = 1 To t_triees.Count
        ligne = i * 2 + 4
        diff = CInt(t_triees(i).get_fin) - CInt(t_triees(i).get_debut)
        case_debut = CInt(t_triees(i).get_debut) / 2 + GANTT_horizontal_margin
        case_fin = CInt(t_triees(i).get_fin) / 2 + GANTT_horizontal_margin - 1
            
        'If i <> 1 Then
            'Call DrawArrows(Range(Cells(ligne - 2, case_debut - 1), Cells(ligne - 2, case_debut - 1)), Range(Cells(ligne, case_debut), Cells(ligne, case_debut)))
            If t_triees(i).get_preds <> "" Then
                Dim preds() As String
                preds = Split(t_triees(i).get_preds, ",")
                Dim k As Integer
                For k = 0 To UBound(preds)
                    Dim pred_case_fin As Integer, indice_pred As Integer
                    indice_pred = get_task_index_by_id(CInt(preds(k)), t_triees)
                    
                    If t_triees(indice_pred).get_type <> 4 Then
                        pred_case_fin = CInt(t_triees(indice_pred).get_fin) / 2 + GANTT_horizontal_margin - 1
                        Call draw_arrows(indice_pred * 2 + 4, pred_case_fin, ligne, case_debut)
                    End If
                    
                Next k
            ElseIf t_triees(i).get_type = 3 Then ' Traçage des flèches pour les tâches bleues
                'MsgBox "In it"
                Call draw_arrows(ligne, case_fin, row_buffer, column_buffer)
            End If
            
            
        'End If
            
        For Each sh In sheets
            sh.Cells(ligne, 1) = t_triees(i).get_ID
            sh.Cells(ligne, 2) = t_triees(i).get_Intitule
        Next
        ThisWorkbook.Worksheets("LOGS_AV").Cells(i + 1, 1).value = t_triees(i).get_ID
        s.Cells(ligne, case_debut) = t_triees(i).get_ID
        For j = case_debut To case_fin
            Select Case t_triees(i).get_type
                Case Is = 1
                    s.Cells(ligne, j).Interior.Color = RGB(255, 0, 0)
                Case Is = 2
                    s.Cells(ligne, j).Interior.Color = RGB(0, 255, 0)
                Case Is = 3
                    s.Cells(ligne, j).Interior.Color = RGB(0, 0, 255)
                Case Is = 4
                    s.Cells(ligne, j).Interior.Color = RGB(200, 200, 200)
            End Select
        Next j
    Next i

    s.Range("P2") = s.Cells(4, ThisWorkbook.Worksheets("LOGS").Cells(2, 1).value).value 'date de fin estimée
    Call create_charts
    Call retrieve_fv_points
    s.Select

End Function


Sub creer_calendrier(t_triees)
    
        'on travaille 8h 5 jours sur 7 (jours ouvrés)
        'dans le calendrier on va sauter les weekends pour que la date de fin soit réaliste
    
    Dim s As Worksheet, i As Integer, j  As Date, r As Range
    Set s = ThisWorkbook.Worksheets("GANTT")
    j = s.Range("C2")
    i = 1
    Dim case_fin As Integer
    While i < (last_task(t_triees) / 2 + 4)
        s.Cells(GANTT_vertical_margin - 2, i + GANTT_horizontal_margin - 1) = Format(j, "dd.mm.yy")
        Set r = Range(s.Cells(GANTT_vertical_margin - 2, i + GANTT_horizontal_margin - 1), s.Cells(GANTT_vertical_margin - 2, i + GANTT_horizontal_margin + 2))
        case_fin = i + GANTT_horizontal_margin - 1
        r.Merge
        r.HorizontalAlignment = xlCenter
        r.VerticalAlignment = xlCenter
        r.Interior.Color = RGB(255, 242, 204)
        
        Range(s.Cells(GANTT_vertical_margin - 2, i + GANTT_horizontal_margin - 1), s.Cells(t_triees.Count * 2 + 5, i + GANTT_horizontal_margin + 2)).BorderAround (xlDash)
        
        r.Borders.LineStyle = xlContinuous
        r.Borders.Weight = xlThin
        With r.Borders(xlEdgeTop)
            .Weight = xlMedium
        End With
        
        If i = 1 Then
            With Range(s.Cells(GANTT_vertical_margin - 2, i + GANTT_horizontal_margin - 1), s.Cells(t_triees.Count * 2 + 5, i + GANTT_horizontal_margin + 2)).Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
        End If
        If Format(j, "w", 2) = 1 Then
            With Range(s.Cells(GANTT_vertical_margin - 2, i + GANTT_horizontal_margin - 1), s.Cells(t_triees.Count * 2 + 5, i + GANTT_horizontal_margin + 2)).Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
        End If
        
        i = i + 4
        'dernière colonne
        If i > (last_task(t_triees) / 2) Then
            With Range(s.Cells(GANTT_vertical_margin - 2, i - 4 + GANTT_horizontal_margin - 1), s.Cells(t_triees.Count * 2 + 5, i - 4 + GANTT_horizontal_margin + 2)).Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
        End If
        
        
        'End If
        'MsgBox Format(j, "w", 2)
        If Format(j, "w", 2) < 5 Then
            j = j + 1
        Else
            j = j + 3 'saut du weekend
            'Range(s.Cells(GANTT_vertical_margin - 2, i + GANTT_horizontal_margin - 1), s.Cells(t_triees.Count * 2 + 5, i + GANTT_horizontal_margin + 2)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        End If
    Wend
    
    ThisWorkbook.Worksheets("LOGS").Cells(2, 1) = case_fin
    
    'quadrillage horizontal
    Dim k As Integer
    For k = 1 To t_triees.Count * 2 Step 2
        Range(s.Cells(k + GANTT_vertical_margin - 1, GANTT_horizontal_margin), s.Cells(k + GANTT_vertical_margin, i + 4)).Interior.Color = vbWhite
        Range(s.Cells(k + GANTT_vertical_margin - 1, GANTT_horizontal_margin), s.Cells(k + GANTT_vertical_margin, i + 4)).BorderAround (xlDash)
        With Range(s.Cells(k + GANTT_vertical_margin - 1, GANTT_horizontal_margin), s.Cells(k + GANTT_vertical_margin, i + 4)).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With
        With Range(s.Cells(k + GANTT_vertical_margin - 1, GANTT_horizontal_margin), s.Cells(k + GANTT_vertical_margin, i + 4)).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
        End With
        
        If k >= t_triees.Count * 2 - 2 Then
            With Range(s.Cells(k + GANTT_vertical_margin - 1, GANTT_horizontal_margin), s.Cells(k + GANTT_vertical_margin, i + 4)).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlMedium
            End With
        End If
    Next k
End Sub


Sub GANTT_clear()

    Dim i As Integer: i = GANTT_vertical_margin
    Dim j As Integer: j = GANTT_horizontal_margin
    'While Cells(GANTT_vertical_margin - 2, j).value <> RGB(255, 242, 204)
    '    j = j + 1
    'Wend
    
    Dim s As Worksheet
    Set s = ThisWorkbook.Worksheets("GANTT")
    While s.Cells(i, 3).Interior.Color = vbWhite
        i = i + 2
    Wend
    j = ThisWorkbook.Worksheets("LOGS").Cells(2, 1).value + 4 'marge + 4
    
    Range(s.Cells(GANTT_vertical_margin, 1), s.Cells(i, j)).UnMerge
    Range(s.Cells(GANTT_vertical_margin, 1), s.Cells(i, j)).Clear
    Range(s.Cells(GANTT_vertical_margin, 1), s.Cells(i, j)).Interior.Color = RGB(255, 242, 204)
    
    Dim s2 As Worksheet
    Set s2 = ThisWorkbook.Worksheets("DASHBOARD")
    
    ThisWorkbook.Worksheets("LOGS_FV_CHART").Cells.Clear 'on supprime les données d'avancement
    
    'supprimer les graphs
    With ThisWorkbook.Worksheets("DASHBOARD")
        'suppression tâches à gauche
        .Range(.Cells(GANTT_vertical_margin, 1), .Cells(i, j)).UnMerge
        .Range(.Cells(GANTT_vertical_margin, 1), .Cells(i, j)).Clear
        .Range(.Cells(GANTT_vertical_margin, 1), .Cells(i, j)).Interior.Color = RGB(255, 242, 204)
        While .ChartObjects.Count > 1
            .ChartObjects(.ChartObjects.Count).Delete
        Wend
        .ChartObjects(1).Chart.ChartTitle.Text = "Buffer projet"
    End With

End Sub


Sub draw_arrows(c1_row As Integer, c1_column As Integer, c2_row As Integer, c2_column As Integer)

    Dim s As Worksheet
    Dim from As Range, toRange As Range
    
    Set s = ThisWorkbook.Worksheets("GANTT")
    Set from = Range(s.Cells(c1_row, c1_column), s.Cells(c1_row, c1_column))
    Set toRange = Range(s.Cells(c2_row, c2_column), s.Cells(c2_row, c2_column))
    
    s.Select
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, from.Left + from.Width, from.Top + from.Height / 2, toRange.Left, toRange.Top + toRange.Height / 2).Select
    
    'format line
    With Selection.ShapeRange.Line
        .EndArrowheadStyle = msoArrowheadTriangle
        .Weight = 1
        .ForeColor.RGB = RGB(0, 0, 0)
    End With
 
End Sub


'-----------FEVER CHARTS------------'

'affichage d'un graph, titre en paramètre
Sub add_chart(n As String)
    Dim c As ChartObject
    'Set c = Charts.Add
    With ThisWorkbook.Worksheets("DASHBOARD")
    
        Dim i As Integer
        i = .ChartObjects.Count - 1
        Set c = .ChartObjects.Add(.Cells(6 + 18 * i, 38).Left, .Cells(6 + 18 * i, 38).Top, 200, 100)
        
        'size
        .ChartObjects(.ChartObjects.Count).Width = .ChartObjects(.ChartObjects.Count - 1).Width
        .ChartObjects(.ChartObjects.Count).Height = .ChartObjects(.ChartObjects.Count - 1).Height
        
        i = .ChartObjects.Count
        
        With .ChartObjects(.ChartObjects.Count).Chart
            'defininf title
            .HasTitle = True
            .ChartTitle.Text = "Chaîne : " + n
            
            'définition des zones colorées
            .SeriesCollection.NewSeries
            .SeriesCollection(1).ChartType = xlAreaStacked
            .SeriesCollection(1).XValues = ThisWorkbook.Worksheets("LOGS").Range("H2:H12")
            .SeriesCollection(1).Values = ThisWorkbook.Worksheets("LOGS").Range("I2:I12")
            .SeriesCollection(1).Interior.Color = RGB(146, 208, 80)
            
            .SeriesCollection.NewSeries
            .SeriesCollection(2).ChartType = xlAreaStacked
            .SeriesCollection(2).XValues = ThisWorkbook.Worksheets("LOGS").Range("H2:H12")
            .SeriesCollection(2).Values = ThisWorkbook.Worksheets("LOGS").Range("J2:J12")
            .SeriesCollection(2).Interior.Color = RGB(255, 255, 0)
            
            .SeriesCollection.NewSeries
            .SeriesCollection(3).ChartType = xlAreaStacked
            .SeriesCollection(3).XValues = ThisWorkbook.Worksheets("LOGS").Range("H2:H12")
            .SeriesCollection(3).Values = ThisWorkbook.Worksheets("LOGS").Range("K2:K12")
            .SeriesCollection(3).Interior.Color = RGB(255, 0, 0)
            
            .HasLegend = False 'no legend to series
            
            'titres des axes
            .Axes(xlCategory).HasTitle = True
            .Axes(xlCategory).AxisTitle.Text = "% avancement de la chaîne"
            .Axes(xlValue).HasTitle = True
            .Axes(xlValue).AxisTitle.Text = "% consommation du buffer"
            .Axes(xlValue).MaximumScale = 100
            
            'la coubre de consommation
            .SeriesCollection.NewSeries
            .SeriesCollection(4).ChartType = xlXYScatterLines
            
            Dim range_string As String
            Dim column_letter As String
            column_letter = Split(Cells(1, 4 * i + 3).Address, "$")(1) 'convert index to letter
            range_string = column_letter + CStr(16) + ":" + column_letter + CStr(16 + ThisWorkbook.Worksheets("LOGS").Cells(2, 1).value)
            .SeriesCollection(4).XValues = ThisWorkbook.Worksheets("LOGS_FV_CHART").Range(range_string)
            
            column_letter = Split(Cells(1, 4 * i + 2).Address, "$")(1)
            range_string = column_letter + CStr(16) + ":" + column_letter + CStr(16 + ThisWorkbook.Worksheets("LOGS").Cells(2, 1).value)
            .SeriesCollection(4).Values = ThisWorkbook.Worksheets("LOGS_FV_CHART").Range(range_string)
            
            'style
            .SeriesCollection(4).Format.Line.ForeColor.RGB = RGB(0, 0, 0)
            .SeriesCollection(4).MarkerStyle = xlMarkerStyleCircle
            .SeriesCollection(4).MarkerBackgroundColor = RGB(0, 0, 0)
        
        End With
        
    End With
End Sub


'genèrer un graph par chaîne
Sub create_charts()

    With ThisWorkbook.Worksheets("DASHBOARD")
        'updating critical chain
        .ChartObjects(1).Visible = True
        .ChartObjects(1).Chart.ChartTitle.Text = "Buffer projet (" + CStr(ThisWorkbook.Worksheets("LOGS").Cells(15, 15).value) + " )"
        
        'selecting values
        Dim range_string As String
        Dim column_letter As String
        column_letter = Split(Cells(1, 7).Address, "$")(1)
        range_string = column_letter + CStr(16) + ":" + column_letter + CStr(16 + ThisWorkbook.Worksheets("LOGS").Cells(2, 1).value)
        .ChartObjects(1).Chart.SeriesCollection(4).XValues = ThisWorkbook.Worksheets("LOGS_FV_CHART").Range(range_string)
        column_letter = Split(Cells(1, 6).Address, "$")(1)
        range_string = column_letter + CStr(16) + ":" + column_letter + CStr(16 + ThisWorkbook.Worksheets("LOGS").Cells(2, 1).value)
        .ChartObjects(1).Chart.SeriesCollection(4).Values = ThisWorkbook.Worksheets("LOGS_FV_CHART").Range(range_string)
    

        Dim s As Worksheet
        Set s = ThisWorkbook.Worksheets("LOGS")
        Dim i As Integer: i = 16 'on commence après la chaîne critique
        While s.Cells(i, 15).value <> 0
            add_chart (ThisWorkbook.Worksheets("LOGS").Cells(i, 15).value)
            i = i + 1
        Wend
    End With
End Sub

Sub couleur_avancement()
    Dim s As Worksheet
    Set s = ThisWorkbook.Worksheets("GANTT")
    Dim k As Worksheet
    Set k = ThisWorkbook.Worksheets("LOGS")
    Dim i As Integer, j As Integer
    j = 0
    Dim t As Collection
    Call retrieve_tasks
    Set t = taches
    
    Dim nb_chaines As Integer
    nb_chaines = 0
    
    While k.Cells(j + 15, 15) <> ""
        nb_chaines = nb_chaines + 1
        j = j + 1
    Wend
    
    For i = 1 To t.Count + nb_chaines - 1
        If s.Cells(6 + j, 3) = 1 Then
            s.Cells(6 + j, 3).Interior.ColorIndex = 15
        Else
            s.Cells(6 + j, 3).Interior.ColorIndex = 2
        End If
      
        
        If InStr(CStr(s.Cells(6 + j, 2)), "Buffer") <> 0 Then
            s.Cells(6 + j, 3).Interior.ColorIndex = 15
            s.Cells(6 + j, 3) = "//"
        End If
        j = j + 2
    Next i
End Sub
