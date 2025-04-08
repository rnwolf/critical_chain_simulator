Attribute VB_Name = "utils"
Option Explicit
Public ressources As Collection
Public taches As Collection
Public GANTT_vertical_margin As Integer
Public GANTT_horizontal_margin As Integer
Public RSC_vertical_margin As Integer
Public RSC_horizontal_margin As Integer
Public TSK_vertical_margin As Integer
Public TSK_horizontal_margin As Integer

Public planning_alert As Boolean 'used if an error is found in one of the recursivities


Sub ajouter_tache()
    ajout_tache.Show
End Sub


Sub ajouter_ressource()
    ajout_ressource.Show
End Sub

'---------RETRIEVES---------'

Sub retrieve_tasks() 'ByRef t() As Tache)
    Call retrieve_margins
    Dim s As Worksheet, i As Integer, m As Integer
    Set s = ThisWorkbook.Worksheets("T�CHES")
    Set taches = New Collection
    i = TSK_vertical_margin
    m = TSK_horizontal_margin
    While s.Cells(i, m).value <> ""
        Dim t As Tache
        Set t = New Tache
        t.set_attributes s.Cells(i, m + 1).value, s.Cells(i, m + 2).value * 8, s.Cells(i, m + 5).value, s.Cells(i, m + 4).value, s.Cells(i, m + 3).value * 8
        taches.Add t
        i = i + 1
    Wend
End Sub


Sub retrieve_ressources()
    Call retrieve_margins
    Dim s As Worksheet, i As Integer, m As Integer
    Set s = ThisWorkbook.Worksheets("T�CHES")
    Set ressources = New Collection
    i = RSC_vertical_margin
    m = RSC_horizontal_margin
    While s.Cells(i, m).value <> ""
        Dim r As Ressource
        Set r = New Ressource
        r.set_attributes s.Cells(i, m + 1).value 'We don't recover the tasks here
        ressources.Add r
        i = i + 1
    Wend
End Sub

Sub retrieve_margins()
    Dim s As Worksheet
    Set s = ThisWorkbook.Worksheets("LOGS")
    GANTT_horizontal_margin = s.Cells(6, 1).value
    GANTT_vertical_margin = s.Cells(8, 1).value
    RSC_horizontal_margin = s.Cells(6, 2).value 'Margins associated with the resource table
    RSC_vertical_margin = s.Cells(8, 2).value
    TSK_horizontal_margin = s.Cells(6, 3).value 'Margins associated with the table of tasks
    TSK_vertical_margin = s.Cells(8, 3).value
End Sub


Function retrieve_chains() As Collection

    Call retrieve_tasks

    Dim chains As Collection
    Set chains = New Collection
    Dim s As Worksheet
    Set s = ThisWorkbook.Worksheets("LOGS")
    Dim i As Integer: i = 15
    While s.Cells(i, 15).value <> 0
        Dim chain As Collection
        Set chain = New Collection
        Dim schain As String
        schain = s.Cells(i, 15).value
        Dim t() As String
        t = Split(schain, ",")
        Dim j As Integer
        For j = 0 To UBound(t)
            'If InStr(1, schain, CStr(taches(j).get_ID)) Then

                'Recovery of the ending square
                Dim k As Integer: k = 22
                'While s.Cells(k, 9).value <> CStr(taches(j).get_ID)
                While s.Cells(k, 9).value <> CInt(t(j))
                    k = k + 1
                Wend
                taches(CInt(t(j))).set_fin CInt(s.Cells(k, 11))
                taches(CInt(t(j))).set_debut CInt(s.Cells(k, 10))
                chain.Add taches(CInt(t(j)))
            'End If
        Next j
        chains.Add chain
        i = i + 1
    Wend

    Set retrieve_chains = chains

End Function


Sub retrieve_fv_points()

    Call retrieve_margins

    Dim i As Integer, u As Worksheet
    i = 2
    Set u = ThisWorkbook.Worksheets("LOGS_AV")
    While u.Cells(1, i).value <> ""
        Call consume_buffers(u.Cells(1, i).value, i)
        i = i + 1
    Wend

End Sub


Public Function print_taches()
    Dim res As String, i As Integer

    For i = 1 To taches.Count
        res = res + CStr(taches(i).get_ID)
    Next i
    print_taches = res
End Function


'---------OTHER---------'

'Returns the end date of the project
Function last_task(t_triees)

    Dim value As Integer, i As Integer

    'MsgBox t_triees(2).get_fin

    value = CInt(t_triees(1).get_fin)

    For i = 1 To t_triees.Count
        If CInt(t_triees(i).get_fin) > value Then
            value = CInt(t_triees(i).get_fin)
        End If
    Next i

    last_task = value

End Function


'returns a collection of tasks that are the antecedents of the target
Public Function antecedants(cible As Tache, t As Collection) As Object

    Dim i As Integer
    Set antecedants = New Collection
    For i = 1 To t.Count
        Dim res As Boolean: res = False
        Dim k As Integer, preds() As String
        preds = Split(t(i).get_preds, ",")

        For k = 0 To UBound(preds)
            If preds(k) = CStr(cible.get_ID) Then
                res = True
            End If
        Next k

        If res = True Then
            antecedants.Add t(i)
        End If


        'If InStr(1, t(i).get_preds, cible.get_ID) Then
        '    antecedants.Add t(i)
        'End If
    Next i

End Function

Sub remove_task_by_id(id As Integer, t As Collection)

    Dim i As Integer, res As Integer
    res = 1
    For i = 1 To t.Count
        If t(i).get_ID() = id Then
            res = i
        End If
    Next i
    t.Remove res
End Sub

'returns the index of the task in the array with its id as a parameter
Function get_task_index_by_id(id As Integer, t As Collection) As Integer

    Dim i As Integer, res As Integer
    res = 1

    For i = 1 To t.Count
        If t(i).get_ID = id Then
            res = i
        End If
    Next i

    get_task_index_by_id = res

End Function

Function task_in_tab_by_id(id As Integer, t As Collection) As Boolean

    Dim i As Integer, res As Boolean

    res = False

    For i = 1 To t.Count

        If t(i).get_ID = id Then
            res = True
        End If

    Next i

    task_in_tab_by_id = res

End Function


'insert element to a tab at custom indice
'The idea is to be able to insert the task in the place of the index k used previously, as it becomes the next element considered
Sub insertion_by_indice(element As Tache, t As Collection, i As Integer)

    Dim j As Integer, temp As Collection

    Set temp = New Collection

    For j = 1 To i - 1
        temp.Add t(j)
    Next j

    temp.Add element

    For j = i To t.Count
        temp.Add t(j)
    Next j

    Set t = temp

End Sub


'returns clue end tin
Function last_task_indice(t As Collection) As Integer

    Dim value As Integer, i As Integer

    value = CInt(t(1).get_fin)

    For i = 1 To t.Count
        If CInt(t(i).get_fin) > value Then
            value = i
        End If
    Next i

    last_task_indice = value

End Function


'the calendar column that is selected
Public Function colonne_date_actuelle()
    Dim d As String, colonne_date As Integer
    Dim s As Worksheet
    Set s = ThisWorkbook.Worksheets("GANTT")
    'Use strings because I have a problem using dates.

    colonne_date = 6
    d = s.Cells(4, colonne_date + 1)

    While d <> s.Cells(1, 16)
        colonne_date = colonne_date + 4
        d = s.Cells(4, colonne_date)
        If colonne_date > 5000 Then
            MsgBox " Merci de v�rifier la date du jour saisie (notamment la valeur en case P1)."
            colonne_date_actuelle = 40000
        End If
    Wend
    colonne_date_actuelle = colonne_date
    'The -4 makes it possible to compensate for the first 4 columns that are not part of the calendar.

End Function


Public Function dans_chaine_critique(id)
    'I get the box I fill it and I see if my ID is in it
    Dim splito() As String, check As Boolean, i As Integer
    Dim k As Worksheet
    Set k = ThisWorkbook.Worksheets("LOGS")
    check = False
    id = CStr(id)

    splito = Split(k.Cells(15, 15), ",")
    For i = LBound(splito) To UBound(splito)
        If id = splito(i) Then
            check = True
        End If
    Next i
    dans_chaine_critique = check
End Function
Public Function dans_quel_chaine(id)

    Dim splito() As String, check As Integer, i As Integer, j As Integer, ext As Boolean
    Dim k As Worksheet
    Set k = ThisWorkbook.Worksheets("LOGS")
    j = 15
    id = CStr(id)
    ext = False

    While k.Cells(j, 15) <> "" And ext = False
        splito = Split(k.Cells(j, 15), ",")
        For i = LBound(splito) To UBound(splito)
            If id = splito(i) Then
                check = j - 15 '15 for a to be equal to 0 for the critical chain, 1 for the first secondary chain
                ext = True
            End If
        Next i
        j = j + 1
    Wend
    If ext = False Then ' in no chain
        check = -1
    End If
    dans_quel_chaine = check

End Function

Public Function trouver_ligne_indice(indice)
    Dim s As Worksheet, i As Integer
    Set s = ThisWorkbook.Worksheets("GANTT")
    i = 6
    While s.Cells(i, 1) <> indice
        i = i + 2
    Wend
    trouver_ligne_indice = i

End Function
