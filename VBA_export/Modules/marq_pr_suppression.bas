Attribute VB_Name = "marq_pr_suppression"
'marking for suppression
Sub remove_chains_first_task(sc As Collection)
    Dim i As Integer, j As Integer
    For i = 1 To sc.Count
        If sc(i).Count > 0 Then
            Dim res As Boolean: res = False
            sc(i).Remove 1
            Dim preds() As String, k As Integer
            preds() = Split(sc(i)(1).get_preds, ",") 'we recover the preds of the first element

            For j = 2 To sc(i).Count
                If res = False Then
                    For k = 0 To UBound(preds)
                        If sc(i)(j).get_ID = CInt(preds(k)) Then
                            res = True
                        End If
                    Next k
                    If res = False Then
                        sc(i).Remove j
                    End If
                End If
            Next j
        End If
    Next i
End Sub

Sub update_preds(id As Integer)
    MsgBox id
    Dim i As Integer, preds() As String, k As Integer
    For i = 1 To taches.Count
        preds = Split(taches(i).get_preds, ",")
        Dim countr As String: countr = "" 'remember how many characters of the string are task ids
        For k = 0 To UBound(preds)
            countr = countr + preds(k)

            If id = CInt(preds(k)) Then

                If k = UBound(preds) Then
                    If UBound(preds) < 2 Then
                        taches(i).set_preds (vbNullString)
                    End If
                    taches(i).set_preds (Left(taches(i).get_preds, CInt(Len(taches(i).get_preds)) - CInt(Len(CStr(id)) - 1)))
                Else

                    Dim right_part As String, left_part As String
                    left_part = Left(taches(i).get_preds, Len(countr) - Len(preds(k)) + k)

                    right_part = Right(taches(i).get_preds, Len(taches(i).get_preds) - Len(left_part) - Len(preds(k)) - 1)
                    taches(i).set_preds (left_part + right_part)
                End If
            End If
        Next k
    Next i

End Sub


Public Function max_fin_preds_reel(preds() As String, t As Collection) As Integer
    't is the table of unsorted tasks.
    ' We look at the estimated end date to have an estimated start.
    ' If the predecessor is finished, the estimated date is the date of completion.
    Dim i As Integer, j As Integer, max As Integer, s As Worksheet, k As Worksheet, sauv As Integer
    Set s = ThisWorkbook.Worksheets("GANTT")
    Set k = ThisWorkbook.Worksheets("LOGS")
    max = 0

    For i = LBound(preds) To UBound(preds)
        For j = 1 To t.Count
            If preds(i) = t(j).get_ID Then
                If k.Cells(25 + t(j).get_ID, 5) > max Then

                    max = k.Cells(25 + t(j).get_ID, 5)
                    sauv = j
                End If
            End If
        Next j
    Next i
    If t(sauv).get_preds() <> "" Then
        max = max + 1
    End If
    max_fin_preds_reel = max

End Function
