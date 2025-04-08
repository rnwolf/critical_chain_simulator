VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Call retrieve_margins
    Dim i As Integer: i = GANTT_vertical_margin
    While Cells(i, 3).Interior.Color = vbWhite
        i = i + 2
    Wend
    
    Dim r As Range
    Set r = Range(Cells(6, 3), Cells(i - 1, 3))
    If Not Intersect(Target, r) Is Nothing Then
        With Worksheets("DASHBOARD")
            .Cells(Target.Row, Target.Column) = Target.value
        End With
    End If
    
    If Not Intersect(Target, Range(Cells(2, 3), Cells(2, 3))) Is Nothing Then
        With Worksheets("TÂCHES")
            If .Cells(2, 1).value <> Target.value Then
                .Cells(2, 1) = Target.value
            End If
        End With
        With Worksheets("DASHBOARD")
            If .Cells(Target.Row, Target.Column).value <> Target.value Then
                .Cells(Target.Row, Target.Column) = Target.value
            End If
        End With
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Call retrieve_margins
    If Selection.Count = 4 Then
        
        Dim left_limit As Integer: left_limit = GANTT_horizontal_margin + 4
        Dim i As Integer: i = 1
        Dim s As Worksheet
        Set s = ThisWorkbook.Worksheets("LOGS_FV_CHART")
        While s.Cells(16 + i, 8).value <> ""
            i = i + 1
        Wend
        If s.Cells(15 + i, 8).value <> "" Then
            'pos = (d - GANTT_horizontal_margin) * 2 + 2
            left_limit = (s.Cells(15 + i, 8).value + 2) / 2 + GANTT_horizontal_margin - 1
            
        End If
        'MsgBox left_limit
        
        Dim r As Range
        Dim orange_r As Range
        Dim grey_r As Range
        If left_limit = 0 Then
            MsgBox "Une erreur est survenue. Worksheet_SelectionChange"
            Exit Sub
        End If
        Set r = Range(Cells(GANTT_vertical_margin - 2, left_limit), Cells(GANTT_vertical_margin - 2, ThisWorkbook.Worksheets("LOGS").Cells(2, 1).value))
        Set grey_r = Range(Cells(GANTT_vertical_margin - 2, GANTT_horizontal_margin), Cells(GANTT_vertical_margin - 2, left_limit - 4))
        Set orange_r = Range(Cells(GANTT_vertical_margin - 2, GANTT_horizontal_margin), Cells(GANTT_vertical_margin - 2, ThisWorkbook.Worksheets("LOGS").Cells(2, 1).value))
        If Not Intersect(Target, r) Is Nothing Then 'on a cliqué sur une des cases du calendrier
            orange_r.Interior.Color = RGB(255, 242, 204)
            grey_r.Interior.Color = RGB(200, 200, 200)
            Target.Interior.Color = vbWhite
            ThisWorkbook.Worksheets("GANTT").Range("P1").value = Target.value 'updating current date cell
        End If
    End If
End Sub


