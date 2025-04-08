VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil4"
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
        With Worksheets("GANTT")
            .Cells(Target.Row, Target.Column) = Target.value
        End With
    End If
    
    If Not Intersect(Target, Range(Cells(2, 3), Cells(2, 3))) Is Nothing Then
        With Worksheets("TÂCHES")
            If .Cells(2, 1).value <> Target.value Then
                .Cells(2, 1) = Target.value
            End If
        End With
        With Worksheets("GANTT")
            If .Cells(Target.Row, Target.Column).value <> Target.value Then
                .Cells(Target.Row, Target.Column) = Target.value
            End If
        End With
    End If
End Sub




